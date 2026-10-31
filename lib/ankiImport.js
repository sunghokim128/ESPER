import { openZip } from './parseSheet.js';
import { imageTag } from './imageText.js';

const FIELD_SEPARATOR = '\x1f';
const TOKEN_START = '\uE000';
const TOKEN_END = '\uE001';
const IMG_PATTERN = /<img\b[^>]*?\bsrc\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))[^>]*>/gi;
const SQL_SCRIPT_URL = new URL('../vendor/sql.js/sql-asm.js', import.meta.url).href;

export class AnkiImportError extends Error {
  constructor(code) {
    super(code);
    this.code = code;
  }
}

let sqlPromise = null;

function loadSql() {
  if (!sqlPromise) {
    sqlPromise = (async () => {
      if (!globalThis.initSqlJs) {
        await new Promise((resolve, reject) => {
          const script = document.createElement('script');
          script.src = SQL_SCRIPT_URL;
          script.onload = resolve;
          script.onerror = () => reject(new Error('sql.js load failed'));
          document.head.append(script);
        });
      }
      return globalThis.initSqlJs();
    })().catch((error) => {
      sqlPromise = null;
      throw error;
    });
  }
  return sqlPromise;
}

function readModels(db) {
  const models = {};
  try {
    const json = JSON.parse(db.exec('select models from col')[0].values[0][0]);
    for (const id of Object.keys(json)) {
      const fields = [...(json[id].flds || [])].sort((a, b) => a.ord - b.ord).map((field) => String(field.name));
      models[id] = { name: String(json[id].name || ''), fields };
    }
    if (Object.keys(models).length) return models;
  } catch {
    /* 새 스키마라면 아래 테이블에서 읽습니다. */
  }
  try {
    db.exec('select id, name from notetypes')[0].values.forEach(([id, name]) => {
      models[String(id)] = { name: String(name), fields: [] };
    });
    db.exec('select ntid, name from fields order by ntid, ord')[0].values.forEach(([id, name]) => {
      models[String(id)]?.fields.push(String(name));
    });
  } catch {
    /* 노트 타입 정보를 못 읽어도 필드 번호로 계속 진행합니다. */
  }
  return models;
}

function decodeEntities(text) {
  const named = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };
  const decoded = text.replace(/&(#x[0-9a-f]+|#\d+|[a-z][a-z0-9]*);/gi, (match, body) => {
    if (body[0] === '#') {
      const code = body[1].toLowerCase() === 'x' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
      return Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : match;
    }
    return named[body.toLowerCase()] ?? match;
  });
  if (typeof DOMParser === 'undefined' || !/&[a-z]/i.test(decoded)) return decoded;
  return new DOMParser().parseFromString(`<!doctype html><body>${decoded.replace(/</g, '&lt;')}`, 'text/html').body.textContent;
}

// 파일 읽기 이후의 작업은 이 객체에 모아 둡니다. media 는 숫자 파일명 → 원본 파일명 JSON 입니다.
export async function loadAnkiDeck(file) {
  let zip;
  try {
    zip = await openZip(await file.arrayBuffer());
  } catch {
    throw new AnkiImportError('bad-file');
  }
  const names = new Set(zip.entries.map((entry) => entry.name));
  const collectionName = names.has('collection.anki21')
    ? 'collection.anki21'
    : names.has('collection.anki21b')
      ? null
      : names.has('collection.anki2')
        ? 'collection.anki2'
        : null;
  if (!collectionName) {
    throw new AnkiImportError(names.has('collection.anki21b') ? 'new-format' : 'no-collection');
  }

  const bytes = await zip.readBytes(collectionName);
  let SQL;
  try {
    SQL = await loadSql();
  } catch {
    throw new AnkiImportError('sql-unavailable');
  }
  const notes = [];
  let models;
  let db;
  try {
    db = new SQL.Database(bytes);
    models = readModels(db);
    const statement = db.prepare('select mid, flds from notes order by id');
    while (statement.step()) {
      const [mid, flds] = statement.get();
      notes.push({ mid: String(mid), fields: String(flds).split(FIELD_SEPARATOR) });
    }
    statement.free();
  } catch {
    throw new AnkiImportError('bad-database');
  } finally {
    db?.close();
  }
  if (!notes.length) throw new AnkiImportError('no-notes');

  const media = new Map();
  try {
    const json = JSON.parse((await zip.read('media')) || '{}');
    for (const entryName of Object.keys(json)) media.set(String(json[entryName]), entryName);
  } catch {
    /* media 목록이 없거나 깨져 있으면 이미지 없이 진행합니다. */
  }

  const counts = {};
  notes.forEach((note) => {
    counts[note.mid] = (counts[note.mid] || 0) + 1;
  });
  const noteTypes = Object.keys(counts)
    .sort((a, b) => counts[b] - counts[a])
    .map((id) => {
      const fieldCount = Math.max(...notes.filter((note) => note.mid === id).map((note) => note.fields.length));
      const known = models[id]?.fields || [];
      const fields = Array.from({ length: fieldCount }, (_, index) => known[index] || `Field ${index + 1}`);
      return { id, name: models[id]?.name || 'Unknown', fields, noteCount: counts[id] };
    });

  const lowerMedia = new Map([...media.keys()].map((name) => [name.toLowerCase(), name]));
  const findMedia = (source) => {
    const candidates = [source, decodeEntities(source)];
    for (const value of [...candidates]) {
      try {
        candidates.push(decodeURIComponent(value));
      } catch {
        /* 인코딩이 아닌 파일명은 그대로 둡니다. */
      }
    }
    for (const value of [...candidates]) {
      candidates.push(value.split('/').pop(), value.normalize('NFC'), value.normalize('NFD'));
    }
    for (const value of candidates) {
      if (media.has(value)) return value;
      if (lowerMedia.has(value.toLowerCase())) return lowerMedia.get(value.toLowerCase());
    }
    return null;
  };

  const readMedia = async (name) => {
    const entryName = media.get(name);
    const data = entryName ? await zip.readBytes(entryName) : null;
    if (!data) throw new Error('missing media');
    return new Blob([data], { type: mimeTypeFor(name) });
  };

  return { noteTypes, notes, findMedia, readMedia };
}

function mimeTypeFor(name) {
  const extension = (name.split('.').pop() || '').toLowerCase();
  return (
    { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', webp: 'image/webp', svg: 'image/svg+xml', bmp: 'image/bmp', avif: 'image/avif' }[extension] || ''
  );
}

// 필드마다 연결된 이미지가 몇 개인지 세어서 배치할 때 참고할 수 있게 합니다.
export function countFieldImages(deck, noteTypeId) {
  const noteType = deck.noteTypes.find((type) => type.id === noteTypeId);
  const stats = (noteType?.fields || []).map(() => ({ found: 0, missing: 0 }));
  for (const note of deck.notes) {
    if (note.mid !== noteTypeId) continue;
    note.fields.forEach((field, index) => {
      if (!stats[index]) return;
      for (const match of field.matchAll(IMG_PATTERN)) {
        if (deck.findMedia(match[1] ?? match[2] ?? match[3] ?? '')) stats[index].found += 1;
        else stats[index].missing += 1;
      }
    });
  }
  return stats;
}

// 이미지는 순서를 보존하기 위해 임시 토큰으로 바꿔 두었다가, 저장 후 [img:ID] 로 되돌립니다.
function fieldToText(deck, html) {
  const marked = String(html ?? '')
    .replace(/<(style|script)\b[\s\S]*?<\/\1>/gi, '')
    .replace(IMG_PATTERN, (_, a, b, c) => {
      const name = deck.findMedia(a ?? b ?? c ?? '');
      return name ? `${TOKEN_START}${name}${TOKEN_END}` : '';
    })
    .replace(/\[sound:[^\]]*\]/g, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(div|p|li|tr|h[1-6])>/gi, '\n')
    .replace(/<[^>]+>/g, '');
  return decodeEntities(marked)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

const TOKEN_PATTERN = new RegExp(`${TOKEN_START}([^${TOKEN_END}]*)${TOKEN_END}`, 'g');
const SEPARATORS = { newline: '\n', slash: ' / ', space: ' ' };

function yieldToBrowser() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

/**
 * mapping: { key: number[], value: number[], hint: number[] } (필드 번호)
 * saveImage(blob) => Promise<id>
 */
export async function buildAnkiRows(deck, noteTypeId, mapping, { separator = 'newline', saveImage, onProgress } = {}) {
  const joiner = SEPARATORS[separator] ?? '\n';
  const join = (note, indexes) =>
    indexes
      .map((index) => fieldToText(deck, note.fields[index] ?? ''))
      .filter(Boolean)
      .join(joiner);

  const drafts = [];
  const fileNames = new Set();
  let skippedNotes = 0;
  let scannedNotes = 0;
  for (const note of deck.notes) {
    scannedNotes += 1;
    if (scannedNotes % 200 === 0) await yieldToBrowser();
    if (note.mid !== noteTypeId) continue;
    const draft = [join(note, mapping.key), join(note, mapping.value), join(note, mapping.hint)];
    if (!draft[0] || !draft[1]) {
      skippedNotes += 1;
      continue;
    }
    draft.forEach((text) => text.replace(TOKEN_PATTERN, (_, name) => fileNames.add(name)));
    drafts.push(draft);
  }

  const imageIds = new Map();
  let processed = 0;
  let failedImages = 0;
  for (const name of fileNames) {
    try {
      imageIds.set(name, await saveImage(await deck.readMedia(name)));
    } catch {
      imageIds.set(name, null);
      failedImages += 1;
    }
    processed += 1;
    onProgress?.(processed, fileNames.size);
    if (processed % 5 === 0) await yieldToBrowser();
  }

  const toTags = (text) =>
    text
      .replace(TOKEN_PATTERN, (_, name) => (imageIds.get(name) ? imageTag(imageIds.get(name)) : ''))
      .replace(/[ \t]+\n/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
  const rows = [];
  for (const draft of drafts) {
    const [key, value, hint] = draft.map(toTags);
    if (!key || !value) {
      skippedNotes += 1;
      continue;
    }
    rows.push([key, value, hint]);
  }
  return { rows, skippedNotes, failedImages, imageCount: fileNames.size - failedImages };
}

export function previewAnkiRows(deck, noteTypeId, mapping, separator = 'newline', limit = 5) {
  const joiner = SEPARATORS[separator] ?? '\n';
  const show = (note, indexes) =>
    indexes
      .map((index) => fieldToText(deck, note.fields[index] ?? ''))
      .filter(Boolean)
      .join(joiner)
      .replace(TOKEN_PATTERN, '🖼');
  const rows = [];
  let total = 0;
  for (const note of deck.notes) {
    if (note.mid !== noteTypeId) continue;
    const row = [show(note, mapping.key), show(note, mapping.value), show(note, mapping.hint)];
    if (!row[0] || !row[1]) continue;
    total += 1;
    if (rows.length < limit) rows.push(row);
  }
  return { rows, total };
}

function previewParts(text) {
  const parts = [];
  let last = 0;
  for (const match of String(text).matchAll(new RegExp(TOKEN_PATTERN.source, 'g'))) {
    if (match.index > last) parts.push({ type: 'text', value: text.slice(last, match.index) });
    parts.push({ type: 'image', name: match[1] });
    last = match.index + match[0].length;
  }
  if (last < text.length) parts.push({ type: 'text', value: text.slice(last) });
  return parts;
}

export function previewAnkiCards(deck, noteTypeId, mapping, separator = 'newline', selectedIndex = 0) {
  const joiner = SEPARATORS[separator] ?? '\n';
  const show = (note, indexes) =>
    indexes
      .map((index) => fieldToText(deck, note.fields[index] ?? ''))
      .filter(Boolean)
      .join(joiner);

  let total = 0;
  let card = null;
  for (const note of deck.notes) {
    if (note.mid !== noteTypeId) continue;
    const fields = {
      key: show(note, mapping.key),
      value: show(note, mapping.value),
      hint: show(note, mapping.hint),
    };
    if (!fields.key || !fields.value) continue;
    if (total === selectedIndex) {
      card = Object.fromEntries(Object.entries(fields).map(([field, text]) => [field, previewParts(text)]));
    }
    total += 1;
  }
  return { card, total };
}

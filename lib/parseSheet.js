const utf8 = new TextDecoder('utf-8');

function decodeEntities(value) {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)))
    .replace(/&amp;/g, '&');
}

function cellText(value) {
  return decodeEntities(value)
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => line.replace(/[^\S\n]+/g, ' ').trim())
    .join('\n')
    .trim();
}

function rawCellText(value) {
  return value.replace(/\r\n?/g, '\n');
}

function richTextCell(value) {
  const runs = [...value.matchAll(/<r\b[^>]*>([\s\S]*?)<\/r>/g)].map((match) => {
    const body = match[1];
    const properties = (body.match(/<rPr\b[^>]*>([\s\S]*?)<\/rPr>/) || [])[1] || '';
    const boldTag = properties.match(/<b\b([^>]*)\/?>/);
    const bold = Boolean(boldTag) && !/\bval=["']?0["']?/.test(boldTag[1]);
    const texts = [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((text) => rawCellText(text[1]));
    return { text: texts.join(''), bold };
  });

  if (!runs.length) {
    return [...value.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)]
      .map((match) => rawCellText(match[1]))
      .join('');
  }

  return runs.reduce((result, run, index) => {
    const previous = runs[index - 1];
    const previousChar = previous?.text.slice(-1) || '';
    const firstChar = run.text[0] || '';
    const startsWord = /[\p{L}\p{N}]/u.test(firstChar);
    const needsSpace =
      previous &&
      previous.bold !== run.bold &&
      !/\s/.test(previousChar) &&
      !/\s/.test(firstChar) &&
      ((/[\p{L}\p{N}]/u.test(previousChar) && startsWord) ||
        (previous.bold && /[,;:.!?]$/.test(previousChar) && startsWord));
    return `${result}${needsSpace ? ' ' : ''}${run.text}`;
  }, '');
}

function hintCellText(value) {
  return decodeEntities(value).replace(/\r\n?/g, '\n').trim();
}

function decodeText(buffer) {
  let text = utf8.decode(new Uint8Array(buffer));
  if ((text.match(/\uFFFD/g) || []).length > 2) {
    try {
      text = new TextDecoder('euc-kr').decode(new Uint8Array(buffer));
    } catch {
      /* keep utf-8 result */
    }
  }
  return text.replace(/^\uFEFF/, '');
}

function colToIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export async function openZip(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);

  let eocd = -1;
  const stop = Math.max(0, bytes.length - 66000);
  for (let i = bytes.length - 22; i >= stop; i -= 1) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('엑셀 파일을 읽지 못했어요. .xlsx 형식인지 확인해 주세요.');

  const total = view.getUint16(eocd + 10, true);
  let pointer = view.getUint32(eocd + 16, true);
  const entries = [];

  for (let i = 0; i < total; i += 1) {
    if (view.getUint32(pointer, true) !== 0x02014b50) break;
    const method = view.getUint16(pointer + 10, true);
    const compSize = view.getUint32(pointer + 20, true);
    const nameLen = view.getUint16(pointer + 28, true);
    const extraLen = view.getUint16(pointer + 30, true);
    const commentLen = view.getUint16(pointer + 32, true);
    const local = view.getUint32(pointer + 42, true);
    const name = utf8.decode(bytes.subarray(pointer + 46, pointer + 46 + nameLen));
    entries.push({ name, method, compSize, local });
    pointer += 46 + nameLen + extraLen + commentLen;
  }

  const readBytes = async (name) => {
    const entry = entries.find((e) => e.name === name);
    if (!entry) return null;
    const nameLen = view.getUint16(entry.local + 26, true);
    const extraLen = view.getUint16(entry.local + 28, true);
    const start = entry.local + 30 + nameLen + extraLen;
    const data = bytes.subarray(start, start + entry.compSize);
    if (entry.method === 0) return data.slice();
    if (entry.method !== 8) throw new Error('지원하지 않는 압축 형식이에요.');
    if (typeof DecompressionStream === 'undefined') throw new Error('이 브라우저에서는 엑셀 파일을 읽을 수 없어요.');
    const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  };

  const read = async (name) => {
    const data = await readBytes(name);
    return data ? utf8.decode(data) : null;
  };

  return { entries, read, readBytes };
}

function parseSharedStrings(xml) {
  const strings = [];
  const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
  let si;
  while ((si = siRe.exec(xml))) {
    strings.push(richTextCell(si[1]));
  }
  return strings;
}

function parseWorksheet(xml, shared) {
  const rows = [];
  const rowRe = /<row\b([^>]*?)(?:>([\s\S]*?)<\/row>|\/>)/g;
  let row;
  while ((row = rowRe.exec(xml))) {
    const rowXml = row[2] || '';
    const cells = {};
    const cRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let c;
    while ((c = cRe.exec(rowXml))) {
      const attrs = c[1] || '';
      const body = c[2] || '';
      const ref = attrs.match(/r="([A-Z]+)\d+"/);
      const type = (attrs.match(/t="([^"]+)"/) || [])[1];
      const col = ref ? colToIndex(ref[1]) : Object.keys(cells).length;
      if (col > 3) continue;
      let value = '';
      if (type === 's') {
        const v = (body.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
        value = v != null ? shared[parseInt(v, 10)] ?? '' : '';
      } else if (type === 'inlineStr') {
        value = richTextCell(body);
      } else {
        const v = (body.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
        value = v != null ? cellText(v) : '';
      }
      if (value) cells[col] = value;
    }
    const values = [cells[0] ?? '', cells[1] ?? '', cells[2] ?? '', cells[3] ?? ''];
    values.sourceRow = Number((row[1].match(/\br="(\d+)"/) || [])[1]) || rows.length + 1;
    rows.push(values);
  }
  return rows;
}

async function sheetPathFrom(zip) {
  const workbook = await zip.read('xl/workbook.xml');
  const rels = await zip.read('xl/_rels/workbook.xml.rels');
  if (workbook && rels) {
    const sheet = workbook.match(/<sheet\b[^>]*r:id="([^"]+)"[^>]*\/?>/);
    if (sheet) {
      const rel = rels.match(new RegExp(`<Relationship\\b[^>]*Id="${sheet[1]}"[^>]*>`));
      const target = rel && rel[0].match(/Target="([^"]+)"/);
      if (target) {
        const path = target[1].replace(/^\/?xl\//, '').replace(/^\//, '');
        const sheetPath = `xl/${path}`;
        if (zip.entries.some((entry) => entry.name === sheetPath)) return sheetPath;
      }
    }
  }
  return zip.entries
    .map((e) => e.name)
    .filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))
    .sort()[0];
}

function workbookTargetPath(target) {
  const parts = target.startsWith('/') ? [] : ['xl'];
  for (const part of target.replace(/^\/+/, '').split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') parts.pop();
    else parts.push(part);
  }
  return parts.join('/');
}

async function workbookSheetsFrom(zip) {
  const workbook = await zip.read('xl/workbook.xml');
  const rels = await zip.read('xl/_rels/workbook.xml.rels');
  if (!workbook || !rels) return new Map();
  const result = new Map();
  for (const match of workbook.matchAll(/<sheet\b([^>]*?)\/?>/g)) {
    const name = decodeEntities((match[1].match(/\bname="([^"]*)"/) || [])[1] || '');
    const id = (match[1].match(/\br:id="([^"]+)"/) || [])[1];
    if (!name || !id) continue;
    const relation = [...rels.matchAll(/<Relationship\b([^>]*?)\/?>/g)]
      .find((item) => (item[1].match(/\bId="([^"]+)"/) || [])[1] === id);
    const target = relation && (relation[1].match(/\bTarget="([^"]+)"/) || [])[1];
    if (!target) continue;
    const path = workbookTargetPath(target);
    if (zip.entries.some((entry) => entry.name === path)) result.set(name, path);
  }
  return result;
}

async function sheetXmlFrom(zip) {
  const path = await sheetPathFrom(zip);
  return path ? zip.read(path) : null;
}

function sniffDelimiter(text) {
  const line = (text.split(/\r?\n/)[0] || '');
  const tabs = (line.match(/\t/g) || []).length;
  const commas = (line.match(/,/g) || []).length;
  const semis = (line.match(/;/g) || []).length;
  if (tabs > 0 && tabs >= commas && tabs >= semis) return '\t';
  if (semis > commas) return ';';
  return ',';
}

function parseDelimited(text, delimiter) {
  const rows = [];
  let row = [];
  let cell = '';
  let quoted = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cell += '"';
          i += 1;
        } else quoted = false;
      } else cell += ch;
    } else if (ch === '"') {
      quoted = true;
    } else if (ch === delimiter) {
      row.push(cell);
      cell = '';
    } else if (ch === '\n') {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
    } else if (ch !== '\r') {
      cell += ch;
    }
  }
  row.push(cell);
  rows.push(row);
  return rows;
}

function toPairs(rows) {
  const columnOffset = studyColumnOffset(rows);
  const cleaned = rows
    .map((row, index) => ({
      sourceRow: row.sourceRow ?? index + 1,
      values: [
        cellText(String(row[columnOffset] ?? '')),
        cellText(String(row[columnOffset + 1] ?? '')),
        hintCellText(String(row[columnOffset + 2] ?? '')),
      ],
    }))
    .filter(({ sourceRow }) => sourceRow > 1)
    .filter(({ values: [key, value] }) => key && value);

  const pairs = [];
  for (const { sourceRow, values: [key, value, hint] } of cleaned) {
    if (!key || !value) continue;
    pairs.push({
      id: `k${pairs.length}`,
      sourceRow,
      key,
      value,
      hint,
    });
  }
  return pairs;
}

function normalizedHeader(value) {
  return cellText(String(value ?? '')).toLocaleLowerCase().replace(/[\s_-]+/g, '');
}

function studyColumnOffset(rows) {
  const header = rows.find((row) => row.sourceRow === 1) || rows[0] || [];
  const headerKey = normalizedHeader(header[1]);
  const headerValue = normalizedHeader(header[2]);
  const headerHint = normalizedHeader(header[3]);
  const legacyStatus = new Set([
    '',
    'ignored',
    'status',
    'understoodstatus',
    'understood',
    '이해상태',
    '이해완료',
    '학습상태',
    '사용안함',
    '사용하지않음',
  ]);
  return legacyStatus.has(normalizedHeader(header[0])) &&
    headerKey === 'key' &&
    headerValue === 'value' &&
    headerHint === 'hint'
    ? 1
    : 0;
}

function normalizeStudyGrid(rows) {
  const offset = studyColumnOffset(rows);
  return rows.map((row) => {
    const normalized = Array.from({ length: 3 }, (_, index) =>
      String(row[index + offset] ?? '')
    );
    if (row.sourceRow !== undefined) normalized.sourceRow = row.sourceRow;
    return normalized;
  });
}

export async function parseStudyFile(file) {
  const name = (file.name || '').toLowerCase();

  if (name.endsWith('.xlsx') || name.endsWith('.xlsm')) {
    const zip = await openZip(await file.arrayBuffer());
    const sharedXml = await zip.read('xl/sharedStrings.xml');
    const shared = sharedXml ? parseSharedStrings(sharedXml) : [];
    const sheetXml = await sheetXmlFrom(zip);
    if (!sheetXml) throw new Error('엑셀 파일에서 시트를 찾지 못했어요.');
    return toPairs(parseWorksheet(sheetXml, shared));
  }

  if (name.endsWith('.csv') || name.endsWith('.txt') || name.endsWith('.tsv')) {
    const text = decodeText(await file.arrayBuffer());
    const delimiter = sniffDelimiter(text);
    return toPairs(
      parseDelimited(text, delimiter).map((row, index) => {
        row.sourceRow = index + 1;
        return row;
      })
    );
  }

  if (name.endsWith('.xls') || name.endsWith('.xlsb')) {
    throw new Error('구버전 .xls 형식은 지원하지 않아요. .xlsx, .xlsm, .csv, .txt 또는 .tsv로 저장해 주세요.');
  }

  throw new Error('지원하는 파일 형식은 .xlsx, .xlsm, .csv, .txt, .tsv예요.');
}

export async function normalizeStudyWorkbook(file) {
  const name = (file.name || '').toLowerCase();
  let rows;
  if (name.endsWith('.xlsx') || name.endsWith('.xlsm')) {
    const zip = await openZip(await file.arrayBuffer());
    const sharedXml = await zip.read('xl/sharedStrings.xml');
    const shared = sharedXml ? parseSharedStrings(sharedXml) : [];
    const sheetXml = await sheetXmlFrom(zip);
    if (!sheetXml) throw new Error('엑셀 파일에서 시트를 찾지 못했어요.');
    rows = parseWorksheet(sheetXml, shared);
  } else if (name.endsWith('.csv') || name.endsWith('.txt') || name.endsWith('.tsv')) {
    const text = decodeText(await file.arrayBuffer());
    rows = parseDelimited(text, sniffDelimiter(text)).map((row, index) => {
      row.sourceRow = index + 1;
      return row;
    });
  } else {
    return { file, converted: false };
  }
  if (!studyColumnOffset(rows)) return { file, converted: false };

  const grid = await readSheetGrid(file);
  const workbook = await writeSheetGrid(file, grid);
  return {
    file: new File([workbook], file.name, {
      type: workbook.type,
      lastModified: file.lastModified,
    }),
    converted: true,
  };
}

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

async function compressBytes(bytes) {
  if (typeof CompressionStream === 'undefined') return { method: 0, bytes };
  let compression;
  try {
    compression = new CompressionStream('deflate-raw');
  } catch {
    return { method: 0, bytes };
  }
  const stream = new Blob([bytes]).stream().pipeThrough(compression);
  return { method: 8, bytes: new Uint8Array(await new Response(stream).arrayBuffer()) };
}

function writeZip(entries, mimeType) {
  if (entries.length > 0xffff) throw new Error('항목이 너무 많아 엑셀 파일로 만들 수 없어요.');
  const localParts = [];
  const centralParts = [];
  let localOffset = 0;

  for (const entry of entries) {
    const nameBytes = new TextEncoder().encode(entry.name);
    const checksum = crc32(entry.data);
    if (
      nameBytes.length > 0xffff ||
      entry.data.length > 0xffffffff ||
      entry.compressed.length > 0xffffffff ||
      localOffset > 0xffffffff
    ) {
      throw new Error('파일이 너무 커서 엑셀 형식으로 내보낼 수 없어요.');
    }
    const localHeader = new Uint8Array(30 + nameBytes.length);
    const localView = new DataView(localHeader.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, 0x0800, true);
    localView.setUint16(8, entry.method, true);
    localView.setUint16(10, 0, true);
    localView.setUint16(12, 33, true);
    localView.setUint32(14, checksum, true);
    localView.setUint32(18, entry.compressed.length, true);
    localView.setUint32(22, entry.data.length, true);
    localView.setUint16(26, nameBytes.length, true);
    localView.setUint16(28, 0, true);
    localHeader.set(nameBytes, 30);
    localParts.push(localHeader, entry.compressed);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const centralView = new DataView(centralHeader.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, 0x0800, true);
    centralView.setUint16(10, entry.method, true);
    centralView.setUint16(12, 0, true);
    centralView.setUint16(14, 33, true);
    centralView.setUint32(16, checksum, true);
    centralView.setUint32(20, entry.compressed.length, true);
    centralView.setUint32(24, entry.data.length, true);
    centralView.setUint16(28, nameBytes.length, true);
    centralView.setUint32(38, 0, true);
    centralView.setUint32(42, localOffset, true);
    centralHeader.set(nameBytes, 46);
    centralParts.push(centralHeader);

    localOffset += localHeader.length + entry.compressed.length;
  }

  const centralSize = centralParts.reduce((size, part) => size + part.length, 0);
  if (centralSize > 0xffffffff || localOffset > 0xffffffff) {
    throw new Error('파일이 너무 커서 엑셀 형식으로 내보낼 수 없어요.');
  }
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(8, entries.length, true);
  endView.setUint16(10, entries.length, true);
  endView.setUint32(12, centralSize, true);
  endView.setUint32(16, localOffset, true);

  return new Blob([...localParts, ...centralParts, end], {
    type: mimeType,
  });
}

export async function createZipBlob(files, mimeType) {
  const entries = [];
  for (const { name, data } of files) {
    const compressed = await compressBytes(data);
    entries.push({ name, data, method: compressed.method, compressed: compressed.bytes });
  }
  return writeZip(entries, mimeType);
}

function encodeDelimitedCell(value, delimiter) {
  const text = String(value ?? '');
  return text.includes(delimiter) || /["\r\n]/.test(text)
    ? `"${text.replace(/"/g, '""')}"`
    : text;
}

const STUDY_KEYS_SHEET = 'a';
const STUDY_SESSIONS_SHEET = 'b';
const XLSX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function recordCellXml(reference, value) {
  if (typeof value === 'number') return `<c r="${reference}"><v>${value}</v></c>`;
  return `<c r="${reference}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(String(value))}</t></is></c>`;
}

function recordWorksheetXml(rows) {
  const lastColumn = String.fromCharCode(64 + Math.max(...rows.map((row) => row.length)));
  const rowXml = rows.map((row, rowIndex) => {
    const cells = row.map((value, columnIndex) =>
      recordCellXml(`${String.fromCharCode(65 + columnIndex)}${rowIndex + 1}`, value)
    ).join('');
    return `<row r="${rowIndex + 1}">${cells}</row>`;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${lastColumn}${rows.length}"/><sheetData>${rowXml}</sheetData></worksheet>`;
}

async function createStudyRecordWorkbook(keyRows, sessionRows) {
  const files = [
    ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`],
    ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${STUDY_KEYS_SHEET}" sheetId="1" r:id="rId1"/><sheet name="${STUDY_SESSIONS_SHEET}" sheetId="2" r:id="rId2"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`],
    ['xl/worksheets/sheet1.xml', recordWorksheetXml(keyRows)],
    ['xl/worksheets/sheet2.xml', recordWorksheetXml(sessionRows)],
  ];
  const entries = [];
  for (const [name, content] of files) {
    const data = new TextEncoder().encode(content);
    const compressed = await compressBytes(data);
    entries.push({ name, data, method: compressed.method, compressed: compressed.bytes });
  }
  return writeZip(entries, XLSX_MIME_TYPE);
}

function rowsFromRecordSheet(xml, shared) {
  return parseWorksheet(xml, shared)
    .sort((left, right) => left.sourceRow - right.sourceRow)
    .map((row) => [...row]);
}

export async function readStudyRecordWorkbook(file) {
  const zip = await openZip(await file.arrayBuffer());
  const sheets = await workbookSheetsFrom(zip);
  const keysPath = sheets.get(STUDY_KEYS_SHEET);
  const sessionsPath = sheets.get(STUDY_SESSIONS_SHEET);
  if (!keysPath || !sessionsPath) {
    throw new Error('학습 기록 파일에 필요한 시트가 없어 내용을 읽을 수 없어요.');
  }
  const sharedXml = await zip.read('xl/sharedStrings.xml');
  const shared = sharedXml ? parseSharedStrings(sharedXml) : [];
  const keysXml = await zip.read(keysPath);
  const sessionsXml = await zip.read(sessionsPath);
  if (!keysXml || !sessionsXml) throw new Error('학습 기록 시트를 읽지 못했어요.');
  return {
    keys: rowsFromRecordSheet(keysXml, shared).slice(1)
      .filter((row) => String(row[0] ?? '').trim())
      .map((row) => ({ key: String(row[0]), count: Number(row[1]) || 0 })),
    sessions: rowsFromRecordSheet(sessionsXml, shared).slice(1)
      .filter((row) => String(row[0] ?? '').trim())
      .map((row) => ({
        date: String(row[0]),
        itemCount: Number(row[1]) || 0,
        duration: String(row[2] ?? ''),
      })),
  };
}

export async function exportStudyRecordWorkbook(file, studiedKeys, session) {
  let keyRows = [['key', '학습 횟수']];
  let sessionRows = [['날짜', '학습 항목 수', '소요 시간']];
  if (file) {
    const zip = await openZip(await file.arrayBuffer());
    const sheets = await workbookSheetsFrom(zip);
    const keysPath = sheets.get(STUDY_KEYS_SHEET);
    const sessionsPath = sheets.get(STUDY_SESSIONS_SHEET);
    if (!keysPath || !sessionsPath) {
      throw new Error('기존 학습 기록 파일에 필요한 시트가 없어 내용을 갱신할 수 없어요.');
    }
    const sharedXml = await zip.read('xl/sharedStrings.xml');
    const shared = sharedXml ? parseSharedStrings(sharedXml) : [];
    const existingKeysXml = await zip.read(keysPath);
    const existingSessionsXml = await zip.read(sessionsPath);
    if (!existingKeysXml || !existingSessionsXml) throw new Error('학습 기록 시트를 읽지 못했어요.');
    keyRows = rowsFromRecordSheet(existingKeysXml, shared);
    sessionRows = rowsFromRecordSheet(existingSessionsXml, shared);
    if (keyRows.length === 0) keyRows = [['key', '학습 횟수']];
    if (sessionRows.length === 0) sessionRows = [['날짜', '학습 항목 수', '소요 시간']];
  }

  const counts = new Map();
  for (const row of keyRows.slice(1)) {
    const key = String(row[0] ?? '');
    if (key) counts.set(key, Number(row[1]) || 0);
  }
  for (const key of studiedKeys) {
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  keyRows = [
    keyRows[0],
    ...[...counts].map(([key, count]) => [key, count]),
  ];
  sessionRows.push([session.date, session.itemCount, session.duration]);
  return createStudyRecordWorkbook(keyRows, sessionRows);
}

/* ---------------- 원본 표 편집 ---------------- */

const GRID_COLUMNS = 3;
const COLUMN_LETTERS = ['A', 'B', 'C'];

function normalizeGridRow(row) {
  return Array.from({ length: GRID_COLUMNS }, (_, index) => String(row[index] ?? ''));
}

function trimGrid(rows) {
  const grid = rows.map(normalizeGridRow);
  while (grid.length && grid[grid.length - 1].every((cell) => !cell.trim())) grid.pop();
  return grid;
}

export async function readSheetGrid(file) {
  const name = (file.name || '').toLowerCase();
  if (name.endsWith('.xlsx') || name.endsWith('.xlsm')) {
    const zip = await openZip(await file.arrayBuffer());
    const sharedXml = await zip.read('xl/sharedStrings.xml');
    const shared = sharedXml ? parseSharedStrings(sharedXml) : [];
    const sheetXml = await sheetXmlFrom(zip);
    if (!sheetXml) throw new Error('엑셀 파일에서 시트를 찾지 못했어요.');
    const grid = [];
    for (const row of parseWorksheet(sheetXml, shared)) {
      while (grid.length < row.sourceRow - 1) grid.push(normalizeGridRow([]));
      const values = row.map((cell) => decodeEntities(String(cell)));
      values.sourceRow = row.sourceRow;
      grid[row.sourceRow - 1] = values;
    }
    return trimGrid(normalizeStudyGrid(grid));
  }
  if (name.endsWith('.csv') || name.endsWith('.txt') || name.endsWith('.tsv')) {
    const text = decodeText(await file.arrayBuffer());
    return trimGrid(normalizeStudyGrid(parseDelimited(text, sniffDelimiter(text))));
  }
  throw new Error('편집할 수 있는 파일 형식은 .xlsx, .xlsm, .csv, .txt, .tsv예요.');
}

function escapeXml(value) {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function gridCellXml(letter, rowNumber, value, style) {
  if (!value) return '';
  const styleAttribute = style ? ` s="${style}"` : '';
  return `<c r="${letter}${rowNumber}"${styleAttribute} t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
}

function firstDataRowStyles(sheetData, columnOffset = 0) {
  const styles = {};
  const rows = [...sheetData.matchAll(/<row\b([^>]*?)(?:>([\s\S]*?)<\/row>|\/>)/g)];
  const dataRow = rows.find((match) => Number((match[1].match(/\br="(\d+)"/) || [])[1]) >= 2);
  for (const cell of (dataRow?.[2] || '').matchAll(/<c\b([^>]*?)(?:\/>|>)/g)) {
    const reference = (cell[1].match(/\br="([A-Z]+)\d+"/) || [])[1];
    const style = (cell[1].match(/\bs="(\d+)"/) || [])[1];
    if (!reference || !style) continue;
    const columnIndex = colToIndex(reference) - columnOffset;
    if (columnIndex >= 0 && columnIndex < COLUMN_LETTERS.length) {
      styles[COLUMN_LETTERS[columnIndex]] = style;
    }
  }
  return styles;
}

export async function writeSheetGrid(file, rows) {
  const grid = trimGrid(rows);
  const name = (file.name || '').toLowerCase();

  if (name.endsWith('.csv') || name.endsWith('.txt') || name.endsWith('.tsv')) {
    const text = decodeText(await file.arrayBuffer());
    const delimiter = sniffDelimiter(text);
    const output = grid
      .map((row) => row.map((cell) => encodeDelimitedCell(cell, delimiter)).join(delimiter))
      .join('\r\n');
    const type = name.endsWith('.tsv')
      ? 'text/tab-separated-values;charset=utf-8'
      : name.endsWith('.txt')
        ? 'text/plain;charset=utf-8'
        : 'text/csv;charset=utf-8';
    return new Blob([`\uFEFF${output}`], { type });
  }

  if (!name.endsWith('.xlsx') && !name.endsWith('.xlsm')) {
    throw new Error('편집할 수 있는 파일 형식은 .xlsx, .xlsm, .csv, .txt, .tsv예요.');
  }
  const zip = await openZip(await file.arrayBuffer());
  const sheetPath = await sheetPathFrom(zip);
  if (!sheetPath) throw new Error('엑셀 파일에서 시트를 찾지 못했어요.');
  const sheetXml = await zip.read(sheetPath);
  if (!sheetXml) throw new Error('엑셀 파일의 첫 번째 시트를 읽지 못했어요.');

  const sheetDataMatch = sheetXml.match(/<sheetData\b[^>]*?(?:\/>|>([\s\S]*?)<\/sheetData>)/);
  if (!sheetDataMatch) throw new Error('엑셀 파일의 시트 데이터를 찾지 못했어요.');
  const originalData = sheetDataMatch[1] || '';
  const sharedXml = await zip.read('xl/sharedStrings.xml');
  const originalSheetRows = parseWorksheet(sheetXml, sharedXml ? parseSharedStrings(sharedXml) : []);
  const legacyColumnOffset = studyColumnOffset(originalSheetRows);
  const styles = firstDataRowStyles(originalData, legacyColumnOffset);

  const originalGrid = normalizeStudyGrid(originalSheetRows);
  const originalHeader = originalGrid.find((row) => row.sourceRow === 1);
  const headerXml = (originalData.match(/<row\b[^>]*?\br="1"[^>]*?>[\s\S]*?<\/row>/) || [])[0];
  const headerHasExtraColumns = [...(headerXml?.matchAll(/<c\b([^>]*?)(?:\/>|>)/g) || [])]
    .some((cell) => {
      const reference = (cell[1].match(/\br="([A-Z]+)\d+"/) || [])[1];
      return reference && colToIndex(reference) > 2;
    });
  const headerUnchanged =
    originalHeader &&
    headerXml &&
    !headerHasExtraColumns &&
    normalizeGridRow(originalHeader).every((cell, index) => cell === grid[0]?.[index]);

  const rowsXml = grid.map((row, index) => {
    const rowNumber = index + 1;
    if (rowNumber === 1 && headerUnchanged) return headerXml;
    const cells = row
      .map((cell, column) => gridCellXml(COLUMN_LETTERS[column], rowNumber, cell, rowNumber > 1 ? styles[COLUMN_LETTERS[column]] : ''))
      .join('');
    return `<row r="${rowNumber}">${cells}</row>`;
  });
  const sheetDataXml = `<sheetData>${rowsXml.join('')}</sheetData>`;
  let updatedSheetXml = sheetXml.replace(sheetDataMatch[0], () => sheetDataXml);
  updatedSheetXml = updatedSheetXml.replace(
    /<dimension\b[^>]*\/>/,
    `<dimension ref="A1:C${Math.max(grid.length, 1)}"/>`
  );

  const entries = [];
  for (const entry of zip.entries) {
    const data =
      entry.name === sheetPath
        ? new TextEncoder().encode(updatedSheetXml)
        : await zip.readBytes(entry.name);
    if (!data) throw new Error(`엑셀 파일의 ${entry.name} 항목을 읽지 못했어요.`);
    const compressed = await compressBytes(data);
    entries.push({ name: entry.name, data, method: compressed.method, compressed: compressed.bytes });
  }
  const mimeType = name.endsWith('.xlsm')
    ? 'application/vnd.ms-excel.sheet.macroEnabled.12'
    : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  return writeZip(entries, mimeType);
}


/* ---------------- 새 범용 엑셀 파일 만들기 ---------------- */

export async function createStudyWorkbook(rows, header = ['key', 'value', 'hint']) {
  const grid = [header, ...rows].map(normalizeGridRow);
  const rowXml = grid
    .map((row, rowIndex) => {
      const cells = row
        .map((value, column) => gridCellXml(COLUMN_LETTERS[column], rowIndex + 1, value, ''))
        .join('');
      return `<row r="${rowIndex + 1}">${cells}</row>`;
    })
    .join('');
  const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C${grid.length}"/><cols><col min="1" max="1" width="24" customWidth="1"/><col min="2" max="2" width="48" customWidth="1"/><col min="3" max="3" width="60" customWidth="1"/></cols><sheetData>${rowXml}</sheetData></worksheet>`;
  const files = [
    ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`],
    ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`],
    ['xl/worksheets/sheet1.xml', sheetXml],
  ];
  const entries = [];
  for (const [name, content] of files) {
    const data = new TextEncoder().encode(content);
    const compressed = await compressBytes(data);
    entries.push({ name, data, method: compressed.method, compressed: compressed.bytes });
  }
  return writeZip(entries, XLSX_MIME_TYPE);
}

import '../vendor/ts-fsrs/index.umd.js';
import { readFsrsDataFile, writeFsrsDataFile } from './projectStore.js';

const TS_FSRS_VERSION = '6.0.0';
const FSRS_MODEL = 'FSRS-7';
const SCHEMA_VERSION = 1;
const { DefaultScheduler, Rating, State } = globalThis.FSRS;
export const DEFAULT_DESIRED_RETENTION = 0.9;
export const MIN_DESIRED_RETENTION = 0.7;
export const MAX_DESIRED_RETENTION = 0.99;
const schedulerPromises = new Map();

function normalizeRetention(value) {
  const retention = Number(value);
  if (!Number.isFinite(retention)) return DEFAULT_DESIRED_RETENTION;
  return Math.min(MAX_DESIRED_RETENTION, Math.max(MIN_DESIRED_RETENTION, Math.round(retention * 100) / 100));
}

const DEFAULT_MAXIMUM_INTERVAL = 36500;

export function examMaximumInterval(examDate, now = new Date()) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(examDate || '')) return DEFAULT_MAXIMUM_INTERVAL;
  const examStart = new Date(`${examDate}T00:00:00`);
  const days = Math.floor((examStart.getTime() - now.getTime()) / 86400000);
  return Math.min(DEFAULT_MAXIMUM_INTERVAL, Math.max(1, days));
}

function getScheduler(desiredRetention = DEFAULT_DESIRED_RETENTION, maximumInterval = DEFAULT_MAXIMUM_INTERVAL) {
  const retention = normalizeRetention(desiredRetention);
  const key = `${retention}:${maximumInterval}`;
  if (!schedulerPromises.has(key)) {
    schedulerPromises.set(key, DefaultScheduler({ version: FSRS_MODEL, desiredRetention: retention, maximumInterval }));
  }
  return schedulerPromises.get(key);
}
const schedulerPromise = getScheduler();

function isObject(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function validDate(value) {
  return typeof value === 'string' && Number.isFinite(Date.parse(value));
}

function assertCard(record) {
  const card = record?.card;
  const validState = Object.values(State).includes(card?.state);
  const validScheduleStatus = ['new', 'learning', 'review', 'relearning'].includes(card?.scheduleStatus);
  const numericFields = ['stability', 'stabilityFast', 'difficulty', 'reps', 'lapses', 'scheduledDays', 'learningStep'];
  if (
    !isObject(record) ||
    typeof record.id !== 'string' ||
    record.id.length === 0 ||
    typeof record.fingerprint !== 'string' ||
    !/^[a-f0-9]{64}$/.test(record.fingerprint) ||
    (record.archivedAt !== null && !validDate(record.archivedAt)) ||
    !isObject(card) ||
    card.cardId !== record.id ||
    !validDate(card.dueAt) ||
    (card.lastReviewAt !== null && !validDate(card.lastReviewAt)) ||
    !validState ||
    !validScheduleStatus ||
    numericFields.some((field) => !Number.isFinite(card[field])) ||
    !Number.isInteger(card.reps) ||
    !Number.isInteger(card.lapses) ||
    !Number.isInteger(card.learningStep) ||
    !Array.isArray(record.reviewLogs) ||
    record.reviewLogs.some((log) =>
      !isObject(log) ||
      log.cardId !== record.id ||
      ![Rating.Again, Rating.Hard, Rating.Good, Rating.Easy].includes(log.rating) ||
      !validDate(log.reviewTime)
    )
  ) {
    throw new Error('FSRS 카드 데이터의 구조가 올바르지 않습니다.');
  }
}

function validateData(data, project) {
  if (
    !isObject(data) ||
    data.schemaVersion !== SCHEMA_VERSION ||
    data.projectId !== project.id ||
    data.workbookFileName !== project.fileName ||
    data.library?.name !== 'ts-fsrs' ||
    data.library?.version !== TS_FSRS_VERSION ||
    data.library?.model !== FSRS_MODEL ||
    !Array.isArray(data.cards)
  ) {
    throw new Error('FSRS 파일의 프로젝트 정보, 라이브러리 버전 또는 스키마가 맞지 않습니다. 기존 기록은 변경하지 않았습니다.');
  }
  const ids = new Set();
  for (const record of data.cards) {
    assertCard(record);
    if (ids.has(record.id)) throw new Error('FSRS 파일에 중복된 카드 ID가 있습니다.');
    ids.add(record.id);
  }
  return data;
}

function emptyData(project) {
  return {
    schemaVersion: SCHEMA_VERSION,
    projectId: project.id,
    workbookFileName: project.fileName,
    library: { name: 'ts-fsrs', version: TS_FSRS_VERSION, model: FSRS_MODEL },
    cards: [],
    ...(project.plan ? { plan: project.plan } : {}),
  };
}

async function itemFingerprint(pair) {
  const source = JSON.stringify([String(pair.key), String(pair.value), String(pair.hint)]);
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(source));
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('');
}

async function newRecord(fingerprint) {
  const id = crypto.randomUUID();
  const scheduler = await schedulerPromise;
  const card = scheduler.newCard({ now: new Date(), cardId: id });
  return {
    id,
    fingerprint,
    archivedAt: null,
    card: JSON.parse(JSON.stringify(card)),
    reviewLogs: [],
  };
}

export async function reconcileFsrsData(project, pairs, stored) {
  let data;
  if (stored === null) {
    data = emptyData(project);
  } else {
    try {
      data = validateData(JSON.parse(stored), project);
    } catch (error) {
      throw new Error(`FSRS 학습 기록을 읽지 못했습니다. 파일을 초기화하거나 덮어쓰지 않았습니다. ${error instanceof Error ? error.message : String(error)}`);
    }
    data = structuredClone(data);
  }

  const incoming = await Promise.all(pairs.map(async (pair) => ({
    pair,
    fingerprint: await itemFingerprint(pair),
  })));
  const activeByFingerprint = new Map();
  for (const record of data.cards) {
    if (record.archivedAt !== null) continue;
    const records = activeByFingerprint.get(record.fingerprint) || [];
    records.push(record);
    activeByFingerprint.set(record.fingerprint, records);
  }
  const incomingByFingerprint = new Map();
  for (const item of incoming) {
    const items = incomingByFingerprint.get(item.fingerprint) || [];
    items.push(item);
    incomingByFingerprint.set(item.fingerprint, items);
  }

  const now = new Date().toISOString();
  const synchronizedPairs = new Map();
  for (const [fingerprint, oldRecords] of activeByFingerprint) {
    const newItems = incomingByFingerprint.get(fingerprint) || [];
    if (oldRecords.length === 1 && newItems.length === 1) {
      synchronizedPairs.set(newItems[0].pair, oldRecords[0]);
    } else if (oldRecords.length > 1 && oldRecords.length === newItems.length) {
      oldRecords.sort((a, b) => a.id.localeCompare(b.id));
      newItems.forEach((item, index) => synchronizedPairs.set(item.pair, oldRecords[index]));
    } else if (oldRecords.length !== newItems.length || oldRecords.length > 1) {
      for (const record of oldRecords) record.archivedAt = now;
    }
  }

  for (const item of incoming) {
    if (synchronizedPairs.has(item.pair)) continue;
    const record = await newRecord(item.fingerprint);
    data.cards.push(record);
    synchronizedPairs.set(item.pair, record);
  }

  for (const item of incoming) {
    item.pair.id = synchronizedPairs.get(item.pair).id;
  }
  if (project.plan) data.plan = { examDate: project.plan.examDate, desiredRetention: project.plan.desiredRetention };
  return { data: validateData(data, project), pairs };
}

export async function prepareFsrsSync(project, pairs) {
  const stored = await readFsrsDataFile(project.name, project.fileName);
  return reconcileFsrsData(project, pairs, stored);
}

export async function saveFsrsData(project, data) {
  validateData(data, project);
  await writeFsrsDataFile(project.name, project.fileName, JSON.stringify(data, null, 2));
}

export async function synchronizeFsrsProject(project, pairs) {
  const result = await prepareFsrsSync(project, pairs);
  await saveFsrsData(project, result.data);
  return result;
}

export async function reviewFsrsCard(project, data, cardId, rating, now = new Date(), desiredRetention = DEFAULT_DESIRED_RETENTION, examDate = '') {
  const record = data.cards.find((item) => item.id === cardId && item.archivedAt === null);
  if (!record) throw new Error('복습할 FSRS 카드를 찾을 수 없습니다.');
  if (![Rating.Again, Rating.Hard, Rating.Good, Rating.Easy].includes(rating)) {
    throw new Error('지원하지 않는 FSRS 평가입니다.');
  }
  const scheduler = await getScheduler(desiredRetention, examMaximumInterval(examDate, now));
  const card = {
    ...record.card,
    dueAt: new Date(record.card.dueAt),
    lastReviewAt: record.card.lastReviewAt === null ? null : new Date(record.card.lastReviewAt),
  };
  const result = scheduler.review({ card, grade: rating, now });
  const nextData = structuredClone(data);
  const nextRecord = nextData.cards.find((item) => item.id === cardId);
  nextRecord.card = JSON.parse(JSON.stringify(result.card));
  nextRecord.reviewLogs.push(JSON.parse(JSON.stringify(result.revlog)));
  validateData(nextData, project);
  return { data: nextData, card: nextRecord.card, reviewLog: result.revlog };
}

export async function previewFsrsCard(card, now = new Date(), desiredRetention = DEFAULT_DESIRED_RETENTION, examDate = '') {
  const scheduler = await getScheduler(desiredRetention, examMaximumInterval(examDate, now));
  const fsrsCard = {
    ...card,
    dueAt: new Date(card.dueAt),
    lastReviewAt: card.lastReviewAt === null ? null : new Date(card.lastReviewAt),
  };
  return [...scheduler.preview({ card: fsrsCard, now })].map(({ grade, card: nextCard }) => ({
    grade,
    dueAt: nextCard.dueAt,
  }));
}

export async function getFsrsRetrievability(card, now = new Date()) {
  const scheduler = await schedulerPromise;
  const fsrsCard = {
    ...card,
    dueAt: new Date(card.dueAt),
    lastReviewAt: card.lastReviewAt === null ? null : new Date(card.lastReviewAt),
  };
  return scheduler.prepareReview(fsrsCard, now).retrievability;
}

export async function summarizeExamReadiness(cards, examDate, desiredRetention, now = new Date()) {
  const exam = new Date(examDate);
  const target = normalizeRetention(desiredRetention);
  const reviewed = cards.filter((record) => record.card.state !== State.New && record.card.scheduleStatus !== 'new');
  const newCount = cards.length - reviewed.length;
  const daysLeft = Math.ceil((exam.getTime() - now.getTime()) / 86400000);
  const values = await Promise.all(reviewed.map((record) => getFsrsRetrievability(record.card, exam)));
  const dueBeforeExam = reviewed.filter((record) => Date.parse(record.card.dueAt) <= exam.getTime()).length;
  return {
    daysLeft,
    newCount,
    reviewedCount: reviewed.length,
    dueBeforeExam,
    averageRetention: values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null,
    newPerDay: daysLeft > 0 ? Math.ceil(newCount / daysLeft) : newCount,
    target,
  };
}

export { normalizeRetention as normalizeDesiredRetention };

export async function rescheduleFsrsData(project, data, desiredRetention, examDate, now = new Date()) {
  const scheduler = await getScheduler(desiredRetention);
  const retention = normalizeRetention(desiredRetention);
  const examStart = /^\d{4}-\d{2}-\d{2}$/.test(examDate || '') ? new Date(`${examDate}T00:00:00`).getTime() : null;
  const nextData = structuredClone(data);
  let changed = 0;
  for (const record of nextData.cards) {
    const card = record.card;
    if (record.archivedAt !== null || card.state !== State.Review || card.lastReviewAt === null) continue;
    const lastReview = Date.parse(card.lastReviewAt);
    const intervalDays = Math.max(1, scheduler.model.nextInterval(card, retention));
    let due = lastReview + intervalDays * 86400000;
    if (examStart !== null && examStart > now.getTime()) due = Math.min(due, examStart);
    const dueAt = new Date(due).toISOString();
    if (dueAt !== card.dueAt) changed += 1;
    card.dueAt = dueAt;
    card.scheduledDays = Math.max(1, Math.round((due - lastReview) / 86400000));
  }
  validateData(nextData, project);
  return { data: nextData, changed };
}

export function fsrsReviewCategory(card, now = Date.now()) {
  if (card?.state === State.New || card?.scheduleStatus === 'new') return 'new';
  return Date.parse(card?.dueAt) <= now ? 'due' : 'scheduled';
}

export { Rating as FsrsRating, State as FsrsState };

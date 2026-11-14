import assert from 'node:assert/strict';
import test from 'node:test';
import {
  FsrsRating,
  FsrsState,
  fsrsReviewCategory,
  getFsrsRetrievability,
  previewFsrsCard,
  reconcileFsrsData,
  reviewFsrsCard,
  summarizeExamReadiness,
  examMaximumInterval,
  rescheduleFsrsData,
  updateFsrsCardFingerprint,
} from '../lib/fsrs.js';
import {
  createStudyWorkbook,
  openZip,
  normalizeStudyWorkbook,
  parseStudyFile,
  readSheetGrid,
  writeSheetGrid,
} from '../lib/parseSheet.js';
import { normalizeColumnWidths, resizeAdjacentColumns } from '../lib/columnSizing.js';
import { indexStudyPairs } from '../lib/studyPairs.js';
import { buildPriorityPages, sortByRecallProbability } from '../lib/sheetPages.js';

const project = { id: 'project-1', fileName: 'Example.xlsx' };
const pair = (key, value = `value for ${key}`, hint = '') => ({ key, value, hint });

test('FSRS-7 review state and all four schedule previews come from ts-fsrs', async () => {
  const initial = await reconcileFsrsData(project, [pair('alpha')], null);
  const cardId = initial.pairs[0].id;
  assert.equal(initial.data.cards[0].card.state, FsrsState.New);
  assert.equal(initial.data.cards[0].card.cardId, cardId);

  const previews = await previewFsrsCard(initial.data.cards[0].card, new Date());
  assert.deepEqual(previews.map((item) => item.grade), [
    FsrsRating.Again,
    FsrsRating.Hard,
    FsrsRating.Good,
    FsrsRating.Easy,
  ]);
  assert.ok(previews.every((item) => item.dueAt instanceof Date));

  const reviewed = await reviewFsrsCard(
    project,
    initial.data,
    cardId,
    FsrsRating.Good,
    new Date()
  );
  const record = reviewed.data.cards[0];
  assert.equal(record.id, cardId);
  assert.equal(record.card.reps, 1);
  assert.equal(record.reviewLogs.length, 1);
  assert.equal(record.reviewLogs[0].rating, FsrsRating.Good);
  assert.equal(record.card.state, FsrsState.Review);
});

test('FSRS project snapshots preserve the manual daily pace mode', async () => {
  const plannedProject = {
    ...project,
    plan: {
      examDate: '2026-12-03',
      desiredRetention: 0.9,
      dailyNewCards: 7,
      useRecommendedDailyPace: false,
      firstPassDeadline: '2026-11-27',
    },
  };
  const result = await reconcileFsrsData(plannedProject, [pair('alpha')], null);
  assert.deepEqual(result.data.plan, plannedProject.plan);
});

test('editing card text updates its fingerprint without resetting its review state', async () => {
  const initial = await reconcileFsrsData(project, [pair('alpha')], null);
  const id = initial.pairs[0].id;
  const reviewed = (await reviewFsrsCard(
    project,
    initial.data,
    id,
    FsrsRating.Good,
    new Date('2026-01-02T00:00:00.000Z')
  )).data;
  const originalRecord = reviewed.cards.find((record) => record.id === id);
  const editedPair = { ...initial.pairs[0], value: 'corrected answer' };
  const updated = await updateFsrsCardFingerprint(project, reviewed, editedPair);
  const updatedRecord = updated.data.cards.find((record) => record.id === id);

  assert.notEqual(updated.fingerprint, originalRecord.fingerprint);
  assert.deepEqual(updatedRecord.card, originalRecord.card);
  assert.deepEqual(updatedRecord.reviewLogs, originalRecord.reviewLogs);

  const synchronized = await reconcileFsrsData(project, [editedPair], JSON.stringify(updated.data));
  assert.equal(synchronized.pairs[0].id, id);
  assert.equal(synchronized.data.cards.filter((record) => record.archivedAt === null).length, 1);
});

test('sync preserves identities through reordering and archives changed or deleted cards', async () => {
  const initial = await reconcileFsrsData(project, [pair('alpha'), pair('beta')], null);
  const firstIds = new Map(initial.pairs.map((item) => [item.key, item.id]));

  const reordered = await reconcileFsrsData(
    project,
    [pair('beta'), pair('alpha')],
    JSON.stringify(initial.data)
  );
  assert.equal(reordered.pairs[0].id, firstIds.get('beta'));
  assert.equal(reordered.pairs[1].id, firstIds.get('alpha'));

  const edited = await reconcileFsrsData(
    project,
    [pair('alpha', 'changed answer')],
    JSON.stringify(reordered.data)
  );
  assert.notEqual(edited.pairs[0].id, firstIds.get('alpha'));
  assert.ok(edited.data.cards.filter((card) => card.archivedAt !== null).length >= 2);
  assert.ok(edited.data.cards.some((card) => card.id === edited.pairs[0].id && card.archivedAt === null));
});

test('ambiguous duplicate-count changes archive old records instead of reassigning them', async () => {
  const duplicates = await reconcileFsrsData(project, [pair('same'), pair('same')], null);
  const oldIds = new Set(duplicates.pairs.map((item) => item.id));
  assert.equal(oldIds.size, 2);

  const reduced = await reconcileFsrsData(
    project,
    [pair('same')],
    JSON.stringify(duplicates.data)
  );
  assert.equal(reduced.data.cards.filter((card) => card.archivedAt === null).length, 1);
  assert.equal(reduced.data.cards.filter((card) => card.archivedAt !== null).length, 2);
  assert.ok(!oldIds.has(reduced.pairs[0].id));
});

test('corrupt or incompatible FSRS data is rejected without replacing its source', async () => {
  const corrupt = '{not json';
  await assert.rejects(reconcileFsrsData(project, [pair('alpha')], corrupt), /기록을 읽지 못했습니다/);
  assert.equal(corrupt, '{not json');

  const incompatible = JSON.stringify({ schemaVersion: 999, cards: [] });
  await assert.rejects(reconcileFsrsData(project, [pair('alpha')], incompatible), /기록을 읽지 못했습니다/);
});

test('new workbooks contain only A=key, B=value, C=hint', async () => {
  const file = new File(
    [await createStudyWorkbook([['alpha', 'answer', 'clue']])],
    'new-format.xlsx'
  );
  const normalized = await normalizeStudyWorkbook(file);
  assert.equal(normalized.converted, false);
  const pairs = await parseStudyFile(file);
  assert.deepEqual(
    { key: pairs[0].key, value: pairs[0].value, hint: pairs[0].hint },
    { key: 'alpha', value: 'answer', hint: 'clue' }
  );
  const zip = await openZip(await file.arrayBuffer());
  const sheet = await zip.read('xl/worksheets/sheet1.xml');
  assert.match(sheet, /dimension ref="A1:C2"/);
  assert.doesNotMatch(sheet, /<c r="D/);
});

test('legacy status-column workbooks are parsed and edited into the three-column format', async () => {
  const file = new File(
    ['Understood status,key,value,hint\r\n2025-01-02,alpha,answer,clue\r\n'],
    'legacy.csv',
    { type: 'text/csv' }
  );
  const normalized = await normalizeStudyWorkbook(file);
  assert.equal(normalized.converted, true);
  const [pair] = await parseStudyFile(normalized.file);
  assert.deepEqual(
    { key: pair.key, value: pair.value, hint: pair.hint },
    { key: 'alpha', value: 'answer', hint: 'clue' }
  );

  const grid = await readSheetGrid(normalized.file);
  assert.deepEqual(grid[0], ['key', 'value', 'hint']);
  assert.deepEqual(grid[1], ['alpha', 'answer', 'clue']);

  const converted = new File([await writeSheetGrid(normalized.file, grid)], file.name, { type: file.type });
  assert.deepEqual(await readSheetGrid(converted), grid);
  const convertedPairs = await parseStudyFile(converted);
  assert.deepEqual(
    { key: convertedPairs[0].key, value: convertedPairs[0].value, hint: convertedPairs[0].hint },
    { key: 'alpha', value: 'answer', hint: 'clue' }
  );
});

test('resizing adjacent edit columns preserves the total width and enforces minimums', () => {
  const widths = [250, 300, 354];
  const widened = resizeAdjacentColumns(widths, 0, 80);
  assert.deepEqual(widened, [330, 220, 354]);
  assert.equal(widened.reduce((sum, width) => sum + width, 0), 904);

  const clamped = resizeAdjacentColumns(widths, 1, 1000);
  assert.deepEqual(clamped, [250, 554, 100]);
  assert.deepEqual(normalizeColumnWidths(widened), [330 / 904, 220 / 904, 354 / 904]);
});

test('FSRS status categories distinguish new, due, and future cards', () => {
  const now = Date.parse('2026-10-30T00:00:00.000Z');
  assert.equal(fsrsReviewCategory({ state: FsrsState.New, dueAt: '2026-11-01T00:00:00.000Z' }, now), 'new');
  assert.equal(fsrsReviewCategory({ state: FsrsState.Review, dueAt: '2026-10-29T00:00:00.000Z' }, now), 'due');
  assert.equal(fsrsReviewCategory({ state: FsrsState.Review, dueAt: '2026-10-31T00:00:00.000Z' }, now), 'scheduled');
});

test('FSRS retrievability uses the model forgetting curve and decreases with elapsed time', async () => {
  const reviewedAt = new Date();
  const initial = await reconcileFsrsData(project, [pair('alpha')], null);
  const cardId = initial.pairs[0].id;
  const reviewed = await reviewFsrsCard(project, initial.data, cardId, FsrsRating.Good, reviewedAt);
  const card = reviewed.data.cards[0].card;
  const probabilityNow = await getFsrsRetrievability(card, reviewedAt);
  const probabilityLater = await getFsrsRetrievability(
    card,
    new Date(reviewedAt.getTime() + 10 * 24 * 60 * 60 * 1000)
  );

  assert.ok(probabilityNow > probabilityLater);
  assert.ok(probabilityLater > 0 && probabilityLater < 1);
});

test('Again and Hard are saved and change the next interval and forgetting curve', async () => {
  const initial = await reconcileFsrsData(project, [pair('again'), pair('hard')], null);
  const now = new Date();
  const againId = initial.pairs.find((item) => item.key === 'again').id;
  const hardId = initial.pairs.find((item) => item.key === 'hard').id;

  const again = await reviewFsrsCard(project, initial.data, againId, FsrsRating.Again, now);
  const hard = await reviewFsrsCard(project, initial.data, hardId, FsrsRating.Hard, now);
  const againRecord = again.data.cards.find((record) => record.id === againId);
  const hardRecord = hard.data.cards.find((record) => record.id === hardId);
  const probabilitySoon = async (record) =>
    getFsrsRetrievability(record.card, new Date(now.getTime() + 60_000));

  assert.equal(againRecord.reviewLogs.at(-1).rating, FsrsRating.Again);
  assert.equal(hardRecord.reviewLogs.at(-1).rating, FsrsRating.Hard);
  assert.ok(againRecord.card.stability < hardRecord.card.stability);
  assert.ok(Date.parse(againRecord.card.dueAt) < Date.parse(hardRecord.card.dueAt));
  assert.ok(await probabilitySoon(againRecord) < await probabilitySoon(hardRecord));
});

test('learning status indexes valid pairs and counts malformed entries', () => {
  const validPair = { id: 'card-1', key: 'alpha', value: 'answer', hint: '' };
  const { byId, invalidCount } = indexStudyPairs([validPair, undefined, { key: 'missing id' }]);

  assert.equal(byId.get('card-1'), validPair);
  assert.equal(invalidCount, 2);
  assert.deepEqual(indexStudyPairs(null), { byId: new Map(), invalidCount: 0 });
});

test('A4 learning priority sorts reviewed cards first and starts unreviewed cards on a fresh page', () => {
  const reviewed = sortByRecallProbability([
    { id: 'strong', recallProbability: 0.9 },
    { id: 'weak', recallProbability: 0.2 },
    { id: 'middle', recallProbability: 0.5 },
    { id: 'very-weak', recallProbability: 0.1 },
    { id: 'good', recallProbability: 0.7 },
  ]);
  assert.deepEqual(reviewed.map((item) => item.id), ['very-weak', 'weak', 'middle', 'good', 'strong']);

  const pages = buildPriorityPages(
    reviewed.map((item) => ({ ...item, key: item.id, value: 'answer' })),
    [{ id: 'new-1', key: 'new one', value: 'answer' }],
    2
  );
  assert.deepEqual(pages.map((page) => page.section), ['reviewed', 'reviewed', 'unreviewed']);
  assert.deepEqual(pages.map((page) => page.sectionStart), [true, false, true]);
  assert.deepEqual(pages.flatMap((page) => page.rows.map((row) => row.index)), [1, 2, 3, 4, 5, 6]);
});

test('desired retention changes future intervals and exam summary reports pacing', async () => {
  const initial = await reconcileFsrsData(project, [pair('a'), pair('b')], null);
  const id = initial.pairs[0].id;
  const start = new Date('2026-01-01T00:00:00Z');
  let data = initial.data;
  for (let i = 0; i < 3; i += 1) {
    data = (await reviewFsrsCard(project, data, id, FsrsRating.Good, new Date(start.getTime() + i * 3 * 86400000))).data;
  }
  const record = data.cards.find((item) => item.id === id);
  const at = new Date(start.getTime() + 9 * 86400000);
  const low = await previewFsrsCard(record.card, at, 0.8);
  const high = await previewFsrsCard(record.card, at, 0.97);
  const good = (list) => list.find((item) => item.grade === FsrsRating.Good).dueAt;
  assert.ok(good(low) > good(high));

  const exam = new Date(at.getTime() + 30 * 86400000);
  const summary = await summarizeExamReadiness(data.cards, exam, 0.9, at);
  assert.equal(summary.daysLeft, 30);
  assert.equal(summary.newCount, 1);
  assert.equal(summary.newPerDay, 1);
  assert.ok(summary.averageRetention > 0 && summary.averageRetention < 1);
});

test('exam date caps the next review so it lands before the exam', async () => {
  const initial = await reconcileFsrsData(project, [pair('a')], null);
  const card = initial.data.cards[0].card;
  const now = new Date('2026-11-01T09:00:00');
  assert.equal(examMaximumInterval('', now), 36500);
  const free = await previewFsrsCard(card, now, 0.9, '');
  const capped = await previewFsrsCard(card, now, 0.9, '2026-11-03');
  const easy = (list) => list.find((item) => item.grade === FsrsRating.Easy).dueAt;
  assert.ok(easy(capped) <= new Date('2026-11-03T00:00:00'));
  assert.ok(easy(capped) <= easy(free));
  let data = initial.data;
  for (let i = 0; i < 4; i += 1) {
    data = (await reviewFsrsCard(project, data, data.cards[0].id, FsrsRating.Easy, new Date(now.getTime() + i * 86400000 * 0.01), 0.9, '2026-12-01')).data;
  }
  assert.ok(Date.parse(data.cards[0].card.dueAt) <= new Date('2026-12-01T00:00:00').getTime());
});

test('rescheduling applies retention and exam date to cards already in review', async () => {
  const initial = await reconcileFsrsData(project, [pair('a'), pair('b')], null);
  const id = initial.pairs[0].id;
  const start = new Date('2026-01-01T00:00:00Z');
  let data = initial.data;
  for (let i = 0; i < 4; i += 1) {
    data = (await reviewFsrsCard(project, data, id, FsrsRating.Good, new Date(start.getTime() + i * 5 * 86400000))).data;
  }
  const before = data.cards.find((item) => item.id === id).card;
  assert.equal(before.state, FsrsState.Review);
  const same = await rescheduleFsrsData(project, data, 0.9, '', new Date(start.getTime() + 30 * 86400000));
  const lower = await rescheduleFsrsData(project, data, 0.8, '', new Date(start.getTime() + 30 * 86400000));
  const higher = await rescheduleFsrsData(project, data, 0.97, '', new Date(start.getTime() + 30 * 86400000));
  const due = (result) => Date.parse(result.data.cards.find((item) => item.id === id).card.dueAt);
  assert.ok(due(lower) > due(same) && due(same) > due(higher));
  const capped = await rescheduleFsrsData(project, data, 0.8, '2026-02-10', new Date(start.getTime() + 30 * 86400000));
  assert.ok(due(capped) <= new Date('2026-02-10T00:00:00').getTime());
  const untouched = capped.data.cards.find((item) => item.id !== id);
  assert.equal(untouched.card.state, FsrsState.New);
  assert.equal(data.cards.find((item) => item.id === id).card.dueAt, before.dueAt);
});

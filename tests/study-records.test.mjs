import assert from 'node:assert/strict';
import test from 'node:test';
import { appendStudySession, emptyStudyRecords, normalizeStudyRecords, removeStudySession, summarizeStudySessions } from '../lib/studyRecords.js';

test('appendStudySession adds sessions and stores nothing else', () => {
  const first = appendStudySession(emptyStudyRecords(), { date: '2026-11-04', itemCount: 2, durationSeconds: 61.9 });
  const second = appendStudySession(first, { date: '2026-11-05', itemCount: 1, durationSeconds: 10 });
  assert.deepEqual(Object.keys(second).sort(), ['sessions', 'version']);
  assert.deepEqual(second.sessions[0], { date: '2026-11-04', itemCount: 2, durationSeconds: 61, mode: 'flashcard' });
  assert.equal(second.sessions.length, 2);
});

test('preserves session type and combines flashcard and skimming study time', () => {
  const records = appendStudySession(
    { version: 1, sessions: [{ date: '2026-11-04', itemCount: 2, durationSeconds: 61 }] },
    { date: '2026-11-04', itemCount: 0, durationSeconds: 90, mode: 'skimming' }
  );
  assert.deepEqual(records.sessions.map(({ mode }) => mode), ['flashcard', 'skimming']);
  assert.deepEqual(summarizeStudySessions(records.sessions), { itemCount: 2, durationSeconds: 151 });
});

test('normalizeStudyRecords drops malformed entries', () => {
  const result = normalizeStudyRecords({
    keys: [{ key: 'x', count: 3 }],
    sessions: [
      { date: 'bad' },
      { date: '2026-11-04', itemCount: 1, durationSeconds: 5 },
      { date: '2026-11-05', itemCount: 0, durationSeconds: 30 },
    ],
  });
  assert.equal(result.keys, undefined);
  assert.equal(result.sessions.length, 2);
  assert.equal(result.sessions[0].mode, 'flashcard');
  assert.equal(result.sessions[1].mode, 'skimming');
  assert.deepEqual(normalizeStudyRecords(null), emptyStudyRecords());
});

test('removeStudySession removes only the selected round', () => {
  const records = {
    version: 1,
    sessions: [
      { date: '2026-11-04', itemCount: 2, durationSeconds: 61 },
      { date: '2026-11-04', itemCount: 3, durationSeconds: 90 },
      { date: '2026-11-05', itemCount: 1, durationSeconds: 10 },
    ],
  };
  assert.deepEqual(removeStudySession(records, 1), {
    version: 1,
    sessions: [
      { ...records.sessions[0], mode: 'flashcard' },
      { ...records.sessions[2], mode: 'flashcard' },
    ],
  });
  assert.equal(records.sessions.length, 3);
  assert.throws(() => removeStudySession(records, 3), RangeError);
});

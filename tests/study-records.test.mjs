import assert from 'node:assert/strict';
import test from 'node:test';
import { appendStudySession, emptyStudyRecords, normalizeStudyRecords, removeStudySession } from '../lib/studyRecords.js';

test('appendStudySession adds sessions and stores nothing else', () => {
  const first = appendStudySession(emptyStudyRecords(), { date: '2026-11-04', itemCount: 2, durationSeconds: 61.9 });
  const second = appendStudySession(first, { date: '2026-11-05', itemCount: 1, durationSeconds: 10 });
  assert.deepEqual(Object.keys(second).sort(), ['sessions', 'version']);
  assert.deepEqual(second.sessions[0], { date: '2026-11-04', itemCount: 2, durationSeconds: 61 });
  assert.equal(second.sessions.length, 2);
});

test('normalizeStudyRecords drops malformed entries', () => {
  const result = normalizeStudyRecords({ keys: [{ key: 'x', count: 3 }], sessions: [{ date: 'bad' }, { date: '2026-11-04', itemCount: 1, durationSeconds: 5 }] });
  assert.equal(result.keys, undefined);
  assert.equal(result.sessions.length, 1);
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
    sessions: [records.sessions[0], records.sessions[2]],
  });
  assert.equal(records.sessions.length, 3);
  assert.throws(() => removeStudySession(records, 3), RangeError);
});

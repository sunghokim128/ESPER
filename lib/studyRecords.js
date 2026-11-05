export function emptyStudyRecords() {
  return { version: 1, sessions: [] };
}

export function normalizeStudyRecords(value) {
  const records = emptyStudyRecords();
  if (!value || typeof value !== 'object') return records;
  if (Array.isArray(value.sessions)) {
    records.sessions = value.sessions
      .filter((item) => item && /^\d{4}-\d{2}-\d{2}/.test(String(item.date ?? '')))
      .map((item) => ({
        date: String(item.date).slice(0, 10),
        itemCount: Math.max(0, Math.floor(Number(item.itemCount)) || 0),
        durationSeconds: Math.max(0, Math.floor(Number(item.durationSeconds)) || 0),
      }));
  }
  return records;
}

export function appendStudySession(records, session) {
  const next = normalizeStudyRecords(records);
  next.sessions.push({
    date: session.date,
    itemCount: Math.max(0, Math.floor(session.itemCount)),
    durationSeconds: Math.max(0, Math.floor(session.durationSeconds)),
  });
  return next;
}

export function removeStudySession(records, index) {
  const next = normalizeStudyRecords(records);
  if (!Number.isInteger(index) || index < 0 || index >= next.sessions.length) {
    throw new RangeError('Study session index is out of range.');
  }
  next.sessions.splice(index, 1);
  return next;
}

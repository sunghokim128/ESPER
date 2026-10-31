export function indexStudyPairs(pairs) {
  const sourcePairs = Array.isArray(pairs) ? pairs : [];
  const validPairs = sourcePairs.filter((pair) =>
    pair && typeof pair === 'object' && typeof pair.id === 'string' && pair.id.length > 0
  );
  return {
    byId: new Map(validPairs.map((pair) => [pair.id, pair])),
    invalidCount: sourcePairs.length - validPairs.length,
  };
}

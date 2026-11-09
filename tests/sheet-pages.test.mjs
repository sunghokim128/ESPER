import assert from 'node:assert/strict';
import test from 'node:test';
import { buildPages } from '../lib/sheetPages.js';

test('A4 pagination accounts for full-width characters and expanded hints', () => {
  const items = Array.from({ length: 5 }, (_, index) => ({
    id: `card-${index}`,
    key: '가'.repeat(30),
    value: '답'.repeat(30),
    hint: '예문 '.repeat(300),
  }));
  const pages = buildPages(items);
  const rows = pages.flat();

  assert.ok(pages.length > 1);
  assert.deepEqual(rows.map((row) => row.id), items.map((item) => item.id));
  assert.ok(rows.every((row) => row.estimatedHeightMm > 8.2));

  for (const page of pages) {
    for (const column of [0, 1]) {
      const height = page
        .filter((row) => row.column === column)
        .reduce((total, row) => total + row.estimatedHeightMm, 0);
      assert.ok(height <= 30 * 8.2);
    }
  }
});

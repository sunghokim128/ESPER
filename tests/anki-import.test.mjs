import assert from 'node:assert/strict';
import test from 'node:test';
import { buildAnkiRows, previewAnkiCards } from '../lib/ankiImport.js';

test('yields to the event loop while preparing a large Anki deck', async () => {
  let eventLoopYielded = false;
  setTimeout(() => {
    eventLoopYielded = true;
  }, 0);

  const deck = {
    notes: Array.from({ length: 400 }, (_, index) => ({
      mid: 'basic',
      fields: [`word-${index}`, `meaning-${index}`],
    })),
    findMedia: () => null,
  };
  const result = await buildAnkiRows(deck, 'basic', { key: [0], value: [1], hint: [] });

  assert.equal(eventLoopYielded, true);
  assert.equal(result.rows.length, 400);
});

test('previews one mapped Anki card at a time and keeps embedded image references', () => {
  const deck = {
    notes: [
      { mid: 'basic', fields: ['first', 'first answer'] },
      { mid: 'basic', fields: ['second', 'second <img src="photo.png">'] },
      { mid: 'basic', fields: ['', 'missing key'] },
      { mid: 'other', fields: ['other note', 'other answer'] },
    ],
    findMedia: (source) => source === 'photo.png' ? source : null,
  };
  const mapping = { key: [0], value: [1], hint: [] };
  const first = previewAnkiCards(deck, 'basic', mapping);
  const second = previewAnkiCards(deck, 'basic', mapping, 'newline', 1);
  const pastEnd = previewAnkiCards(deck, 'basic', mapping, 'newline', 2);

  assert.equal(first.total, 2);
  assert.deepEqual(first.card.key, [{ type: 'text', value: 'first' }]);
  assert.deepEqual(second.card.key, [{ type: 'text', value: 'second' }]);
  assert.deepEqual(second.card.value, [
    { type: 'text', value: 'second ' },
    { type: 'image', name: 'photo.png' },
  ]);
  assert.equal(pastEnd.card, null);
});

import assert from 'node:assert/strict';
import test from 'node:test';
import { grayscalePixels } from '../lib/grayscale.js';

test('grayscale conversion replaces RGB channels and preserves alpha', () => {
  const pixels = new Uint8ClampedArray([
    255, 0, 0, 128,
    20, 40, 60, 255,
  ]);

  grayscalePixels(pixels);

  assert.deepEqual([...pixels], [54, 54, 54, 128, 37, 37, 37, 255]);
});

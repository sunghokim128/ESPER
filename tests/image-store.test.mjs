import test from 'node:test';
import assert from 'node:assert/strict';
import { webcrypto } from 'node:crypto';
import { configureImageDirectory, getImageBlob, saveImageFile } from '../lib/imageStore.js';

if (!globalThis.crypto) globalThis.crypto = webcrypto;

function directory() {
  const entries = new Map();
  return {
    entries,
    async getDirectoryHandle(name, { create = false } = {}) {
      const current = entries.get(name);
      if (current?.kind === 'directory') return current;
      if (current) throw new DOMException('Not a directory', 'TypeMismatchError');
      if (!create) throw new DOMException('Directory not found', 'NotFoundError');
      const handle = directory();
      handle.kind = 'directory';
      handle.name = name;
      entries.set(name, handle);
      return handle;
    },
    async getFileHandle(name, { create = false } = {}) {
      const current = entries.get(name);
      if (current?.kind === 'file') return current;
      if (current) throw new DOMException('Not a file', 'TypeMismatchError');
      if (!create) throw new DOMException('File not found', 'NotFoundError');
      const handle = {
        kind: 'file',
        name,
        blob: null,
        async getFile() {
          if (!this.blob) throw new DOMException('File not found', 'NotFoundError');
          return this.blob;
        },
        async createWritable() {
          return {
            async write(blob) {
              handle.nextBlob = blob;
            },
            async close() {
              handle.blob = handle.nextBlob;
              handle.nextBlob = null;
            },
            async abort() {
              handle.nextBlob = null;
            },
          };
        },
      };
      entries.set(name, handle);
      return handle;
    },
    async *entriesIterator() {
      yield* entries;
    },
  };
}

async function imageId(blob) {
  const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', await blob.arrayBuffer()));
  return [...digest.slice(0, 12)].map((byte) => byte.toString(16).padStart(2, '0')).join('');
}

test('stores and reads images only inside their project image folder', async () => {
  const root = directory();
  configureImageDirectory(async () => root);
  const image = new Blob([new Uint8Array([1, 2, 3, 4])], { type: 'image/jpeg' });
  const id = await imageId(image);

  await saveImageFile(id, image, 'Words');

  const project = root.entries.get('Words');
  const images = project.entries.get('images');
  assert.ok(images.entries.has(`${id}.jpg`));
  assert.equal((await getImageBlob(id, 'Words')).size, image.size);
  const savedImage = images.entries.get(`${id}.jpg`).blob;
  await saveImageFile(id, image, 'Words');
  assert.equal(images.entries.get(`${id}.jpg`).blob, savedImage);
  await saveImageFile(id, image, 'Other project');
  assert.ok(root.entries.get('Other project').entries.get('images').entries.has(`${id}.jpg`));
});

test('rejects a damaged image and repairs a corrupt same-ID file on retry', async () => {
  const root = directory();
  configureImageDirectory(async () => root);
  const image = new Blob([new Uint8Array([10, 20, 30])], { type: 'image/jpeg' });
  const corrupt = new Blob([new Uint8Array([0])], { type: 'image/jpeg' });
  const id = await imageId(image);

  await assert.rejects(saveImageFile(id, corrupt, 'Words'), /손상되었거나 ID와 일치하지 않습니다/);
  await saveImageFile(id, image, 'Words');
  const images = root.entries.get('Words').entries.get('images');
  const file = images.entries.get(`${id}.jpg`);
  file.blob = corrupt;

  await saveImageFile(id, image, 'Words');

  assert.equal((await getImageBlob(id, 'Words')).size, image.size);
});

test('accepts named image IDs used by packaged ESPER samples', async () => {
  const root = directory();
  configureImageDirectory(async () => root);
  const image = new Blob([new Uint8Array([5, 6, 7])], { type: 'image/jpeg' });

  await saveImageFile('sample-eat', image, 'Vocabulary');

  assert.equal((await getImageBlob('sample-eat', 'Vocabulary')).size, image.size);
});

import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createProjectDirectory,
  readFsrsDataFile,
  readLibraryFile,
  readProjectWorkbook,
  readStudyRecordFile,
  removeProjectDirectory,
  projectNameExists,
  projectNameFromFileName,
  projectNameKey,
  writeFsrsDataFile,
  writeLibraryFile,
  writeProjectWorkbook,
  writeStudyRecordFile,
} from '../lib/projectStore.js';

test('uses the uploaded workbook basename as the project folder name', () => {
  assert.equal(projectNameFromFileName('Words.xlsx'), 'Words');
  assert.equal(projectNameFromFileName('phrase-list.CSV'), 'phrase-list');
});

test('normalizes duplicate project names across case and Unicode forms', () => {
  assert.equal(projectNameExists(['Words'], 'words'), true);
  assert.equal(projectNameExists(['Cafe\u0301'], 'Café'), true);
  assert.equal(projectNameExists(['Words'], 'Other'), false);
  assert.equal(projectNameKey('Cafe\u0301'), projectNameKey('Café'));
});

function memoryDirectory() {
  const entries = new Map();
  return {
    kind: 'directory',
    entriesMap: entries,
    async queryPermission() {
      return 'granted';
    },
    async *entries() {
      yield* entries;
    },
    async getDirectoryHandle(name, { create = false } = {}) {
      const current = entries.get(name);
      if (current?.kind === 'directory') return current;
      if (current) throw new DOMException('Not a directory', 'TypeMismatchError');
      if (!create) throw new DOMException('Directory not found', 'NotFoundError');
      const handle = memoryDirectory();
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
    async removeEntry(name, { recursive = false } = {}) {
      const entry = entries.get(name);
      if (!entry) throw new DOMException('Entry not found', 'NotFoundError');
      if (entry.kind === 'directory' && entry.entriesMap.size && !recursive) {
        throw new DOMException('Directory not empty', 'InvalidModificationError');
      }
      entries.delete(name);
    },
  };
}

function fakeIndexedDb(records) {
  const database = {
    objectStoreNames: { contains: () => true },
    transaction() {
      const transaction = {};
      transaction.objectStore = () => ({
        get(key) {
          const request = {};
          queueMicrotask(() => {
            request.result = records.get(key);
            request.onsuccess?.();
            transaction.oncomplete?.();
          });
          return request;
        },
      });
      return transaction;
    },
    close() {},
  };
  return {
    open() {
      const request = {};
      queueMicrotask(() => {
        request.result = database;
        request.onsuccess?.();
      });
      return request;
    },
  };
}

test('writes and reads extensible library preferences', async () => {
  const previousIndexedDb = globalThis.indexedDB;
  const records = new Map([['workbookDirectory', { handle: memoryDirectory() }]]);
  globalThis.indexedDB = fakeIndexedDb(records);

  try {
    const library = {
      version: 1,
      preferences: {
        version: 1,
        language: 'ko',
        futurePreference: { enabled: true },
      },
    };
    assert.equal(await writeLibraryFile(library), true);
    assert.deepEqual(await readLibraryFile(), library);
  } finally {
    if (previousIndexedDb === undefined) delete globalThis.indexedDB;
    else globalThis.indexedDB = previousIndexedDb;
  }
});

test('stores the workbook, FSRS data, and study record inside one project folder', async () => {
  const previousIndexedDb = globalThis.indexedDB;
  const records = new Map([['workbookDirectory', { handle: memoryDirectory() }]]);
  globalThis.indexedDB = fakeIndexedDb(records);

  try {
    const root = records.get('workbookDirectory').handle;
    await createProjectDirectory('Words');
    await writeProjectWorkbook('Words', 'Words.xlsx', new Blob(['workbook']));
    await writeFsrsDataFile('Words', 'Words.xlsx', '{"projectId":"id"}');
    await writeStudyRecordFile('Words', { version: 1, sessions: [{ date: '2026-11-04', itemCount: 3, durationSeconds: 90 }] });

    assert.deepEqual([...root.entriesMap.keys()], ['Words']);
    const project = root.entriesMap.get('Words');
    assert.deepEqual([...project.entriesMap.keys()].sort(), [
      'Words study record.json',
      'Words.json',
      'Words.xlsx',
    ]);
    assert.equal((await readProjectWorkbook('Words', 'Words.xlsx')).size, 8);
    assert.equal(await readFsrsDataFile('Words', 'Words.xlsx'), '{"projectId":"id"}');
    assert.deepEqual(await readStudyRecordFile('Words'), {
      version: 1,
      sessions: [{ date: '2026-11-04', itemCount: 3, durationSeconds: 90 }],
    });
    await assert.rejects(createProjectDirectory('words'), /이미 있어요/);

    await removeProjectDirectory('Words');
    assert.equal(root.entriesMap.has('Words'), false);
  } finally {
    if (previousIndexedDb === undefined) delete globalThis.indexedDB;
    else globalThis.indexedDB = previousIndexedDb;
  }
});

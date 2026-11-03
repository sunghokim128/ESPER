import test from 'node:test';
import assert from 'node:assert/strict';
import { createStudyWorkbook } from '../lib/parseSheet.js';
import { recoverWorkbookProjects } from '../lib/projectRecovery.js';

function workbook(name, content) {
  const blob = new Blob([content]);
  return {
    kind: 'file',
    name,
    async getFile() {
      return Object.assign(blob, { name, lastModified: 123 });
    },
  };
}

function directory(entries) {
  return {
    kind: 'directory',
    async *entries() {
      yield* entries;
    },
  };
}

test('recovers projects from name-matched folders and reuses matching FSRS identity', async () => {
  const wordsWorkbook = workbook(
    'Words.xlsx',
    await createStudyWorkbook([['apple', '사과', 'red fruit']])
  );
  const scienceWorkbook = workbook('Science.csv', 'key,value,hint\nwater,H2O,\n');
  const root = directory([
    ['Words', directory([['Words.xlsx', wordsWorkbook]])],
    ['Science', directory([['Science.csv', scienceWorkbook]])],
    ['unrelated', directory([['other.xlsx', workbook('other.xlsx', 'key,value\nx,y\n')]])],
    ['root.xlsx', workbook('root.xlsx', 'key,value\nx,y\n')],
  ]);
  const requestedProgress = [];
  const recovered = await recoverWorkbookProjects(root, async (projectName, fileName) => {
    requestedProgress.push([projectName, fileName]);
    return projectName === 'Words'
      ? JSON.stringify({
          projectId: 'original-project-id',
          workbookFileName: fileName,
          plan: {
            examDate: '2026-12-03',
            desiredRetention: 0.9,
            dailyNewCards: 18,
            useRecommendedDailyPace: false,
            firstPassDeadline: '2026-11-27',
          },
        })
      : null;
  });

  assert.equal(recovered.projects.length, 2);
  assert.deepEqual(recovered.projects.map(({ fileName }) => fileName), ['Words.xlsx', 'Science.csv']);
  assert.equal(recovered.projects[0].id, 'original-project-id');
  assert.equal(recovered.projects[0].name, 'Words');
  assert.equal(recovered.projects[0].pairs[0].key, 'apple');
  assert.deepEqual(recovered.projects[0].plan, {
    examDate: '2026-12-03',
    desiredRetention: 0.9,
    dailyNewCards: 18,
    useRecommendedDailyPace: false,
    firstPassDeadline: '2026-11-27',
  });
  assert.equal(recovered.projects[1].name, 'Science');
  assert.deepEqual(requestedProgress, [['Words', 'Words.xlsx'], ['Science', 'Science.csv']]);
  assert.deepEqual(recovered.errors, []);
});

test('reports malformed workbooks and FSRS data without dropping valid projects', async () => {
  const root = directory([
    ['valid', directory([['valid.xlsx', workbook(
      'valid.xlsx',
      await createStudyWorkbook([['term', 'definition', '']])
    )]])],
    ['broken', directory([['broken.xlsx', workbook('broken.xlsx', 'not an excel workbook')]])],
  ]);
  const recovered = await recoverWorkbookProjects(root, async (projectName, fileName) =>
    projectName === 'valid' ? '{broken json' : null
  );

  assert.equal(recovered.projects.length, 1);
  assert.equal(recovered.projects[0].fileName, 'valid.xlsx');
  assert.equal(recovered.projects[0].name, 'valid');
  assert.equal(recovered.errors.length, 2);
  assert.match(recovered.errors.join('\n'), /valid\/valid\.xlsx/);
  assert.match(recovered.errors.join('\n'), /broken\/broken\.xlsx/);
});

test('does not recover ambiguous folders with multiple workbook candidates', async () => {
  const root = directory([
    ['Words', directory([
      ['Words.xlsx', workbook('Words.xlsx', 'key,value\nx,y\n')],
      ['Words.csv', workbook('Words.csv', 'key,value\nx,y\n')],
    ])],
  ]);
  const recovered = await recoverWorkbookProjects(root, async () => null);

  assert.deepEqual(recovered.projects, []);
  assert.match(recovered.errors[0], /multiple workbook files/);
});

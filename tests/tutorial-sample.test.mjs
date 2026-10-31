import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { openZip, parseStudyFile } from '../lib/parseSheet.js';
import { readEsprFile } from '../lib/projectShare.js';

const imageIds = ['sample-eat', 'sample-adversity', 'sample-truculent', 'sample-fsrs7', 'sample-retention'];
const samples = [
  {
    language: 'Korean',
    path: '../assets/tutorial/vocabulary-fsrs-7-starter-ko.espr',
    expectedValues: [
      '먹다; 음식물을 입에 넣어 씹고 삼키다.',
      '역경; 삶을 어렵게 만드는 힘들고 불운한 상황.',
      '호전적인, 싸우기 좋아하는; 공격적으로 맞서는.',
      'FSRS-7 간격 반복 알고리즘; 복습 평가로 카드별 기억 상태를 추정해 다음 복습 시점을 정하는 스케줄러.',
      '기억 유지; 배운 내용을 일정 시간이 지난 뒤에도 기억하고 떠올리는 것.',
    ],
  },
  {
    language: 'English',
    path: '../assets/tutorial/vocabulary-fsrs-7-starter-en.espr',
    expectedValues: [
      'To take food into the mouth, chew it, and swallow it.',
      'A difficult or unfavorable situation; hardship.',
      'Eager to argue or fight; aggressively defiant.',
      'A spaced-repetition scheduler that estimates memory from review history and chooses when each card should be reviewed again.',
      'The ability to remember or keep information over time.',
    ],
  },
];

for (const sample of samples) {
  test(`${sample.language} tutorial ESPR has localized values and matching images`, async () => {
    const sampleUrl = new URL(sample.path, import.meta.url);
    const file = new File([await readFile(sampleUrl)], 'tutorial.espr', { type: 'application/x-espr' });
    const shared = await readEsprFile(file);
    const cards = await parseStudyFile(shared.workbook);

    assert.deepEqual(cards.map(({ key, value, hint }) => [key, value, hint]), [
      'Eat', 'Adversity', 'Truculent', 'FSRS-7 algorithm', 'Retention',
    ].map((word, index) => [word, sample.expectedValues[index], `[img:${imageIds[index]}]`]));
    assert.deepEqual(shared.images.map(({ id }) => id).sort(), [...imageIds].sort());
    assert.ok(shared.images.every(({ blob }) =>
      blob.type === 'image/jpeg' && blob.size > 2
    ));
    const zip = await openZip(await file.arrayBuffer());
    const credits = await zip.read('CREDITS.txt');
    assert.match(credits, /CC BY-SA 2.0/);
    assert.match(credits, /Leitner system alternative/);
  });
}

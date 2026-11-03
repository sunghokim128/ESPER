import test from 'node:test';
import assert from 'node:assert/strict';
import { dailyNewCardPlan, firstPassDeadline, manualDailyNewCardPlan } from '../lib/studyPace.js';

test('sets first-pass deadline at 80 percent of the exam period', () => {
  assert.equal(firstPassDeadline('2026-12-03', new Date(2026, 10, 3, 12)), '2026-11-27');
});

test('keeps the configured daily amount unless a faster pace is needed', () => {
  const plan = dailyNewCardPlan({
    examDate: '2026-12-03',
    firstPassDate: '2026-11-27',
    remainingCards: 100,
    configuredDailyCards: 10,
    now: new Date(2026, 10, 3, 12),
  });
  assert.equal(plan.recommended, 10);
  assert.equal(plan.completionDate, '2026-11-12');
});

test('automatically raises the daily amount to meet the first-pass deadline', () => {
  const plan = dailyNewCardPlan({
    examDate: '2026-12-03',
    firstPassDate: '2026-11-13',
    remainingCards: 200,
    configuredDailyCards: 10,
    now: new Date(2026, 10, 3, 12),
  });
  assert.equal(plan.recommended, 20);
  assert.equal(plan.completionDate, '2026-11-12');
});

test('allows more than the default pace when the remaining deck requires it', () => {
  const plan = dailyNewCardPlan({
    examDate: '2026-12-03',
    firstPassDate: '2026-11-13',
    remainingCards: 200,
    configuredDailyCards: 10,
    now: new Date(2026, 10, 3, 12),
  });
  assert.equal(plan.recommended, 20);
});

test('stops recommending new cards after the first-pass deadline', () => {
  const plan = dailyNewCardPlan({
    examDate: '2026-12-03',
    firstPassDate: '2026-11-02',
    remainingCards: 100,
    configuredDailyCards: 10,
    now: new Date(2026, 10, 3, 12),
  });
  assert.equal(plan.recommended, 0);
  assert.equal(plan.isPastDeadline, true);
});

test('manual pace stays independent from D-day and estimates completion from the chosen amount', () => {
  assert.deepEqual(
    manualDailyNewCardPlan({ remainingCards: 25, dailyCards: 7, now: new Date(2026, 10, 3, 12) }),
    { recommended: 7, completionDate: '2026-11-06', isPastDeadline: false }
  );
});

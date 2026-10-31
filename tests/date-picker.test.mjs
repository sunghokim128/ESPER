import assert from 'node:assert/strict';
import test from 'node:test';
import {
  calendarMonthDays,
  isDatePickerEvent,
  localDateValue,
  parseLocalDate,
  shiftCalendarMonth,
} from '../lib/datePicker.js';

test('calendar month navigation crosses year boundaries in both directions', () => {
  assert.equal(shiftCalendarMonth('2026-12', 1), '2027-01');
  assert.equal(shiftCalendarMonth('2027-01', -1), '2026-12');
  assert.equal(shiftCalendarMonth('2028-02', -1), '2028-01');
  assert.equal(shiftCalendarMonth('2028-01', 1), '2028-02');
});

test('calendar month grid includes the right days and weekday-aligned empty cells', () => {
  const february = calendarMonthDays('2028-02', '', '2028-02-29');
  assert.equal(february.length, 35);
  assert.equal(february.filter(Boolean).length, 29);
  assert.equal(february[0], null);
  assert.deepEqual(february.find((day) => day?.value === '2028-02-29'), {
    day: 29, value: '2028-02-29', disabled: false, today: true,
  });

  const april = calendarMonthDays('2026-04', '', '');
  assert.equal(april.length, 35);
  assert.equal(april.filter(Boolean).length, 30);
  assert.equal(april[3].value, '2026-04-01');
});

test('calendar minimum date disables only earlier days', () => {
  const days = calendarMonthDays('2026-11', '2026-11-15', '');
  assert.equal(days.find((day) => day?.value === '2026-11-14').disabled, true);
  assert.equal(days.find((day) => day?.value === '2026-11-15').disabled, false);
  assert.equal(days.find((day) => day?.value === '2026-11-30').disabled, false);
});

test('local dates reject invalid calendar values and format valid values without UTC shifting', () => {
  assert.equal(parseLocalDate('2026-02-29'), null);
  assert.equal(parseLocalDate('not-a-date'), null);
  assert.equal(localDateValue(parseLocalDate('2028-02-29')), '2028-02-29');
});

test('outside-click detection still recognizes a picker after a month button is detached', () => {
  const picker = { matches: (selector) => selector === '.date-picker' };
  const removedButton = { closest: () => null };
  assert.equal(isDatePickerEvent({
    target: removedButton,
    composedPath: () => [removedButton, picker, {}],
  }), true);
  assert.equal(isDatePickerEvent({ target: { closest: () => null }, composedPath: () => [] }), false);
});

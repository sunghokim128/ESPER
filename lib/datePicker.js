export function parseLocalDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) return null;
  const [year, month, day] = value.split('-').map(Number);
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
  return date;
}

export function localDateValue(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

export function shiftCalendarMonth(monthValue, amount) {
  const date = parseLocalDate(`${monthValue}-01`) || new Date();
  date.setMonth(date.getMonth() + amount);
  return localDateValue(date).slice(0, 7);
}

export function calendarMonthDays(monthValue, minValue, todayValue) {
  const parsedMonth = parseLocalDate(`${monthValue}-01`);
  const monthDate = parsedMonth || new Date();
  const firstDay = new Date(monthDate.getFullYear(), monthDate.getMonth(), 1);
  const dayCount = new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0).getDate();
  const leadingDays = firstDay.getDay();
  const cellCount = Math.ceil((leadingDays + dayCount) / 7) * 7;
  const minDate = parseLocalDate(minValue);
  return Array.from({ length: cellCount }, (_, index) => {
    const day = index - leadingDays + 1;
    if (day < 1 || day > dayCount) return null;
    const date = new Date(monthDate.getFullYear(), monthDate.getMonth(), day);
    const value = localDateValue(date);
    return { day, value, disabled: Boolean(minDate && date < minDate), today: value === todayValue };
  });
}

export function isDatePickerEvent(event) {
  const path = event.composedPath?.();
  if (path?.some((node) => node?.matches?.('.date-picker'))) return true;
  return Boolean(event.target?.closest?.('.date-picker'));
}

export function attachDatePicker(picker, { renderCalendar, formatDate, beforeOpen = () => {} }) {
  const input = picker.querySelector('input[type="hidden"]');
  const trigger = picker.querySelector('[data-date-picker-trigger]');
  const calendar = picker.querySelector('[data-date-picker-calendar]');
  trigger.addEventListener('click', () => {
    beforeOpen(calendar);
    calendar.hidden = !calendar.hidden;
    trigger.setAttribute('aria-expanded', String(!calendar.hidden));
    if (!calendar.hidden) {
      const monthValue = calendar.dataset.datePickerMonth ||
        localDateValue(parseLocalDate(input.value) || new Date()).slice(0, 7);
      calendar.dataset.datePickerMonth = monthValue;
      calendar.innerHTML = renderCalendar(monthValue, picker.dataset.dateMin);
    }
  });
  calendar.addEventListener('click', (event) => {
    const monthButton = event.target.closest('[data-date-month-step]');
    if (monthButton) {
      calendar.dataset.datePickerMonth = shiftCalendarMonth(
        calendar.dataset.datePickerMonth,
        Number(monthButton.dataset.dateMonthStep)
      );
      calendar.innerHTML = renderCalendar(calendar.dataset.datePickerMonth, picker.dataset.dateMin);
      return;
    }
    const dayButton = event.target.closest('[data-date-value]');
    if (!dayButton || dayButton.disabled) return;
    input.value = dayButton.dataset.dateValue;
    picker.querySelector('[data-date-picker-label]').textContent = formatDate(input.value);
    calendar.hidden = true;
    trigger.setAttribute('aria-expanded', 'false');
    trigger.focus();
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  });
}

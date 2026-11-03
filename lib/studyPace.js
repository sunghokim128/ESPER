const DAY_MS = 86400000;
export const DEFAULT_DAILY_NEW_CARDS = 10;

function localDateStart(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function validDateValue(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) return false;
  const [year, month, day] = value.split('-').map(Number);
  const date = new Date(year, month - 1, day);
  return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day;
}

export function localStudyDate(date = new Date()) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

export function firstPassDeadline(examDate, now = new Date()) {
  if (!validDateValue(examDate)) return '';
  const today = localDateStart(now);
  const exam = new Date(Number(examDate.slice(0, 4)), Number(examDate.slice(5, 7)) - 1, Number(examDate.slice(8, 10)));
  const daysUntilExam = Math.max(0, Math.round((exam.getTime() - today.getTime()) / DAY_MS));
  today.setDate(today.getDate() + Math.floor(daysUntilExam * 0.8));
  return localStudyDate(today);
}

export function dailyNewCardPlan({
  examDate = '',
  firstPassDate = '',
  remainingCards,
  configuredDailyCards = DEFAULT_DAILY_NEW_CARDS,
  now = new Date(),
}) {
  const remaining = Math.max(0, Math.floor(Number(remainingCards) || 0));
  const configured = Math.min(1000, Math.max(1, Math.floor(Number(configuredDailyCards) || DEFAULT_DAILY_NEW_CARDS)));
  const deadline = validDateValue(firstPassDate)
    ? firstPassDate
    : firstPassDeadline(examDate, now);
  const today = localDateStart(now);
  const deadlineDate = validDateValue(deadline)
    ? new Date(Number(deadline.slice(0, 4)), Number(deadline.slice(5, 7)) - 1, Number(deadline.slice(8, 10)))
    : null;
  const isPastDeadline = Boolean(deadlineDate && deadlineDate < today);
  const availableDays = deadlineDate
    ? Math.max(1, Math.ceil((deadlineDate.getTime() - today.getTime()) / DAY_MS))
    : 0;
  const requiredPerDay = availableDays ? Math.ceil(remaining / availableDays) : 0;
  const recommended = !remaining || isPastDeadline
    ? 0
    : Math.max(configured, requiredPerDay);
  const completionDate = remaining === 0
    ? localStudyDate(today)
    : recommended > 0
      ? (() => {
          const date = new Date(today);
          date.setDate(date.getDate() + Math.ceil(remaining / recommended) - 1);
          return localStudyDate(date);
        })()
      : '';
  return { deadline, recommended, completionDate, isPastDeadline };
}

export function manualDailyNewCardPlan({ remainingCards, dailyCards, now = new Date() }) {
  const remaining = Math.max(0, Math.floor(Number(remainingCards) || 0));
  const daily = Math.min(1000, Math.max(1, Math.floor(Number(dailyCards) || DEFAULT_DAILY_NEW_CARDS)));
  const completionDate = remaining === 0
    ? localStudyDate(now)
    : (() => {
        const date = localDateStart(now);
        date.setDate(date.getDate() + Math.ceil(remaining / daily) - 1);
        return localStudyDate(date);
      })();
  return { recommended: daily, completionDate, isPastDeadline: false };
}

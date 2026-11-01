import {
  exportStudyRecordWorkbook,
  parseStudyFile,
  normalizeStudyWorkbook,
  readStudyRecordWorkbook,
  readSheetGrid,
  writeSheetGrid,
  createStudyWorkbook,
} from './lib/parseSheet.js';
import { normalizeColumnWidths, resizeAdjacentColumns } from './lib/columnSizing.js';
import { indexStudyPairs } from './lib/studyPairs.js';
import { AnkiImportError, buildAnkiRows, countFieldImages, loadAnkiDeck, previewAnkiCards } from './lib/ankiImport.js';
import { configureImageDirectory, getImageBlob, getImageUrl, hydrateImages, processImage, releaseProjectImages, saveImageBlob, saveImageFile } from './lib/imageStore.js';
import { createEsprBlob, readEsprFile, ESPR_EXTENSION } from './lib/projectShare.js';
import { attachDatePicker, calendarMonthDays, isDatePickerEvent, localDateValue, parseLocalDate } from './lib/datePicker.js';
import { TRANSPARENT_PIXEL, imageTag, renderImageText, splitImageText, stripImageTags } from './lib/imageText.js';
import {
  SHEET_H_PX,
  SHEET_W_PX,
  buildPages,
  buildPriorityPages,
  shuffle,
  sortByRecallProbability,
} from './lib/sheetPages.js';
import {
  FsrsRating,
  FsrsState,
  fsrsReviewCategory,
  getFsrsRetrievability,
  DEFAULT_DESIRED_RETENTION,
  MIN_DESIRED_RETENTION,
  MAX_DESIRED_RETENTION,
  normalizeDesiredRetention,
  summarizeExamReadiness,
  rescheduleFsrsData,
  prepareFsrsSync,
  previewFsrsCard,
  reviewFsrsCard,
  saveFsrsData,
  synchronizeFsrsProject,
} from './lib/fsrs.js';
import {
  deleteProjectRecord,
  createProjectDirectory,
  getProject,
  getProjectOrder,
  getLanguage,
  getStudySettings,
  saveLanguage,
  getWorkbookDirectory,
  getStoredWorkbookDirectory,
  listProjects,
  projectNameFromFileName,
  projectNameKey,
  projectNameExists,
  removeProjectDirectory,
  readFsrsDataFile,
  saveProject,
  saveWorkbookDirectory,
  writeProjectWorkbook,
  readProjectWorkbook,
  saveProjectOrder,
  saveStudySettings as saveStudySettingsToStore,
  readStudyRecordFile,
  writeStudyRecordFile,
} from './lib/projectStore.js';
import { recoverWorkbookProjects } from './lib/projectRecovery.js';

const EDIT_COLUMN_WIDTHS_KEY = 'esper.editColumnWidths.v1';
const DEFAULT_STUDY_SETTINGS = {
  showHintOnBack: true,
  swapKeyValue: false,
  randomizeKeyValue: false,
  fontSizes: { front: 72, back: 43 },
  examDate: '',
  desiredRetention: DEFAULT_DESIRED_RETENTION,
  imageSizes: { study: 'fill', a4: 'medium', status: 'medium' },
  shortcuts: {
    flip: 'Space',
    hint: 'ArrowDown',
    again: '1',
    hard: '2',
    good: '3',
    easy: '4',
    previous: 'ArrowLeft',
    speak: 's',
    exit: 'Escape',
  },
};
const STUDY_ACTIONS = [
  ['flip', '카드 뒤집기'],
  ['hint', '힌트 보기'],
  ['again', 'Again'],
  ['hard', 'Hard'],
  ['good', 'Good'],
  ['easy', 'Easy'],
  ['speak', '발음 듣기'],
  ['previous', '이전 카드'],
  ['exit', '회차 종료'],
];
const STUDY_ACTION_LABELS_EN = {
  flip: 'Flip card',
  hint: 'Show hint',
  again: 'Again',
  hard: 'Hard',
  good: 'Good',
  easy: 'Easy',
  speak: 'Pronounce',
  previous: 'Previous card',
  exit: 'Exit Study',
};

// 언어 설정은 화면 표시만 바꾸며 엑셀 데이터는 수정하지 않습니다.
let language = 'en';
const isEnglish = () => language === 'en';
const L = (ko, en) => (language === 'en' ? en : ko);
const studyActionLabel = ([action, ko]) => L(ko, STUDY_ACTION_LABELS_EN[action]);
const PROMPT_FILE_URLS = {
  ko: './assets/guide/downloads/word-list-generation-prompt.md',
  en: './assets/guide/downloads/word-list-generation-prompt-en.md',
};
const SUPPORT_QR_SRC = { ko: 'assets/guide/support-qr.png', en: 'assets/guide/support-qr-eng.png' };

const appEl = document.getElementById('app');
const printEl = document.getElementById('printRoot');
const studyTooltip = document.createElement('div');
studyTooltip.className = 'study-tooltip';
studyTooltip.hidden = true;
studyTooltip.setAttribute('role', 'tooltip');
document.body.appendChild(studyTooltip);

let studyTooltipTarget = null;
function positionStudyTooltip(clientX, clientY) {
  const rect = studyTooltip.getBoundingClientRect();
  const left = Math.max(8 + rect.width / 2, Math.min(window.innerWidth - 8 - rect.width / 2, clientX));
  const top = Math.max(8, clientY - rect.height - 10);
  studyTooltip.style.left = `${left}px`;
  studyTooltip.style.top = `${top}px`;
}

function showStudyTooltip(target, clientX, clientY) {
  const action = target.dataset.studyAction;
  const id = target.dataset.id;
  if (action === 'value') {
    studyTooltip.textContent = isA4ValueVisible(id)
      ? L('value 가리기', 'Hide value')
      : L('value 보기', 'Show value');
  } else if (action === 'hint') {
    studyTooltip.textContent = state.expandedA4Hints.has(id)
      ? L('hint 숨기기', 'Hide hint')
      : L('hint 보기', 'Show hint');
  } else {
    studyTooltip.textContent = target.dataset.studyHint || '';
  }
  if (!studyTooltip.textContent) return;
  studyTooltipTarget = target;
  studyTooltip.hidden = false;
  positionStudyTooltip(clientX, clientY);
}

function hideStudyTooltip() {
  studyTooltipTarget = null;
  studyTooltip.hidden = true;
}

appEl.addEventListener('mouseover', (event) => {
  const target = event.target.closest('[data-study-action], [data-study-hint]');
  if (!target || target.contains(event.relatedTarget)) return;
  showStudyTooltip(target, event.clientX, event.clientY);
});
appEl.addEventListener('mousemove', (event) => {
  if (studyTooltipTarget && (event.target === studyTooltipTarget || studyTooltipTarget.contains(event.target))) {
    positionStudyTooltip(event.clientX, event.clientY);
  }
});
appEl.addEventListener('mouseout', (event) => {
  if (
    studyTooltipTarget &&
    (event.target === studyTooltipTarget || studyTooltipTarget.contains(event.target)) &&
    !studyTooltipTarget.contains(event.relatedTarget)
  ) {
    hideStudyTooltip();
  }
});
appEl.addEventListener('focusin', (event) => {
  const target = event.target.closest('[data-study-action], [data-study-hint]');
  if (!target) return;
  const rect = target.getBoundingClientRect();
  showStudyTooltip(target, rect.left + rect.width / 2, rect.top);
});
appEl.addEventListener('focusout', (event) => {
  if (
    studyTooltipTarget &&
    (event.target === studyTooltipTarget || studyTooltipTarget.contains(event.target)) &&
    !studyTooltipTarget.contains(event.relatedTarget)
  ) {
    hideStudyTooltip();
  }
});
appEl.addEventListener('dragstart', (event) => {
  const card = event.target.closest?.('[data-project-card]');
  if (!card) return;
  if (state.busy || processingProjectIds.size) {
    event.preventDefault();
    return;
  }
  draggedProjectId = card.dataset.projectCard;
  event.dataTransfer.effectAllowed = 'move';
  event.dataTransfer.setData('text/plain', draggedProjectId);
  card.classList.add('is-dragging');
});
appEl.addEventListener('dragover', (event) => {
  const card = event.target.closest('[data-project-card]');
  if (!draggedProjectId || !card || card.dataset.projectCard === draggedProjectId) return;
  event.preventDefault();
  event.dataTransfer.dropEffect = 'move';
  appEl.querySelectorAll('[data-project-card]').forEach((item) => {
    item.classList.remove('is-drop-before', 'is-drop-after');
  });
  const rect = card.getBoundingClientRect();
  const deltaX = (event.clientX - rect.left - rect.width / 2) / rect.width;
  const deltaY = (event.clientY - rect.top - rect.height / 2) / rect.height;
  const insertAfter = Math.abs(deltaX) > Math.abs(deltaY) ? deltaX > 0 : deltaY > 0;
  card.classList.add(insertAfter ? 'is-drop-after' : 'is-drop-before');
});
appEl.addEventListener('drop', (event) => {
  const card = event.target.closest('[data-project-card]');
  if (!draggedProjectId || !card) return;
  event.preventDefault();
  const insertAfter = card.classList.contains('is-drop-after');
  const sourceId = draggedProjectId;
  draggedProjectId = '';
  appEl.querySelectorAll('[data-project-card]').forEach((item) => {
    item.classList.remove('is-dragging', 'is-drop-before', 'is-drop-after');
  });
  moveProject(sourceId, card.dataset.projectCard, insertAfter);
});
appEl.addEventListener('change', (event) => {
  const select = event.target.closest?.('[data-language-select]');
  if (select) changeLanguage(select.value);
});
appEl.addEventListener('click', (event) => {
  const navButton = event.target.closest?.('[data-global-nav]');
  if (!navButton || navButton.disabled) return;
  event.preventDefault();
  const destination = navButton.dataset.globalNav;
  const previousPage = state.page;
  window.setTimeout(() => {
    navigateGlobalPage(destination).catch((error) => {
      const detail = error instanceof Error ? error.message : String(error);
      console.error(`Could not navigate to ${destination}.`, error);
      state.page = previousPage;
      const message = L(`페이지를 열지 못했어요: ${detail}`, `Could not open the page: ${detail}`);
      if (previousPage === 'study') state.studyError = message;
      else if (previousPage === 'study-records') state.studyRecordsError = message;
      else state.error = message;
      render();
    });
  }, 0);
}, true);
appEl.addEventListener('click', (event) => {
  const choice = event.target.closest?.('[data-choose-language]');
  if (choice) changeLanguage(choice.dataset.chooseLanguage);
});
appEl.addEventListener('dragend', () => {
  draggedProjectId = '';
  appEl.querySelectorAll('[data-project-card]').forEach((item) => {
    item.classList.remove('is-dragging', 'is-drop-before', 'is-drop-after');
  });
});

configureImageDirectory(() => getWorkbookDirectory());

const state = {
  page: 'workspace',
  language: null,
  languageChoicePending: false,
  projects: [],
  project: null,
  directoryName: '',
  directoryHandle: null,
  directorySaved: false,
  directoryConnected: false,
  pairs: [],
  fsrsData: null,
  pendingReview: null,
  a4Order: [],
  a4Query: '',
  a4Sort: 'order',
  a4RecallRates: new Map(),
  a4PriorityData: null,
  a4PriorityPairs: null,
  a4PriorityError: '',
  revealedA4Values: new Set(),
  hiddenA4Values: new Set(),
  expandedA4Hints: new Set(),
  showAllA4Values: true,
  sourceName: '',
  sourceFile: null,
  title: '',
  a4Title: null,
  a4Info: null,
  busy: false,
  processingMessage: '',
  error: '',
  notice: '',
  directoryAccessIssue: '',
  generationPrompt: '',
  generationPromptError: '',
  projectOrder: [],
  editRows: [],
  editDirty: false,
  editMessage: '',
  editMessageIsError: false,
  editFocus: null,
  anki: null,
  editActiveCell: null,
  study: null,
  studySettings: structuredClone(DEFAULT_STUDY_SETTINGS),
  studyShortcutsOpen: false,
  studyKeyCapture: '',
  studyKeyError: '',
  studyError: '',
  studySaving: false,
  studyRecords: null,
  studyCalendarMonth: '',
  studyRecordsOpen: { sessions: false, keys: false },
  studyRecordsLoading: false,
  studyRecordsError: '',
  learningStatusFilter: 'all',
  learningStatusShowValue: false,
  learningStatusPlanMessage: '',
  planPanelOpen: false,
  learningStatusQuery: '',
};

const processingProjectIds = new Set();
window.addEventListener('beforeunload', (event) => {
  const hasPendingWork =
    state.busy ||
    Boolean(state.processingMessage) ||
    Boolean(state.anki?.busy) ||
    processingProjectIds.size > 0 ||
    state.studySaving ||
    Boolean(state.pendingReview) ||
    (state.page === 'edit' && state.editDirty);
  if (!hasPendingWork) return;
  event.preventDefault();
  event.returnValue = L(
    '처리 중인 작업이 있어요. 페이지를 떠나면 작업이 중단되거나 데이터가 저장되지 않을 수 있습니다.',
    'A task is in progress. Leaving this page may interrupt it or prevent data from being saved.'
  );
});

let draggedProjectId = '';
let studyTimerInterval = null;
let learningStatusInterval = null;
let learningStatusRefreshPromise = null;
let learningStatusRefreshRequested = false;
let a4PriorityInterval = null;
let a4PriorityRefreshPromise = null;
let a4PriorityRefreshGeneration = 0;
let workspaceExamGaugeInterval = null;
let workspaceExamGaugeGeneration = 0;
let studySidebarCollapsed = false;

function shortcutKeyId(key) {
  if (key === ' ') return 'space';
  return String(key).toLocaleLowerCase();
}

function shortcutKeyLabel(key) {
  return key === 'Space' ? 'Space' : key === 'ArrowDown' ? '↓' :
    key === 'ArrowRight' ? '→' : key === 'ArrowLeft' ? '←' :
      key === 'Escape' ? 'Esc' : key === 'Enter' ? 'Enter' :
        key.length === 1 ? key.toLocaleUpperCase() : key;
}

function normalizeStudySettings(settings) {
  const savedShortcuts = settings?.shortcuts || {};
  const shortcuts = { ...DEFAULT_STUDY_SETTINGS.shortcuts };
  for (const [action] of STUDY_ACTIONS) {
    const key = savedShortcuts[action];
    if (typeof key === 'string' && key.length > 0) shortcuts[action] = key;
  }
  const uniqueKeys = new Set(Object.values(shortcuts).map(shortcutKeyId));
  if (uniqueKeys.size !== STUDY_ACTIONS.length) {
    Object.assign(shortcuts, DEFAULT_STUDY_SETTINGS.shortcuts);
  }
  return {
    showHintOnBack: typeof settings?.showHintOnBack === 'boolean'
      ? settings.showHintOnBack
      : DEFAULT_STUDY_SETTINGS.showHintOnBack,
    swapKeyValue: settings?.randomizeKeyValue === true ? false : settings?.swapKeyValue === true,
    randomizeKeyValue: settings?.randomizeKeyValue === true,
    fontSizes: {
      front: Number.isFinite(Number(settings?.fontSizes?.front))
        ? Math.min(120, Math.max(18, Number(settings.fontSizes.front)))
        : DEFAULT_STUDY_SETTINGS.fontSizes.front,
      back: Number.isFinite(Number(settings?.fontSizes?.back))
        ? Math.min(72, Math.max(16, Number(settings.fontSizes.back)))
        : DEFAULT_STUDY_SETTINGS.fontSizes.back,
    },
    examDate: /^\d{4}-\d{2}-\d{2}$/.test(settings?.examDate || '') ? settings.examDate : '',
    desiredRetention: normalizeDesiredRetention(settings?.desiredRetention ?? DEFAULT_DESIRED_RETENTION),
    imageSizes: Object.fromEntries(IMAGE_SIZE_VIEWS.map((view) => {
      const saved = settings?.imageSizes?.[view];
      return [view, IMAGE_SIZE_IDS.includes(saved) ? saved : DEFAULT_STUDY_SETTINGS.imageSizes[view]];
    })),
    shortcuts,
  };
}

const IMAGE_SIZE_VIEWS = ['study', 'a4', 'status'];
const IMAGE_SIZE_IDS = ['small', 'medium', 'large', 'fill'];

function fitStudyCardText() {
  const face = document.querySelector('.study-card-face');
  const primaryText = face?.querySelector('.study-key, .study-value');
  if (!face || !primaryText) return;

  const side = primaryText.classList.contains('study-key') ? 'front' : 'back';
  const baseSize = state.studySettings.fontSizes[side];
  const texts = side === 'back'
    ? [primaryText, ...face.querySelectorAll('.study-back-extras .study-hint')]
    : [primaryText];
  texts.forEach((text) => {
    text.style.maxHeight = 'none';
    text.style.overflowY = 'visible';
  });
  let low = 12;
  let high = baseSize;
  let best = low;
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const size = Math.floor((low + high) / 2);
    texts.forEach((text) => { text.style.fontSize = `${size}px`; });
    if (face.scrollHeight <= face.clientHeight + 1) {
      best = size;
      low = size + 1;
    } else {
      high = size - 1;
    }
  }
  const fittedSize = `${Math.max(12, Math.min(baseSize, best))}px`;
  texts.forEach((text) => { text.style.fontSize = fittedSize; });
}

window.addEventListener('resize', fitStudyCardText);
document.addEventListener('load', (event) => {
  if (event.target instanceof HTMLImageElement && event.target.closest('.study-card')) {
    fitStudyCardText();
  }
}, true);

function applyImageSizes() {
  const sizes = state.studySettings.imageSizes;
  IMAGE_SIZE_VIEWS.forEach((view) => {
    document.documentElement.dataset[`img${view[0].toUpperCase()}${view.slice(1)}`] = sizes[view];
  });
}

function imageSizeControlHTML(view) {
  const labels = {
    small: L('작게', 'Small'),
    medium: L('보통', 'Medium'),
    large: L('크게', 'Large'),
    fill: L('꽉 차게', 'Fill'),
  };
  const current = state.studySettings.imageSizes[view];
  return `<label class="image-size-control image-size-control-${view}"><span>${L('이미지 크기', 'Image size')}</span>
    <select data-image-size="${view}">${IMAGE_SIZE_IDS.map((id) => `<option value="${id}"${id === current ? ' selected' : ''}>${labels[id]}</option>`).join('')}</select>
  </label>`;
}

function closeProjectMenus() {
  document.querySelectorAll('.project-menu-list').forEach((list) => {
    list.hidden = true;
    list.previousElementSibling?.setAttribute('aria-expanded', 'false');
  });
}
document.addEventListener('click', (event) => {
  if (!event.target.closest?.('.project-menu')) closeProjectMenus();
  if (!isDatePickerEvent(event)) {
    document.querySelectorAll('[data-date-picker-calendar]').forEach((calendar) => {
      calendar.hidden = true;
      calendar.previousElementSibling?.setAttribute('aria-expanded', 'false');
    });
  }
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') {
    closeProjectMenus();
    document.querySelectorAll('[data-date-picker-calendar]').forEach((calendar) => {
      calendar.hidden = true;
      calendar.previousElementSibling?.setAttribute('aria-expanded', 'false');
    });
  }
});

let studyPanelWidth = 290;
let settingsCollapsed = false;
let shuffleEffectTimer;

function playShuffleEffect(button) {
  clearTimeout(shuffleEffectTimer);
  button.classList.remove('is-shuffling');
  void button.offsetWidth;
  button.classList.add('is-shuffling');
  button.onanimationend = (event) => {
    if (event.animationName !== 'shuffle-button-light') return;
    clearTimeout(shuffleEffectTimer);
    button.classList.remove('is-shuffling');
    button.onanimationend = null;
  };
  shuffleEffectTimer = window.setTimeout(() => {
    button.classList.remove('is-shuffling');
    button.onanimationend = null;
  }, 700);
}

const fileInput = document.createElement('input');
fileInput.type = 'file';
fileInput.accept = '.xlsx,.xlsm,.csv,.txt,.tsv,.apkg,.espr';
fileInput.hidden = true;
document.body.appendChild(fileInput);
fileInput.addEventListener('change', () => {
  const file = fileInput.files && fileInput.files[0];
  fileInput.value = '';
  if (file) handleFile(file);
});

const esc = (value) =>
  String(value ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

const richText = (value) => renderImageText(value, esc);

function hydrateRenderedImages() {
  hydrateImages(document.body, L('이미지 없음', 'Image not found'), state.project?.name);
}

// 인쇄 전에 이미지가 모두 그려질 때까지 잠시 기다립니다.
async function waitForImages(root, timeoutMs = 3000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const pending = [...root.querySelectorAll('img[data-img-id]')].filter((image) => !image.complete || image.dataset.imgState === 'loading');
    if (!pending.length) return;
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
}

let imageHydrationQueued = false;
new MutationObserver(() => {
  if (imageHydrationQueued) return;
  imageHydrationQueued = true;
  setTimeout(() => {
    imageHydrationQueued = false;
    hydrateRenderedImages();
  }, 0);
}).observe(document.body, { childList: true, subtree: true });

const todayLabel = () =>
  new Date().toLocaleDateString(isEnglish() ? 'en-US' : 'ko-KR', { year: 'numeric', month: 'long', day: 'numeric' });

function formatExamDate(value, options = { year: 'numeric', month: 'long', day: 'numeric' }) {
  const date = parseLocalDate(value);
  return date ? date.toLocaleDateString(isEnglish() ? 'en-US' : 'ko-KR', options) : '';
}

function datePickerCalendarHTML(monthValue, minValue) {
  const monthDate = parseLocalDate(`${monthValue}-01`) || new Date();
  const locale = isEnglish() ? 'en-US' : 'ko-KR';
  const weekDays = Array.from({ length: 7 }, (_, index) =>
    new Intl.DateTimeFormat(locale, { weekday: 'short' }).format(new Date(2023, 0, index + 1))
  );
  const cells = calendarMonthDays(monthValue, minValue, localDateValue(new Date()))
    .map((cell) => {
      if (!cell) return '<span class="date-picker-empty" aria-hidden="true"></span>';
      return `<button type="button" class="date-picker-day${cell.today ? ' is-today' : ''}" data-date-value="${cell.value}"${cell.disabled ? ' disabled' : ''}>${new Intl.NumberFormat(locale).format(cell.day)}</button>`;
    }).join('');
  const title = new Intl.DateTimeFormat(locale, { year: 'numeric', month: 'long' }).format(monthDate);
  return `<div class="date-picker-calendar-head">
      <button type="button" data-date-month-step="-1" aria-label="${L('이전 달', 'Previous month')}">‹</button>
      <strong>${title}</strong>
      <button type="button" data-date-month-step="1" aria-label="${L('다음 달', 'Next month')}">›</button>
    </div>
    <div class="date-picker-weekdays" aria-hidden="true">${weekDays.map((day) => `<span>${day}</span>`).join('')}</div>
    <div class="date-picker-days">${cells}</div>`;
}

function datePickerHTML(attribute, value, minValue = '', disabled = false) {
  const selectedDate = parseLocalDate(value);
  const monthValue = localDateValue(selectedDate || new Date()).slice(0, 7);
  const label = value ? formatExamDate(value) : L('날짜 선택', 'Choose a date');
  const pickerId = attribute.includes('setup') ? 'setup-exam-date-calendar' : 'plan-exam-date-calendar';
  return `<div class="date-picker" data-date-picker data-date-min="${minValue}">
    <input type="hidden" ${attribute} value="${esc(value)}">
    <button type="button" class="date-picker-trigger" data-date-picker-trigger aria-haspopup="dialog" aria-controls="${pickerId}" aria-expanded="false"${disabled ? ' disabled' : ''}>
      <span data-date-picker-label>${esc(label)}</span><span aria-hidden="true">▦</span>
    </button>
    <div id="${pickerId}" class="date-picker-popover" data-date-picker-calendar data-date-picker-month="${monthValue}" role="dialog" aria-label="${L('달력에서 시험 날짜 선택', 'Choose exam date from calendar')}" hidden>
      ${datePickerCalendarHTML(monthValue, minValue)}
    </div>
  </div>`;
}

function attachDatePickers() {
  appEl.querySelectorAll('[data-date-picker]').forEach((picker) => {
    attachDatePicker(picker, {
      formatDate: formatExamDate,
      renderCalendar: datePickerCalendarHTML,
      beforeOpen: (calendar) => {
        appEl.querySelectorAll('[data-date-picker-calendar]').forEach((otherCalendar) => {
          if (otherCalendar !== calendar) {
            otherCalendar.hidden = true;
            otherCalendar.previousElementSibling?.setAttribute('aria-expanded', 'false');
          }
        });
      },
    });
  });
}

// key-value 반전이 켜져 있으면 화면에 표시할 때만 앞면과 뒷면을 맞바꿉니다.
function swappedPair(pair) {
  return state.studySettings.swapKeyValue ? { ...pair, key: pair.value, value: pair.key } : pair;
}

function studyShownPair(pair) {
  const session = state.study;
  const shouldSwap = state.studySettings.randomizeKeyValue
    ? session?.randomizedSwapIds?.has(pair.id)
    : state.studySettings.swapKeyValue;
  return shouldSwap ? { ...pair, key: pair.value, value: pair.key } : pair;
}

function randomizeStudyPairDirections(session) {
  session.randomizedSwapIds = state.studySettings.randomizeKeyValue
    ? new Set(session.cards.filter(() => Math.random() < 0.5))
    : new Set();
}

function a4SheetItems(pairs) {
  return pairs.map((rawPair) => {
    const pair = swappedPair(rawPair);
    const hasStudyHint = Boolean(pair.hint);
    return state.expandedA4Hints.has(pair.id)
      ? {
          ...pair,
          hasStudyHint,
        }
      : { ...pair, hint: '', hasStudyHint };
  });
}

function a4PairMatchesQuery(pair, query = state.a4Query.trim().toLocaleLowerCase()) {
  return (
    !query ||
    stripImageTags(pair.key).toLocaleLowerCase().includes(query) ||
    stripImageTags(pair.value).toLocaleLowerCase().includes(query)
  );
}

function a4PriorityStatusText() {
  if (state.a4PriorityError) return state.a4PriorityError;
  if (state.a4PriorityData !== state.fsrsData || state.a4PriorityPairs !== state.pairs) {
    return L('기억률을 계산하는 중…', 'Calculating recall probabilities…');
  }
  return '';
}

function refreshA4Priority() {
  const data = state.fsrsData;
  const pairs = state.pairs;
  if (
    state.a4PriorityData === data &&
    state.a4PriorityPairs === pairs
  ) return Promise.resolve();
  if (a4PriorityRefreshPromise) return a4PriorityRefreshPromise;

  const generation = ++a4PriorityRefreshGeneration;
  const pairIds = new Set(pairs.map((pair) => pair.id));
  const records = (data?.cards || []).filter((record) =>
    record.archivedAt === null &&
    pairIds.has(record.id) &&
    record.reviewLogs.length > 0
  );
  const now = new Date();
  const task = (async () => {
    try {
      const entries = await Promise.all(records.map(async (record) => [
        record.id,
        await getFsrsRetrievability(record.card, now),
      ]));
      if (generation !== a4PriorityRefreshGeneration) return;
      state.a4RecallRates = new Map(entries);
      state.a4PriorityError = '';
    } catch (error) {
      if (generation !== a4PriorityRefreshGeneration) return;
      state.a4RecallRates = new Map();
      state.a4PriorityError = L(
        `기억률을 계산하지 못했어요: ${error instanceof Error ? error.message : String(error)}`,
        `Could not calculate recall probabilities: ${error instanceof Error ? error.message : String(error)}`
      );
      console.error('Could not calculate A4 learning priorities.', error);
    } finally {
      if (generation === a4PriorityRefreshGeneration) {
        state.a4PriorityData = data;
        state.a4PriorityPairs = pairs;
        a4PriorityRefreshPromise = null;
        if (state.page === 'a4' && state.a4Sort === 'priority') renderA4View();
      }
    }
  })();
  a4PriorityRefreshPromise = task;
  return task;
}

function a4Pages() {
  const pairsById = new Map(state.pairs.map((pair) => [pair.id, pair]));
  const items = state.a4Order
    .map((id) => pairsById.get(id))
    .filter(Boolean)
    .filter((pair) => a4PairMatchesQuery(pair));
  if (state.a4Sort === 'priority') {
    const reviewed = [];
    const unreviewed = [];
    items.forEach((pair) => {
      const record = fsrsRecordForCard(pair.id);
      if (record?.reviewLogs.length) {
        reviewed.push({
          ...pair,
          recallProbability: state.a4RecallRates.get(pair.id) ?? 1,
        });
      } else {
        unreviewed.push(pair);
      }
    });
    const priorityItems = sortByRecallProbability(reviewed);
    return buildPriorityPages(
      a4SheetItems(priorityItems),
      a4SheetItems(unreviewed)
    );
  }
  return buildPages(a4SheetItems(items)).map((rows) => ({ rows }));
}

function visibleA4HintPairs() {
  const visibleIds = new Set(
    a4Pages().flatMap(({ rows }) =>
      rows.filter((row) => row.hasStudyHint).map((row) => row.id)
    )
  );
  return state.pairs.filter((pair) => visibleIds.has(pair.id));
}

function a4ItemCounts() {
  return { totalCount: state.pairs.length };
}

function syncA4HintControls(button) {
  const pairs = visibleA4HintPairs();
  const allExpanded = pairs.length > 0 && pairs.every((pair) => state.expandedA4Hints.has(pair.id));
  button.textContent = allExpanded ? L('전체 힌트 숨기기', 'Hide all hints') : L('전체 힌트 보기', 'Show all hints');
  button.setAttribute('aria-pressed', String(allExpanded));
  button.dataset.studyHint = allExpanded
    ? L('현재 표시 중인 key의 힌트를 모두 접습니다.', 'Collapse the hints for all visible keys.')
    : L('현재 표시 중인 key의 힌트를 모두 펼칩니다.', 'Expand the hints for all visible keys.');
  button.disabled = !pairs.length;
}

function a4SheetTitle() {
  return state.a4Title ?? state.title ?? L('학습 목록', 'Study set');
}

function a4SheetInfo() {
  return state.a4Info ?? L(`${todayLabel()} · 총 ${state.a4Order.length}개`, `${todayLabel()} · ${state.a4Order.length} items`);
}

function isA4ValueVisible(id) {
  return state.showAllA4Values
    ? !state.hiddenA4Values.has(id)
    : state.revealedA4Values.has(id);
}

function shuffleA4Order(pairs) {
  return shuffle(pairs.map((pair) => pair.id));
}

/* ---------------- 저장 / 불러오기 ---------------- */

async function save() {
  if (!state.project) return false;
  try {
    state.project.pairs = state.pairs;
    state.project.title = state.title;
    state.project.a4Title = state.a4Title;
    state.project.a4Info = state.a4Info;
    await saveProject(state.project);
    state.projects = state.projects.map((project) =>
      project.id === state.project.id ? state.project : project
    );
    if ((state.error.startsWith('학습 목록을 저장하지 못했어요:') || state.error.startsWith('Could not save the study set:'))) {
      state.error = '';
      const errorEl = document.getElementById('a4StorageError');
      if (errorEl) {
        errorEl.textContent = '';
        errorEl.hidden = true;
      }
    }
    return true;
  } catch (error) {
    state.error = L(`학습 목록을 저장하지 못했어요: ${error instanceof Error ? error.message : String(error)}`, `Could not save the study set: ${error instanceof Error ? error.message : String(error)}`);
    const errorEl = document.getElementById('a4StorageError');
    if (errorEl) {
      errorEl.textContent = state.error;
      errorEl.hidden = false;
    }
    return false;
  }
}

async function load() {
  try {
    state.projects = await listProjects();
    const savedProjectOrder = await getProjectOrder();
    const projectById = new Map(state.projects.map((project) => [project.id, project]));
    state.projects = [
      ...savedProjectOrder.map((id) => projectById.get(id)).filter(Boolean),
      ...state.projects.filter((project) => !savedProjectOrder.includes(project.id)),
    ];
    state.projectOrder = state.projects.map((project) => project.id);
    const savedLanguage = await getLanguage();
    if (savedLanguage) setLanguageValue(savedLanguage);
    else setLanguageValue(Intl.DateTimeFormat().resolvedOptions().timeZone === 'Asia/Seoul' ? 'ko' : 'en');
    state.studySettings = normalizeStudySettings(await getStudySettings());
    const savedDirectory = await getStoredWorkbookDirectory();
    state.directoryHandle = savedDirectory;
    const directory = await getWorkbookDirectory();
    state.directorySaved = Boolean(savedDirectory);
    state.directoryConnected = Boolean(directory);
    state.directoryName = savedDirectory?.name || directory?.name || '';
    state.page = 'workspace';
  } catch (error) {
    state.error = L(`프로젝트를 불러오지 못했어요: ${error instanceof Error ? error.message : String(error)}`, `Could not load projects: ${error instanceof Error ? error.message : String(error)}`);
  }

  await loadGenerationPrompt();
}

async function loadGenerationPrompt() {
  state.generationPrompt = '';
  state.generationPromptError = '';
  const requestedLanguage = language;
  try {
    const response = await fetch(new URL(PROMPT_FILE_URLS[requestedLanguage], import.meta.url));
    if (!response.ok) {
      throw new Error(L(`프롬프트 파일 요청에 실패했어요. (${response.status})`, `Could not load the prompt file. (${response.status})`));
    }
    const text = await response.text();
    if (requestedLanguage === language) state.generationPrompt = text;
  } catch (error) {
    if (requestedLanguage === language) state.generationPromptError = error instanceof Error ? error.message : String(error);
  }
}

function setLanguageValue(value) {
  language = value === 'ko' ? 'ko' : 'en';
  state.language = language;
  document.documentElement.lang = language;
  document.title = 'ESPER';
}

// 언어 설정만 바꿉니다. 프로젝트, 엑셀 파일, FSRS 기록은 건드리지 않습니다.
async function changeLanguage(value) {
  const next = value === 'ko' ? 'ko' : 'en';
  const wasPending = state.languageChoicePending;
  if (next === language && !wasPending) return;
  try {
    await saveLanguage(next);
  } catch (error) {
    state.error = `${L('언어 설정을 저장하지 못했어요', 'Could not save the language setting')}: ${error instanceof Error ? error.message : String(error)}`;
    render();
    return;
  }
  const previousLanguage = language;
  setLanguageValue(next);
  state.languageChoicePending = false;
  state.error = '';
  state.notice = '';
  mountSupportWidget();
  if (previousLanguage !== language) loadGenerationPrompt().then(() => {
    if (['workspace', 'home'].includes(state.page)) render();
  });
  render();
}

/* ---------------- 파일 읽기 ---------------- */

async function handleFile(file) {
  if (/\.apkg$/i.test(file.name)) return startAnkiImport(file);
  if (/\.espr$/i.test(file.name)) return importEsprFile(file);
  return importWorkbookFile(file);
}

const TUTORIAL_DECK_FILES = {
  ko: '튜토리얼 영단어 덱.xlsx',
  en: 'Tutorial Vocab Deck.xlsx',
};

async function startTutorialDeck() {
  const fileName = TUTORIAL_DECK_FILES[isEnglish() ? 'en' : 'ko'];
  try {
    const response = await fetch(`assets/tutorial/${encodeURIComponent(fileName)}`);
    if (!response.ok) throw new Error(String(response.status));
    const blob = await response.blob();
    await importWorkbookFile(new File([blob], fileName, { type: blob.type }));
  } catch {
    state.error = L('튜토리얼 덱을 불러오지 못했어요.', 'Could not load the tutorial deck.');
    render();
  }
}

async function importEsprFile(file) {
  state.busy = true;
  state.processingMessage = L('ESPER 공유 파일을 읽는 중…', 'Reading the ESPER share file…');
  render();
  try {
    const shared = await readEsprFile(file);
    state.processingMessage = L('공유 파일의 이미지를 저장하는 중…', 'Saving images from the share file…');
    renderProcessingIndicator();
    return await importWorkbookFile(shared.workbook, shared.images);
  } catch {
    state.error = L('.espr 파일을 읽지 못했어요. 올바른 ESPER 공유 파일인지 확인해 주세요.', 'Could not read the .espr file. Check that it is a valid ESPER share file.');
    render();
  } finally {
    state.busy = false;
    state.processingMessage = '';
    render();
  }
}

async function shareProject(id) {
  const project = state.projects.find((item) => item.id === id);
  if (!project) return;
  processingProjectIds.add(id);
  state.error = '';
  state.notice = '';
  state.processingMessage = L('공유 파일을 만드는 중…', 'Creating the share file…');
  render();
  try {
    const stored = await readProjectWorkbook(project.name, project.fileName);
    const { file } = await normalizeStudyWorkbook(stored);
    const safeName = project.name.replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').trim() || 'project';
    const workbook = new File([await file.arrayBuffer()], `${safeName}.xlsx`, { type: file.type });
    const { blob, imageCount, missingImages } = await createEsprBlob(
      workbook,
      (imageId) => getImageBlob(imageId, project.name)
    );
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${safeName}${ESPR_EXTENSION}`;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
    state.notice = L(
      `${safeName}${ESPR_EXTENSION} 파일을 만들었어요 (이미지 ${imageCount}개).${missingImages ? ` 찾지 못한 이미지 ${missingImages}개는 제외했어요.` : ''}`,
      `Created ${safeName}${ESPR_EXTENSION} (${imageCount} images).${missingImages ? ` ${missingImages} missing images were skipped.` : ''}`
    );
  } catch (error) {
    state.error = `${L('프로젝트를 공유 파일로 만들지 못했어요', 'Could not create the share file')}: ${error instanceof Error ? error.message : String(error)}`;
  } finally {
    processingProjectIds.delete(id);
    state.processingMessage = '';
    render();
  }
}

async function importWorkbookFile(file, images = []) {
  state.busy = true;
  state.processingMessage = L('학습 파일을 저장하는 중…', 'Saving the study file…');
  state.error = '';
  state.notice = '';
  render();
  try {
    const directory = await getWorkbookDirectory();
    if (!directory) {
      throw new Error(L('파일을 저장할 폴더를 먼저 선택해 주세요.', 'Please choose a folder to save your files in first.'));
    }
    const normalized = await normalizeStudyWorkbook(file);
    file = normalized.file;
    const items = await parseStudyFile(file);
    if (!items.length) {
      throw new Error(L('key와 value가 있는 항목을 찾지 못했어요. A열 key와 B열 value를 확인해 주세요.', 'No items with both a key and value were found. Check column A for key and column B for value.'));
    }
    const name = projectNameFromFileName(file.name);
    if (projectNameExists(state.projects.map((project) => project.name), name)) {
      throw new Error(`"${name}" 이름의 프로젝트가 이미 있어요. 중복되지 않는 파일 이름을 사용해 주세요.`);
    }
    const project = {
      id: crypto.randomUUID(),
      name,
      fileName: file.name,
      sourceName: file.name,
      createdAt: Date.now(),
      pairs: items,
      title: name,
      a4Title: null,
      a4Info: null,
      formatVersion: 12,
      cartridgeColorIndex: state.projects.length % 4,
    };
    await createProjectDirectory(project.name);
    let fsrsData;
    try {
      await writeProjectWorkbook(project.name, project.fileName, file);
      for (let index = 0; index < images.length; index += 1) {
        state.processingMessage = L(
          `이미지 저장 중 ${index + 1}/${images.length}`,
          `Saving images ${index + 1}/${images.length}`
        );
        renderProcessingIndicator();
        await saveImageFile(images[index].id, images[index].blob, project.name);
      }
      state.processingMessage = L('학습 데이터를 저장하는 중…', 'Saving study data…');
      renderProcessingIndicator();
      const synchronized = await synchronizeFsrsProject(project, items);
      fsrsData = synchronized.data;
      project.pairs = synchronized.pairs;
      await saveProject(project);
    } catch (error) {
      try {
        await removeProjectDirectory(project.name);
      } catch (cleanupError) {
        throw new Error(
          `${error instanceof Error ? error.message : String(error)} ${L('프로젝트 폴더를 정리하지 못했습니다.', 'The project folder could not be cleaned up.')}: ${cleanupError instanceof Error ? cleanupError.message : String(cleanupError)}`
        );
      }
      throw error;
    }
    releaseProjectImages(project.name);
    state.projects = [project, ...state.projects];
    state.projectOrder = state.projects.map((item) => item.id);
    try {
      await saveProjectOrder(state.projectOrder);
    } catch (error) {
      state.notice = `${L('프로젝트는 추가했지만 목록 순서를 저장하지 못했어요', 'The project was added, but its list position could not be saved')}: ${error instanceof Error ? error.message : String(error)}`;
    }
    state.project = project;
    state.fsrsData = fsrsData;
    state.pairs = project.pairs;
    state.sourceFile = file;
    state.revealedA4Values = new Set();
    state.hiddenA4Values = new Set();
    state.expandedA4Hints = new Set();
    state.showAllA4Values = true;
    state.a4Title = null;
    state.a4Info = null;
    state.a4Order = shuffleA4Order(items);
    state.a4Query = '';
    state.sourceName = file.name;
    state.title = project.title;
    startStudy();
    state.page = 'project-setup';
  } catch (e) {
    state.error = e && e.message ? e.message : L('파일을 읽는 중 문제가 생겼어요.', 'Something went wrong while reading the file.');
  } finally {
    state.busy = false;
    state.processingMessage = '';
    render();
  }
}

/* ---------------- Anki .apkg 가져오기 ---------------- */

function ankiErrorMessage(error) {
  switch (error instanceof AnkiImportError ? error.code : '') {
    case 'new-format':
      return L("Anki에서 'Support older Anki versions'로 다시 내보내 주세요.", "Please re-export from Anki with 'Support older Anki versions'.");
    case 'no-collection':
      return L('덱 데이터(collection.anki21 또는 collection.anki2)를 찾지 못했어요.', 'Could not find the deck data (collection.anki21 or collection.anki2).');
    case 'no-notes':
      return L('가져올 노트를 찾지 못했어요.', 'No notes were found to import.');
    case 'sql-unavailable':
      return L('Anki 데이터베이스를 읽는 도구를 불러오지 못했어요.', 'Could not load the tool that reads the Anki database.');
    case 'bad-file':
    case 'bad-database':
      return L('.apkg 파일을 읽지 못했어요. 올바른 Anki 덱인지 확인해 주세요.', 'Could not read the .apkg file. Check that it is a valid Anki deck.');
    default:
      return error instanceof Error && error.message ? error.message : L('파일을 읽는 중 문제가 생겼어요.', 'Something went wrong while reading the file.');
  }
}

function defaultAnkiMapping(deck, noteTypeId) {
  const noteType = deck.noteTypes.find((type) => type.id === noteTypeId);
  const fieldCount = noteType.fields.length;
  const imageStats = countFieldImages(deck, noteTypeId);
  const imageFields = imageStats.map((stat, index) => (stat.found && index >= 2 ? index : -1)).filter((index) => index >= 0);
  return {
    key: fieldCount > 0 ? [0] : [],
    value: fieldCount > 1 ? [1] : [],
    hint: imageFields.length ? imageFields : fieldCount > 2 ? [2] : [],
  };
}

async function startAnkiImport(file) {
  state.busy = true;
  state.processingMessage = L('Anki 덱을 읽는 중…', 'Reading the Anki deck…');
  state.error = '';
  state.notice = '';
  render();
  try {
    if (!(await getWorkbookDirectory())) {
      throw new Error(L('파일을 저장할 폴더를 먼저 선택해 주세요.', 'Please choose a folder to save your files in first.'));
    }
    await new Promise((resolve) => window.setTimeout(resolve, 0));
    const deck = await loadAnkiDeck(file);
    const noteTypeId = deck.noteTypes[0].id;
    state.anki = {
      fileName: file.name,
      deck,
      noteTypeId,
      mapping: defaultAnkiMapping(deck, noteTypeId),
      separator: 'newline',
      previewIndex: 0,
      previewFlipped: false,
      previewUrls: new Set(),
      busy: false,
      progress: '',
      error: '',
      focus: '',
    };
  } catch (error) {
    state.error = ankiErrorMessage(error);
  } finally {
    state.busy = false;
    state.processingMessage = '';
    render();
  }
}

function ankiImportHTML() {
  const anki = state.anki;
  const noteType = anki.deck.noteTypes.find((type) => type.id === anki.noteTypeId);
  const imageStats = countFieldImages(anki.deck, anki.noteTypeId);
  const preview = previewAnkiCards(anki.deck, anki.noteTypeId, anki.mapping, anki.separator, anki.previewIndex);
  const areas = [
    ['key', 'Key', L('학습 카드 앞면', 'Front of the card')],
    ['value', 'Value', L('학습 카드 뒷면', 'Back of the card')],
    ['hint', 'Hint', L('추가 정보 또는 단서', 'Extra info or clues')],
  ];
  const disabled = anki.busy ? ' disabled' : '';
  const renderParts = (parts) => parts.map((part) => (
    part.type === 'image'
      ? `<img class="cell-img" data-anki-media="${esc(part.name)}" src="${TRANSPARENT_PIXEL}" alt="${L('Anki 이미지 미리보기', 'Anki image preview')}">`
      : esc(part.value)
  )).join('');
  const previewFace = preview.card
    ? anki.previewFlipped
      ? `<div class="anki-preview-face is-back">
          <span class="anki-preview-label">Value</span>
          <div class="anki-preview-content">${renderParts(preview.card.value)}</div>
          ${preview.card.hint.length ? `<div class="anki-preview-hint"><span class="anki-preview-label">Hint</span><div class="anki-preview-content">${renderParts(preview.card.hint)}</div></div>` : ''}
          <button type="button" class="btn" data-anki-flip>${L('앞면 보기', 'Show front')}</button>
        </div>`
      : `<div class="anki-preview-face">
          <span class="anki-preview-label">Key</span>
          <div class="anki-preview-content">${renderParts(preview.card.key)}</div>
          <button type="button" class="btn" data-anki-flip>${L('뒷면 보기', 'Show back')}</button>
        </div>`
    : `<p class="anki-preview-empty">${L('미리 볼 수 있는 카드가 없어요. Key와 Value 필드 배치를 확인해 주세요.', 'No cards to preview. Check the Key and Value field mapping.')}</p>`;
  const badge = (index) => {
    const { found, missing } = imageStats[index];
    return `${found ? ` <span class="anki-badge">🖼 ${found}</span>` : ''}${missing ? ` <span class="anki-badge is-warn">⚠ ${missing}</span>` : ''}`;
  };
  return `
    <div class="anki-overlay">
      <div class="anki-dialog" role="dialog" aria-modal="true" aria-labelledby="ankiTitle">
        <h2 id="ankiTitle">${L('Anki 덱 가져오기', 'Import Anki deck')}</h2>
        <p class="anki-sub">${esc(anki.fileName)} · ${L('오디오는 가져오지 않아요.', 'Audio is not imported.')}</p>
        <div class="anki-row">
          <label>${L('노트 타입', 'Note type')}
            <select data-anki-type${disabled}>
              ${anki.deck.noteTypes.map((type) => `<option value="${esc(type.id)}"${type.id === anki.noteTypeId ? ' selected' : ''}>${esc(type.name)} (${type.noteCount})</option>`).join('')}
            </select>
          </label>
          <label>${L('여러 필드를 합칠 때', 'When combining fields')}
            <select data-anki-separator${disabled}>
              <option value="newline"${anki.separator === 'newline' ? ' selected' : ''}>${L('줄바꿈', 'New line')}</option>
              <option value="slash"${anki.separator === 'slash' ? ' selected' : ''}>/</option>
              <option value="space"${anki.separator === 'space' ? ' selected' : ''}>${L('공백', 'Space')}</option>
            </select>
          </label>
        </div>
        <div class="anki-areas">
          ${areas.map(([area, title, description]) => `
            <fieldset class="anki-area"${disabled}>
              <legend>${title}</legend>
              <small>${description}</small>
              ${noteType.fields.map((field, index) => `
                <label><input type="checkbox" data-anki-area="${area}" value="${index}"${anki.mapping[area].includes(index) ? ' checked' : ''}> <span>${esc(field)}${badge(index)}</span></label>`).join('')}
            </fieldset>`).join('')}
        </div>
        <p class="anki-count">${L(`가져올 항목 ${preview.total}개`, `${preview.total} items to import`)}</p>
        <section class="anki-card-preview" aria-label="${L('학습 카드 미리보기', 'Flashcard preview')}">
          ${previewFace}
          <div class="anki-preview-navigation">
            <button type="button" class="btn" data-anki-nav="-1"${disabled}${anki.previewIndex <= 0 ? ' disabled' : ''} aria-label="${L('이전 카드', 'Previous card')}">←</button>
            <span aria-live="polite" aria-atomic="true">${preview.total ? `${anki.previewIndex + 1} / ${preview.total}` : '0 / 0'}</span>
            <button type="button" class="btn" data-anki-nav="1"${disabled}${anki.previewIndex >= preview.total - 1 ? ' disabled' : ''} aria-label="${L('다음 카드', 'Next card')}">→</button>
          </div>
        </section>
        <p class="anki-message${anki.error ? ' is-error' : ''}" role="status" data-anki-message>${esc(anki.error || anki.progress)}</p>
        <div class="anki-actions">
          <button type="button" class="btn" data-anki-cancel${disabled}>${L('취소', 'Cancel')}</button>
          <button type="button" class="btn btn-primary" data-anki-confirm${disabled}>${L('가져오기', 'Import')}</button>
        </div>
      </div>
    </div>`;
}

function releaseAnkiPreviewUrls(anki) {
  for (const url of anki.previewUrls) URL.revokeObjectURL(url);
  anki.previewUrls.clear();
}

function hydrateAnkiPreviewImages(anki, overlay) {
  overlay.querySelectorAll('img[data-anki-media]').forEach((image) => {
    const name = image.dataset.ankiMedia;
    anki.deck.readMedia(name).then((blob) => {
      if (!image.isConnected || state.anki !== anki) return;
      const url = URL.createObjectURL(blob);
      anki.previewUrls.add(url);
      image.src = url;
    }).catch((error) => {
      if (!image.isConnected || state.anki !== anki) return;
      console.error(`Could not load Anki preview image "${name}".`, error);
      const placeholder = document.createElement('span');
      placeholder.className = 'cell-img-missing';
      placeholder.textContent = L('이미지를 불러오지 못했어요.', 'Could not load image.');
      image.replaceWith(placeholder);
    });
  });
}

function attachAnkiDialog() {
  const anki = state.anki;
  const overlay = appEl.querySelector('.anki-overlay');
  if (!anki || !overlay) return;
  hydrateAnkiPreviewImages(anki, overlay);
  const refresh = (focus) => {
    anki.focus = focus;
    anki.error = '';
    render();
  };
  overlay.querySelector('[data-anki-type]').addEventListener('change', (event) => {
    anki.noteTypeId = event.target.value;
    anki.mapping = defaultAnkiMapping(anki.deck, anki.noteTypeId);
    anki.previewIndex = 0;
    anki.previewFlipped = false;
    releaseAnkiPreviewUrls(anki);
    refresh('[data-anki-type]');
  });
  overlay.querySelector('[data-anki-separator]').addEventListener('change', (event) => {
    anki.separator = event.target.value;
    anki.previewIndex = 0;
    anki.previewFlipped = false;
    releaseAnkiPreviewUrls(anki);
    refresh('[data-anki-separator]');
  });
  overlay.querySelectorAll('[data-anki-area]').forEach((box) => {
    box.addEventListener('change', () => {
      const area = box.dataset.ankiArea;
      const index = Number(box.value);
      const next = new Set(anki.mapping[area]);
      if (box.checked) next.add(index);
      else next.delete(index);
      anki.mapping[area] = [...next].sort((a, b) => a - b);
      anki.previewIndex = 0;
      anki.previewFlipped = false;
      releaseAnkiPreviewUrls(anki);
      refresh(`[data-anki-area="${area}"][value="${index}"]`);
    });
  });
  overlay.querySelector('[data-anki-flip]')?.addEventListener('click', () => {
    anki.previewFlipped = !anki.previewFlipped;
    refresh('[data-anki-flip]');
  });
  overlay.querySelectorAll('[data-anki-nav]').forEach((button) => {
    button.addEventListener('click', () => {
      anki.previewIndex += Number(button.dataset.ankiNav);
      anki.previewFlipped = false;
      releaseAnkiPreviewUrls(anki);
      refresh(`[data-anki-nav="${button.dataset.ankiNav}"]`);
    });
  });
  overlay.querySelector('[data-anki-cancel]').addEventListener('click', () => {
    releaseAnkiPreviewUrls(anki);
    state.anki = null;
    render();
  });
  overlay.querySelector('[data-anki-confirm]').addEventListener('click', confirmAnkiImport);
  overlay.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !anki.busy) {
      releaseAnkiPreviewUrls(anki);
      state.anki = null;
      render();
    }
  });
  const target = (anki.focus && overlay.querySelector(`${anki.focus}:not(:disabled)`)) || overlay.querySelector('[data-anki-confirm]');
  anki.focus = '';
  target?.focus();
}

async function confirmAnkiImport() {
  const anki = state.anki;
  if (!anki || anki.busy) return;
  if (!anki.mapping.key.length || !anki.mapping.value.length) {
    anki.error = L('Key와 Value에 필드를 하나 이상 배치해 주세요.', 'Place at least one field in Key and in Value.');
    render();
    return;
  }
  anki.busy = true;
  anki.error = '';
  anki.progress = L('가져오는 중…', 'Importing…');
  render();
  try {
    const images = new Map();
    const result = await buildAnkiRows(anki.deck, anki.noteTypeId, anki.mapping, {
      separator: anki.separator,
      saveImage: async (blob) => {
        const processed = await processImage(blob);
        images.set(processed.id, processed.blob);
        return processed.id;
      },
      onProgress: (done, total) => {
        anki.progress = L(`이미지 처리 중 ${done}/${total}`, `Processing images ${done}/${total}`);
        const message = appEl.querySelector('[data-anki-message]');
        if (message) message.textContent = anki.progress;
        renderProcessingIndicator();
      },
    });
    if (!result.rows.length) {
      throw new Error(L('Key와 Value가 모두 있는 노트가 없어요. 필드 배치를 확인해 주세요.', 'No notes have both a key and a value. Check the field placement.'));
    }
    anki.progress = L('엑셀 학습 파일을 만드는 중…', 'Creating the workbook…');
    renderProcessingIndicator();
    const workbook = await createStudyWorkbook(result.rows, ['key', 'value', 'hint']);
    const baseName = anki.fileName.replace(/\.apkg$/i, '') || 'anki';
    const file = new File([workbook], `${baseName}.xlsx`, {
      type: workbook.type,
      lastModified: Date.now(),
    });
    releaseAnkiPreviewUrls(anki);
    state.anki = null;
    await importWorkbookFile(file, [...images].map(([id, blob]) => ({ id, blob })));
    if (!state.error && result.failedImages) {
      state.notice = L(`이미지 ${result.failedImages}개는 읽지 못해 건너뛰었어요.`, `${result.failedImages} images could not be read and were skipped.`);
      render();
    }
  } catch (error) {
    anki.busy = false;
    anki.progress = '';
    anki.error = error instanceof Error ? error.message : String(error);
    render();
  }
}

async function openProject(id, destination = 'study') {
  state.error = '';
  state.notice = '';
  state.directoryAccessIssue = '';
  if (!state.directoryConnected) {
    state.directoryAccessIssue = L('현재 폴더에 접근할 수 없습니다. 저장 위치를 다시 선택해 주세요.', 'Cannot access the current folder. Please choose the save location again.');
    const allowed = await ensureWorkbookDirectoryAccess();
    if (!allowed) return;
    state.directoryAccessIssue = '';
    state.notice = L('저장 폴더 권한을 허용했어요. 프로젝트를 다시 열어 주세요.', 'Folder access granted. Please open the project again.');
    render();
    return;
  }
  processingProjectIds.add(id);
  state.processingMessage = L('프로젝트를 여는 중…', 'Opening project…');
  render();
  try {
    const project = await getProject(id);
    if (!project) throw new Error(L('프로젝트를 찾을 수 없어요.', 'Project not found.'));
    const storedFile = await readProjectWorkbook(project.name, project.fileName);
    const normalized = await normalizeStudyWorkbook(storedFile);
    const file = normalized.file;
    if (normalized.converted) {
      await writeProjectWorkbook(project.name, project.fileName, file);
    }
    const synchronized = await synchronizeFsrsProject(
      project,
      project.pairs.map(normalizeStudyPair)
    );
    project.pairs = synchronized.pairs;
    await saveProject(project);
    state.project = project;
    state.sourceFile = file;
    state.fsrsData = synchronized.data;
    state.pendingReview = null;
    state.pairs = synchronized.pairs.map(normalizeStudyPair);
    state.sourceName = project.sourceName || project.fileName;
    state.title = project.title || project.name;
    state.a4Title = typeof project.a4Title === 'string' ? project.a4Title : null;
    state.a4Info = typeof project.a4Info === 'string' ? project.a4Info : null;
    state.revealedA4Values = new Set();
    state.hiddenA4Values = new Set();
    state.expandedA4Hints = new Set();
    state.showAllA4Values = true;
    state.a4Order = shuffleA4Order(state.pairs);
    state.a4Query = '';
    startStudy();
    if (destination !== 'study') await routeToProjectDestination(destination);
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    state.directoryAccessIssue = detail.includes('FSRS')
      ? ''
      : L('현재 폴더 또는 프로젝트 파일에 접근할 수 없습니다. 저장 위치를 다시 선택해 주세요.', 'Cannot access the folder or project file. Please choose the save location again.');
    state.error = `${L('프로젝트를 열지 못했어요', 'Could not open the project')}: ${detail}`;
    state.page = 'workspace';
  } finally {
    processingProjectIds.delete(id);
    state.processingMessage = '';
  }
  render();
}

async function recoverProjectsFromFolder(directory) {
  const knownIds = new Set(state.projects.map((project) => project.id));
  const knownNames = new Set(state.projects.map((project) => projectNameKey(project.name)));
  const recovery = await recoverWorkbookProjects(
    directory,
    (projectName, fileName) => readFsrsDataFile(projectName, fileName)
  );
  const recoveredProjects = recovery.projects.filter((project) => {
    const nameKey = projectNameKey(project.name);
    if (knownIds.has(project.id) || knownNames.has(nameKey)) return false;
    knownNames.add(nameKey);
    return true;
  });
  for (const project of recoveredProjects) {
    await saveProject(project);
  }
  if (recoveredProjects.length) {
    const projects = await listProjects();
    const order = await getProjectOrder();
    const projectById = new Map(projects.map((project) => [project.id, project]));
    state.projects = [
      ...order.map((id) => projectById.get(id)).filter(Boolean),
      ...projects.filter((project) => !order.includes(project.id)),
    ];
    state.projectOrder = state.projects.map((project) => project.id);
    await saveProjectOrder(state.projectOrder);
  }
  return { count: recoveredProjects.length, errors: recovery.errors };
}

async function prepareWorkbookDirectory(directory) {
  const recovered = await recoverProjectsFromFolder(directory);
  if (recovered.errors.length) {
    throw new Error(recovered.errors.join(' · '));
  }
  return { recoveredCount: recovered.count };
}

async function connectWorkbookDirectory() {
  try {
    if (typeof window.showDirectoryPicker !== 'function') {
      state.error = L('이 브라우저에서는 폴더 선택 기능을 사용할 수 없어요. 최신 Chrome에서 다시 시도해 주세요.', 'This browser cannot choose folders. Please try again in the latest Chrome.');
      render();
      return;
    }
    const directory = await window.showDirectoryPicker({
      id: 'esper-excel-projects',
      mode: 'readwrite',
    });
    await saveWorkbookDirectory(directory);
    state.directoryHandle = directory;
    state.directorySaved = true;
    state.directoryConnected = true;
    state.directoryName = directory.name;
    state.directoryAccessIssue = '';
    state.error = '';
    state.notice = L(`${directory.name} 폴더를 저장 위치로 연결했어요.`, `Connected the ${directory.name} folder as your save location.`);
    let prepared;
    try {
      prepared = await prepareWorkbookDirectory(directory);
    } catch (error) {
      state.error = L(
        `폴더는 연결했지만 프로젝트를 불러오지 못했어요: ${error instanceof Error ? error.message : String(error)}`,
        `The folder was connected, but projects could not be loaded: ${error instanceof Error ? error.message : String(error)}`
      );
      render();
      return;
    }
    if (prepared.recoveredCount) {
      state.page = 'workspace';
      state.notice = L(
        `저장 위치에서 기존 프로젝트 ${prepared.recoveredCount}개를 불러왔어요.`,
        `Loaded ${prepared.recoveredCount} existing project(s) from the save location.`
      );
    }
    render();
  } catch (error) {
    if (error instanceof DOMException && error.name === 'AbortError') return;
    state.error = `${L('저장 폴더를 연결하지 못했어요', 'Could not connect the save folder')}: ${error instanceof Error ? error.message : String(error)}`;
    render();
  }
}

async function ensureWorkbookDirectoryAccess() {
  try {
    const directory = state.directoryHandle;
    if (!directory) {
      state.directoryConnected = false;
      state.error = L('먼저 파일을 저장할 폴더를 선택해 주세요.', 'Please choose a folder to save your files in first.');
      render();
      return false;
    }
    if (state.directoryConnected) return true;
    const permission = await directory.requestPermission({ mode: 'readwrite' });
    if (permission !== 'granted') {
      state.directoryConnected = false;
      state.directorySaved = true;
      state.directoryName = directory.name;
      state.error = L('저장 폴더 권한을 허용해 주세요.', 'Please allow access to the save folder.');
      render();
      return false;
    }
    state.directoryConnected = true;
    state.directorySaved = true;
    state.directoryName = directory.name;
    return true;
  } catch (error) {
    state.error = `${L('저장 폴더 권한을 확인하지 못했어요', 'Could not check folder access')}: ${error instanceof Error ? error.message : String(error)}`;
    render();
    return false;
  }
}

const EDIT_COLUMN_LABELS = ['key', 'value', 'hint'];
const EDIT_COLUMN_LABELS_EN = ['Key', 'Value', 'Hint'];
const DEFAULT_EDIT_COLUMN_WIDTHS = [0.28, 0.32, 0.4];
const MIN_EDIT_COLUMN_WIDTH = 100;

function loadEditColumnWidths() {
  const saved = localStorage.getItem(EDIT_COLUMN_WIDTHS_KEY);
  if (saved === null) return [...DEFAULT_EDIT_COLUMN_WIDTHS];
  const widths = JSON.parse(saved);
  if (
    !Array.isArray(widths) ||
    widths.length !== 3 ||
    widths.some((width) => !Number.isFinite(width) || width <= 0) ||
    Math.abs(widths.reduce((sum, width) => sum + width, 0) - 1) > 0.001
  ) {
    throw new Error(L('저장된 열 너비 설정이 올바르지 않습니다.', 'Saved column width settings are invalid.'));
  }
  return widths;
}

function saveEditColumnWidths(widths) {
  const normalized = normalizeColumnWidths(widths);
  localStorage.setItem(EDIT_COLUMN_WIDTHS_KEY, JSON.stringify(normalized));
  return normalized;
}

function applyEditColumnWidths(table, ratios) {
  const wrapper = table.closest('.edit-sheet-wrap');
  const tableWidth = Math.max(wrapper.clientWidth, 960);
  const contentWidth = tableWidth - 56;
  table.style.width = `${tableWidth}px`;
  table.querySelector('[data-edit-row-number-column]').style.width = '56px';
  table.querySelectorAll('[data-edit-column]').forEach((column, index) => {
    column.style.width = `${contentWidth * ratios[index]}px`;
  });
}

async function editProject(id) {
  if (processingProjectIds.has(id)) return;
  state.error = '';
  state.notice = '';
  if (!state.directoryConnected) {
    const allowed = await ensureWorkbookDirectoryAccess();
    if (!allowed) return;
    state.notice = L('저장 폴더 권한을 허용했어요. 프로젝트 편집을 다시 눌러 주세요.', 'Folder access granted. Please click Edit project again.');
    render();
    return;
  }
  processingProjectIds.add(id);
  state.processingMessage = L('프로젝트를 여는 중…', 'Opening project…');
  render();
  try {
    const project = await getProject(id);
    if (!project) throw new Error(L('프로젝트를 찾을 수 없어요.', 'Project not found.'));
    const storedFile = await readProjectWorkbook(project.name, project.fileName);
    const normalized = await normalizeStudyWorkbook(storedFile);
    const file = normalized.file;
    if (normalized.converted) {
      await writeProjectWorkbook(project.name, project.fileName, file);
    }
    const synchronized = await synchronizeFsrsProject(
      project,
      project.pairs.map(normalizeStudyPair)
    );
    project.pairs = synchronized.pairs;
    await saveProject(project);
    const rows = await readSheetGrid(file);
    if (!rows.length) rows.push(['', '', '']);
    state.project = project;
    state.fsrsData = synchronized.data;
    state.pairs = synchronized.pairs.map(normalizeStudyPair);
    state.sourceFile = file;
    state.editRows = rows;
    state.editDirty = false;
    state.editMessage = '';
    state.editMessageIsError = false;
    state.editFocus = null;
    state.editActiveCell = null;
    state.page = 'edit';
  } catch (error) {
    state.error = `${L('프로젝트를 편집용으로 열지 못했어요', 'Could not open the project for editing')}: ${error instanceof Error ? error.message : String(error)}`;
  } finally {
    processingProjectIds.delete(id);
    state.processingMessage = '';
  }
  render();
}

async function createBlankProject() {
  if (state.busy) return;
  state.error = '';
  state.notice = '';
  if (!state.directoryConnected) {
    const allowed = await ensureWorkbookDirectoryAccess();
    if (!allowed) return;
  }
  const input = window.prompt(L('새 프로젝트 이름을 입력해 주세요.', 'Enter a name for the new project.'));
  const name = input?.trim();
  if (!name) return;
  if (projectNameExists(state.projects.map((project) => project.name), name)) {
    state.error = L(`"${name}" 이름의 프로젝트가 이미 있어요. 다른 이름을 사용해 주세요.`, `A project named "${name}" already exists. Please use a different name.`);
    render();
    return;
  }
  state.busy = true;
  state.processingMessage = L('새 프로젝트를 만드는 중…', 'Creating the project…');
  render();
  let created = null;
  try {
    const fileName = `${name}.xlsx`;
    const workbook = await createStudyWorkbook([['', '', '']], ['key', 'value', 'hint']);
    const file = new File([workbook], fileName, { type: workbook.type });
    const project = {
      id: crypto.randomUUID(),
      name,
      fileName,
      sourceName: fileName,
      createdAt: Date.now(),
      pairs: [],
      title: name,
      a4Title: null,
      a4Info: null,
      formatVersion: 12,
      cartridgeColorIndex: state.projects.length % 4,
    };
    await createProjectDirectory(project.name);
    try {
      await writeProjectWorkbook(project.name, project.fileName, file);
      const synchronized = await synchronizeFsrsProject(project, []);
      project.pairs = synchronized.pairs;
      await saveProject(project);
    } catch (error) {
      await removeProjectDirectory(project.name).catch(() => {});
      throw error;
    }
    created = project;
    state.projects = [project, ...state.projects];
    state.projectOrder = state.projects.map((item) => item.id);
    await saveProjectOrder(state.projectOrder).catch(() => {});
  } catch (error) {
    state.error = `${L('프로젝트를 만들지 못했어요', 'Could not create the project')}: ${error instanceof Error ? error.message : String(error)}`;
  } finally {
    state.busy = false;
    state.processingMessage = '';
  }
  render();
  if (created) await editProject(created.id);
}

async function discardEmptyProject(project) {
  try {
    await removeProjectDirectory(project.name);
    releaseProjectImages(project.name);
    await deleteProjectRecord(project.id);
    state.projects = state.projects.filter((item) => item.id !== project.id);
    state.projectOrder = state.projects.map((item) => item.id);
    await saveProjectOrder(state.projectOrder).catch(() => {});
  } catch {
    // The empty project stays in the list if cleanup fails; the user can delete it manually.
  }
}

async function leaveEditPage() {
  if (state.editDirty && !window.confirm(L('저장하지 않은 변경 사항이 있어요. 나가시겠어요?', 'You have unsaved changes. Leave anyway?'))) return;
  const leaving = state.project;
  state.page = 'workspace';
  state.project = null;
  state.sourceFile = null;
  state.editRows = [];
  state.editDirty = false;
  render();
  if (leaving && !leaving.pairs.length) {
    await discardEmptyProject(leaving);
    render();
  }
}

async function saveEditedProject() {
  const project = state.project;
  const previousFile = state.sourceFile;
  if (!project || !previousFile || processingProjectIds.has(project.id)) return;
  const rows = state.editRows.map((row) => row.map((cell) => cell.replace(/\r\n?/g, '\n')));
  for (let index = 1; index < rows.length; index += 1) {
    const [key, value, hint] = rows[index];
    const hasKey = Boolean(key.trim());
    const hasValue = Boolean(value.trim());
    const hasHint = Boolean(hint.trim());
    if (hasKey !== hasValue || (!hasKey && hasHint)) {
      state.editMessage = L(`${index + 1}행: key(A열)와 value(B열)를 함께 입력해 주세요.`, `Row ${index + 1}: fill in both the key (column A) and value (column B).`);
      state.editMessageIsError = true;
      render();
      return;
    }
  }
  processingProjectIds.add(project.id);
  state.processingMessage = L('변경 사항을 저장하는 중…', 'Saving changes…');
  state.editMessage = L('저장하는 중…', 'Saving…');
  state.editMessageIsError = false;
  render();
  let fileWritten = false;
  let projectSaved = false;
  try {
    const workbook = await writeSheetGrid(previousFile, rows);
    const nextFile = new File([workbook], project.fileName, {
      type: previousFile.type,
      lastModified: Date.now(),
    });
    const pairs = await parseStudyFile(nextFile);
    if (!pairs.length) throw new Error(L('key와 value가 있는 행이 최소 1개 필요해요.', 'At least one row with both a key and value is required.'));
    const synchronized = await prepareFsrsSync(project, pairs);
    const nextRows = await readSheetGrid(nextFile);
    const updated = { ...project, pairs: synchronized.pairs };
    await writeProjectWorkbook(project.name, project.fileName, nextFile);
    fileWritten = true;
    await saveProject(updated);
    projectSaved = true;
    await saveFsrsData(updated, synchronized.data);
    state.projects = state.projects.map((item) => (item.id === updated.id ? updated : item));
    state.project = updated;
    state.fsrsData = synchronized.data;
    state.pairs = synchronized.pairs.map(normalizeStudyPair);
    state.sourceFile = nextFile;
    state.editRows = nextRows;
    state.editDirty = false;
    state.editMessage = L(`저장했어요. 학습 항목 ${pairs.length}개`, `Saved ${pairs.length} study items`);
    state.editMessageIsError = false;
  } catch (error) {
    const rollbackErrors = [];
    if (fileWritten) {
      try {
        await writeProjectWorkbook(project.name, project.fileName, previousFile);
      } catch (rollbackError) {
        rollbackErrors.push(rollbackError);
      }
    }
    if (projectSaved) {
      try {
        await saveProject(project);
      } catch (rollbackError) {
        rollbackErrors.push(rollbackError);
      }
    }
    const rollbackMessage = rollbackErrors.length
      ? ` ${L('원본으로 되돌리지 못했습니다.', 'Rollback failed.')}: ${rollbackErrors.map((item) => item instanceof Error ? item.message : String(item)).join('; ')}`
      : '';
    state.editMessage = `${L('저장하지 못했어요', 'Could not save')}: ${error instanceof Error ? error.message : String(error)}${rollbackMessage}`;
    state.editMessageIsError = true;
  } finally {
    processingProjectIds.delete(project.id);
    state.processingMessage = '';
  }
  render();
}

function editImageElement(id) {
  const wrapper = document.createElement('span');
  wrapper.className = 'edit-img';
  wrapper.contentEditable = 'false';
  wrapper.dataset.imgId = id;
  wrapper.innerHTML = `${renderImageText(imageTag(id), esc)}<span class="edit-img-tools" role="toolbar" aria-label="${L('이미지 편집', 'Edit image')}"><button type="button" data-edit-img="crop" title="${L('자르기', 'Crop')}" aria-label="${L('자르기', 'Crop')}"><svg viewBox="0 0 16 16" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 1.5v10a1 1 0 0 0 1 1h9.5M1.5 4H11a1 1 0 0 1 1 1v9.5"/></svg></button><button type="button" data-edit-img="replace" title="${L('교체', 'Replace')}" aria-label="${L('교체', 'Replace')}"><svg viewBox="0 0 16 16" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2.5 6.5a5.5 5.5 0 0 1 9.6-2.4l1.4 1.6M13.5 2v3.7h-3.7M13.5 9.5a5.5 5.5 0 0 1-9.6 2.4l-1.4-1.6M2.5 14v-3.7h3.7"/></svg></button><button type="button" class="is-danger" data-edit-img="delete" title="${L('삭제', 'Delete')}" aria-label="${L('삭제', 'Delete')}"><svg viewBox="0 0 16 16" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2.5 4h11M6 4V2.5h4V4M4 4l.7 9.2a1 1 0 0 0 1 .8h4.6a1 1 0 0 0 1-.8L12 4M6.7 7v4M9.3 7v4"/></svg></button></span>`;
  return wrapper;
}

function editCellHTML(text) {
  return splitImageText(text)
    .map((part) => (part.type === 'image' ? editImageElement(part.id).outerHTML : esc(part.value)))
    .join('');
}

// 셀 안의 글자와 이미지를 화면에 놓인 순서 그대로 [img:ID] 가 섞인 텍스트로 되돌립니다.
function serializeEditNode(node) {
  let text = '';
  node.childNodes.forEach((child) => {
    if (child.nodeType === Node.TEXT_NODE) {
      text += child.textContent.replace(/\u00a0/g, ' ');
    } else if (child.classList?.contains('edit-img')) {
      text += imageTag(child.dataset.imgId);
    } else if (child.nodeName === 'BR') {
      text += '\n';
    } else if (child.nodeType === Node.ELEMENT_NODE) {
      const inner = serializeEditNode(child);
      text += (text && !text.endsWith('\n') ? '\n' : '') + inner;
    }
  });
  return text;
}

const serializeEditCell = (cell) => serializeEditNode(cell).replace(/\n+$/, '');

function setEditStatus(message, isError = false) {
  state.editMessage = message;
  state.editMessageIsError = isError;
  const status = appEl.querySelector('.edit-status');
  if (status) {
    status.textContent = message;
    status.classList.toggle('is-error', isError);
  }
}

function syncEditCell(cell) {
  state.editRows[Number(cell.dataset.editRow)][Number(cell.dataset.editCol)] = serializeEditCell(cell);
  state.editDirty = true;
  const saveButton = appEl.querySelector('[data-edit-save]');
  if (saveButton && !processingProjectIds.size) saveButton.disabled = false;
}

function insertEditImage(cell, id) {
  const wrapper = editImageElement(id);
  const selection = window.getSelection();
  if (selection.rangeCount && cell.contains(selection.anchorNode)) {
    const range = selection.getRangeAt(0);
    range.deleteContents();
    range.insertNode(wrapper);
    range.setStartAfter(wrapper);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
  } else {
    cell.append(wrapper);
  }
  syncEditCell(cell);
}

async function addEditImages(cell, files) {
  if (!files.length || !state.project) return;
  state.processingMessage = L('이미지를 저장하는 중…', 'Saving images…');
  renderProcessingIndicator();
  try {
    for (let index = 0; index < files.length; index += 1) {
      state.processingMessage = L(
        `이미지 처리 중 ${index + 1}/${files.length}`,
        `Processing images ${index + 1}/${files.length}`
      );
      renderProcessingIndicator();
      try {
        insertEditImage(cell, await saveImageBlob(files[index], state.project.name));
      } catch {
        setEditStatus(L('이미지를 읽지 못했어요.', 'Could not read the image.'), true);
      }
    }
  } finally {
    state.processingMessage = '';
    renderProcessingIndicator();
  }
}

function pickImageFile() {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    input.hidden = true;
    const finish = (file) => {
      input.remove();
      resolve(file);
    };
    input.addEventListener('change', () => finish(input.files?.[0] || null));
    input.addEventListener('cancel', () => finish(null));
    document.body.append(input);
    input.click();
  });
}

// 선택한 영역만 남긴 이미지를 돌려줍니다. 취소하면 null. 원본 파일은 다른 셀이 쓸 수 있어 그대로 둡니다.
async function openCropDialog(id) {
  const url = await getImageUrl(id, state.project.name);
  if (!url) throw new Error('missing image');
  const ratios = [
    ['free', L('자유', 'Free'), 0],
    ['original', L('원본', 'Original'), -1],
    ['1:1', '1:1', 1],
    ['4:3', '4:3', 4 / 3],
    ['16:9', '16:9', 16 / 9],
  ];
  return new Promise((resolve) => {
    const overlay = document.createElement('div');
    overlay.className = 'crop-overlay';
    overlay.innerHTML = `
      <div class="crop-dialog" role="dialog" aria-modal="true" aria-label="${L('이미지 자르기', 'Crop image')}">
        <div class="crop-toolbar">
          <div class="crop-ratios" role="group" aria-label="${L('비율', 'Aspect ratio')}">
            ${ratios.map(([key, label]) => `<button type="button" class="crop-chip${key === 'free' ? ' is-active' : ''}" data-crop-ratio="${key}" aria-pressed="${key === 'free'}">${label}</button>`).join('')}
          </div>
          <button type="button" class="crop-chip" data-crop-reset>${L('초기화', 'Reset')}</button>
        </div>
        <div class="crop-stage">
          <div class="crop-area">
            <img src="${url}" alt="" draggable="false">
            <div class="crop-shade"></div>
            <div class="crop-box" tabindex="0" aria-label="${L('자를 영역', 'Crop area')}">
              <span class="crop-grid" aria-hidden="true"></span>
              ${['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].map((handle) => `<span class="crop-handle" data-h="${handle}"></span>`).join('')}
            </div>
          </div>
        </div>
        <p class="crop-hint">${L('모서리와 변을 끌어 크기를, 안쪽을 끌어 위치를 조절하세요. 바깥을 끌면 새로 선택해요.', 'Drag the edges to resize and the inside to move. Drag outside to start a new selection.')}</p>
        <div class="crop-actions">
          <span class="crop-size" data-crop-size></span>
          <button type="button" class="btn" data-crop-cancel>${L('취소', 'Cancel')}</button>
          <button type="button" class="btn btn-primary" data-crop-apply>${L('자르기 적용', 'Apply crop')}</button>
        </div>
      </div>`;
    document.body.append(overlay);
    const area = overlay.querySelector('.crop-area');
    const img = area.querySelector('img');
    const shade = area.querySelector('.crop-shade');
    const box = area.querySelector('.crop-box');
    const sizeEl = overlay.querySelector('[data-crop-size]');
    const applyButton = overlay.querySelector('[data-crop-apply]');
    const MIN = 16;
    let W = 0;
    let H = 0;
    let ratio = 0;
    let r = { l: 0, t: 0, r: 0, b: 0 };
    let drag = null;
    const clamp = (value, min, max) => Math.min(Math.max(value, min), max);

    const paint = () => {
      Object.assign(box.style, { left: `${r.l}px`, top: `${r.t}px`, width: `${r.r - r.l}px`, height: `${r.b - r.t}px` });
      shade.style.clipPath = `polygon(evenodd, 0 0, ${W}px 0, ${W}px ${H}px, 0 ${H}px, 0 0, ${r.l}px ${r.t}px, ${r.l}px ${r.b}px, ${r.r}px ${r.b}px, ${r.r}px ${r.t}px, ${r.l}px ${r.t}px)`;
      const outW = Math.round((r.r - r.l) * (img.naturalWidth / W));
      const outH = Math.round((r.b - r.t) * (img.naturalHeight / H));
      sizeEl.textContent = `${outW} × ${outH}px`;
    };

    const fitRatio = () => {
      if (!ratio) return;
      const cx = (r.l + r.r) / 2;
      const cy = (r.t + r.b) / 2;
      let w = r.r - r.l;
      let h = w / ratio;
      if (h > r.b - r.t) {
        h = r.b - r.t;
        w = h * ratio;
      }
      r = { l: cx - w / 2, t: cy - h / 2, r: cx + w / 2, b: cy + h / 2 };
    };

    const initial = () => {
      const inset = 0.08;
      r = { l: W * inset, t: H * inset, r: W * (1 - inset), b: H * (1 - inset) };
      fitRatio();
    };

    const resize = (handle, start, dx, dy) => {
      let { l, t, r: right, b } = start;
      if (handle.includes('w')) l = clamp(start.l + dx, 0, start.r - MIN);
      if (handle.includes('e')) right = clamp(start.r + dx, start.l + MIN, W);
      if (handle.includes('n')) t = clamp(start.t + dy, 0, start.b - MIN);
      if (handle.includes('s')) b = clamp(start.b + dy, start.t + MIN, H);
      if (ratio) {
        const horizontal = handle.includes('w') || handle.includes('e');
        const vertical = handle.includes('n') || handle.includes('s');
        if (horizontal && vertical) {
          const anchorX = handle.includes('w') ? start.r : start.l;
          const anchorY = handle.includes('n') ? start.b : start.t;
          const availW = handle.includes('w') ? start.r : W - start.l;
          const availH = handle.includes('n') ? start.b : H - start.t;
          let w = Math.max(right - l, (b - t) * ratio);
          w = Math.min(w, availW, availH * ratio);
          const h = w / ratio;
          l = handle.includes('w') ? anchorX - w : anchorX;
          right = handle.includes('w') ? anchorX : anchorX + w;
          t = handle.includes('n') ? anchorY - h : anchorY;
          b = handle.includes('n') ? anchorY : anchorY + h;
        } else if (horizontal) {
          const cy = (start.t + start.b) / 2;
          let h = (right - l) / ratio;
          h = Math.min(h, cy * 2, (H - cy) * 2);
          const w = h * ratio;
          if (handle === 'w') l = right - w; else right = l + w;
          t = cy - h / 2;
          b = cy + h / 2;
        } else {
          const cx = (start.l + start.r) / 2;
          let w = (b - t) * ratio;
          w = Math.min(w, cx * 2, (W - cx) * 2);
          const h = w / ratio;
          if (handle === 'n') t = b - h; else b = t + h;
          l = cx - w / 2;
          right = cx + w / 2;
        }
      }
      r = { l, t, r: right, b };
    };

    const point = (event) => {
      const bounds = img.getBoundingClientRect();
      return { x: clamp(event.clientX - bounds.left, 0, W), y: clamp(event.clientY - bounds.top, 0, H) };
    };

    area.addEventListener('pointerdown', (event) => {
      const p = point(event);
      const handle = event.target.dataset?.h;
      if (handle) {
        drag = { mode: 'resize', handle, start: { ...r }, x: p.x, y: p.y };
      } else if (box.contains(event.target)) {
        drag = { mode: 'move', start: { ...r }, x: p.x, y: p.y };
        box.classList.add('is-moving');
      } else {
        r = { l: p.x, t: p.y, r: p.x, b: p.y };
        drag = { mode: 'resize', handle: 'se', start: { ...r }, x: p.x, y: p.y };
      }
      area.setPointerCapture(event.pointerId);
      event.preventDefault();
    });
    area.addEventListener('pointermove', (event) => {
      if (!drag) return;
      const p = point(event);
      const dx = p.x - drag.x;
      const dy = p.y - drag.y;
      if (drag.mode === 'move') {
        const w = drag.start.r - drag.start.l;
        const h = drag.start.b - drag.start.t;
        const l = clamp(drag.start.l + dx, 0, W - w);
        const t = clamp(drag.start.t + dy, 0, H - h);
        r = { l, t, r: l + w, b: t + h };
      } else {
        resize(drag.handle, drag.start, dx, dy);
      }
      paint();
    });
    const endDrag = () => {
      if (drag && drag.mode === 'resize' && r.r - r.l < MIN && r.b - r.t < MIN) {
        initial();
        paint();
      }
      drag = null;
      box.classList.remove('is-moving');
    };
    area.addEventListener('pointerup', endDrag);
    area.addEventListener('pointercancel', endDrag);

    overlay.querySelectorAll('[data-crop-ratio]').forEach((chip) => {
      chip.addEventListener('click', () => {
        const entry = ratios.find(([key]) => key === chip.dataset.cropRatio);
        ratio = entry[2] === -1 ? img.naturalWidth / img.naturalHeight : entry[2];
        overlay.querySelectorAll('[data-crop-ratio]').forEach((item) => {
          item.classList.toggle('is-active', item === chip);
          item.setAttribute('aria-pressed', String(item === chip));
        });
        fitRatio();
        paint();
      });
    });
    overlay.querySelector('[data-crop-reset]').addEventListener('click', () => {
      initial();
      paint();
    });

    const close = (result) => {
      window.removeEventListener('resize', onResize);
      document.removeEventListener('keydown', onKey, true);
      overlay.remove();
      resolve(result);
    };
    const apply = () => {
      const scaleX = img.naturalWidth / W;
      const scaleY = img.naturalHeight / H;
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.round((r.r - r.l) * scaleX));
      canvas.height = Math.max(1, Math.round((r.b - r.t) * scaleY));
      canvas.getContext('2d').drawImage(img, r.l * scaleX, r.t * scaleY, (r.r - r.l) * scaleX, (r.b - r.t) * scaleY, 0, 0, canvas.width, canvas.height);
      canvas.toBlob((blob) => close(blob), 'image/jpeg', 0.92);
    };
    const onKey = (event) => {
      if (event.key === 'Escape') {
        event.stopPropagation();
        close(null);
      } else if (event.key === 'Enter' && !event.target.closest?.('button')) {
        event.preventDefault();
        event.stopPropagation();
        apply();
      } else if (event.key.startsWith('Arrow') && !drag) {
        event.preventDefault();
        event.stopPropagation();
        const step = event.shiftKey ? 10 : 1;
        const dx = event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0;
        const dy = event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0;
        const w = r.r - r.l;
        const h = r.b - r.t;
        const l = clamp(r.l + dx, 0, W - w);
        const t = clamp(r.t + dy, 0, H - h);
        r = { l, t, r: l + w, b: t + h };
        paint();
      }
    };
    const onResize = () => {
      const nextW = img.clientWidth;
      const nextH = img.clientHeight;
      if (!nextW || !nextH || !W) return;
      const fx = nextW / W;
      const fy = nextH / H;
      r = { l: r.l * fx, t: r.t * fy, r: r.r * fx, b: r.b * fy };
      W = nextW;
      H = nextH;
      paint();
    };
    window.addEventListener('resize', onResize);
    document.addEventListener('keydown', onKey, true);
    overlay.querySelector('[data-crop-cancel]').addEventListener('click', () => close(null));
    overlay.addEventListener('pointerdown', (event) => { if (event.target === overlay) close(null); });
    box.addEventListener('dblclick', apply);
    applyButton.addEventListener('click', apply);

    img.decode().catch(() => {}).then(() => {
      W = img.clientWidth;
      H = img.clientHeight;
      initial();
      paint();
      box.focus({ preventScroll: true });
    });
  });
}

async function handleEditImageAction(button, cell) {
  const wrapper = button.closest('.edit-img');
  if (!wrapper) return;
  if (button.dataset.editImg === 'delete') {
    wrapper.remove();
    syncEditCell(cell);
    return;
  }
  if (button.dataset.editImg === 'crop') {
    try {
      const blob = await openCropDialog(wrapper.dataset.imgId);
      if (!blob) return;
      state.processingMessage = L('자른 이미지를 저장하는 중…', 'Saving the cropped image…');
      renderProcessingIndicator();
      wrapper.replaceWith(editImageElement(await saveImageBlob(blob, state.project.name)));
      syncEditCell(cell);
    } catch {
      setEditStatus(L('이미지를 자르지 못했어요.', 'Could not crop the image.'), true);
    } finally {
      state.processingMessage = '';
      renderProcessingIndicator();
    }
    return;
  }
  const file = await pickImageFile();
  if (!file) return;
  state.processingMessage = L('교체 이미지를 저장하는 중…', 'Saving the replacement image…');
  renderProcessingIndicator();
  try {
    wrapper.replaceWith(editImageElement(await saveImageBlob(file, state.project.name)));
    syncEditCell(cell);
  } catch {
    setEditStatus(L('이미지를 읽지 못했어요.', 'Could not read the image.'), true);
  } finally {
    state.processingMessage = '';
    renderProcessingIndicator();
  }
}

function attachEditCell(cell) {
  const row = Number(cell.dataset.editRow);
  cell.addEventListener('input', () => {
    syncEditCell(cell);
    if (state.editMessage) setEditStatus('');
  });
  cell.addEventListener('focus', () => {
    state.editActiveCell = { row, col: cell.dataset.editCol };
  });
  cell.addEventListener('keydown', (event) => {
    if (event.key !== 'Enter' || event.isComposing) return;
    event.preventDefault();
    if (event.shiftKey) {
      document.execCommand('insertLineBreak');
      return;
    }
    appEl.querySelector(`[data-edit-row="${row + 1}"][data-edit-col="${cell.dataset.editCol}"]`)?.focus();
  });
  cell.addEventListener('paste', (event) => {
    event.preventDefault();
    const images = [...(event.clipboardData?.files || [])].filter((file) => file.type.startsWith('image/'));
    if (images.length) addEditImages(cell, images);
    else document.execCommand('insertText', false, event.clipboardData?.getData('text/plain') || '');
  });
  cell.addEventListener('dragover', (event) => {
    if ([...(event.dataTransfer?.items || [])].some((item) => item.type.startsWith('image/'))) event.preventDefault();
  });
  cell.addEventListener('drop', (event) => {
    const images = [...(event.dataTransfer?.files || [])].filter((file) => file.type.startsWith('image/'));
    if (!images.length) return;
    event.preventDefault();
    event.stopPropagation();
    addEditImages(cell, images);
  });
  cell.addEventListener('click', (event) => {
    const button = event.target.closest?.('[data-edit-img]');
    if (button) handleEditImageAction(button, cell);
  });
}

function editHTML() {
  const columns = ['A', 'B', 'C'];
  const columnLabels = L(EDIT_COLUMN_LABELS, EDIT_COLUMN_LABELS_EN);
  const saving = state.project ? processingProjectIds.has(state.project.id) : false;
  const body = state.editRows.map((row, rowIndex) => `
    <tr class="${rowIndex === 0 ? 'is-header-row' : ''}">
      <th scope="row" class="edit-row-number">
        <span>${rowIndex + 1}</span>
        ${rowIndex > 0 ? `<button type="button" class="edit-row-delete" data-edit-delete-row="${rowIndex}" aria-label="${L(`${rowIndex + 1}행 삭제`, `Delete row ${rowIndex + 1}`)}" title="${L('행 삭제', 'Delete row')}">×</button>` : ''}
      </th>
      ${row.slice(0, columns.length).map((cell, colIndex) => `
        <td><div class="edit-cell" contenteditable="true" role="textbox" aria-multiline="true" spellcheck="false" data-edit-row="${rowIndex}" data-edit-col="${colIndex}" aria-label="${columns[colIndex]}${rowIndex + 1}">${editCellHTML(cell)}</div></td>`).join('')}
    </tr>`).join('');
  return `
    ${headerHTML()}
    <main class="shell edit-screen">
      <div class="edit-toolbar">
        <button type="button" class="btn edit-back" data-edit-back>${L('← 뒤로', '← Back')}</button>
        <h1 class="edit-title">${esc(state.project?.name || '')} <span>${L('편집', 'Edit')}</span></h1>
        <div class="edit-toolbar-actions">
          <span class="edit-status${state.editMessageIsError ? ' is-error' : ''}" role="status">${esc(state.editMessage)}</span>
          <button type="button" class="btn" data-edit-add-image title="${L('마지막으로 선택한 셀에 이미지를 추가해요. 셀에 이미지를 붙여넣거나 끌어다 놓아도 돼요.', 'Add an image to the last selected cell. You can also paste or drop an image into a cell.')}">${L('이미지 추가', 'Add image')}</button>
          <button type="button" class="btn btn-primary" data-edit-save${saving || !state.editDirty ? ' disabled' : ''}>${L('저장', 'Save')}</button>
        </div>
      </div>
      <p class="edit-hint">${L('1행은 제목 행이고, 2행부터 학습 항목이 시작돼요. A열은 key, B열은 value, C열은 hint예요. 열 제목 사이 경계를 드래그하거나 키보드 방향키로 열 너비를 조절할 수 있어요. 셀 안에서 줄을 바꾸려면 Shift+Enter를 누르세요. 이미지는 셀에 붙여넣거나 끌어다 놓거나 \'이미지 추가\'로 넣고, 이미지에 마우스를 올리면 교체·삭제할 수 있어요.', 'Row 1 is the header and study items start at row 2. Column A is key, B is value, and C is hint. Drag a divider between column headers or use the arrow keys to resize columns. Press Shift+Enter to start a new line inside a cell. Add images by pasting, dropping, or using Add image, and hover an image to replace or delete it.')}</p>
      <div class="edit-sheet-wrap">
        <table class="edit-sheet">
          <colgroup>
            <col data-edit-row-number-column>
            ${columns.map((_, index) => `<col data-edit-column="${index}">`).join('')}
          </colgroup>
          <thead>
            <tr>
              <th class="edit-corner"></th>
              ${columns.map((letter, index) => `<th scope="col" class="edit-col-head edit-col-${letter}"><span>${letter}<small>${columnLabels[index]}</small></span>${index < columns.length - 1 ? `<button type="button" class="edit-column-resizer" data-resize-column="${index}" role="separator" aria-orientation="vertical" aria-label="${L(`${columnLabels[index]} 열과 ${columnLabels[index + 1]} 열 너비 조절`, `Resize ${columnLabels[index]} and ${columnLabels[index + 1]} column widths`)}" title="${L('드래그하여 열 너비 조절', 'Drag to resize columns')}"></button>` : ''}</th>`).join('')}
            </tr>
          </thead>
          <tbody>${body}</tbody>
        </table>
      </div>
      <button type="button" class="btn edit-add-row" data-edit-add-row>${L('+ 행 추가', '+ Add row')}</button>
    </main>`;
}

async function deleteProject(id) {
  if (processingProjectIds.has(id)) return;
  const project = state.projects.find((item) => item.id === id);
  if (!project) return;
  if (!state.directoryConnected) {
    const allowed = await ensureWorkbookDirectoryAccess();
    if (!allowed) return;
    state.notice = L('저장 폴더 권한을 허용했어요. 삭제하기를 다시 눌러 주세요.', 'Folder access granted. Please click Delete again.');
    render();
    return;
  }
  const confirmation = [
    L(
      `"${project.name}" 프로젝트 폴더와 그 안의 Excel 파일, FSRS 기록, Study Record, 이미지를 모두 삭제합니다.`,
      `This deletes the "${project.name}" project folder, including its workbook, FSRS data, Study Record, and all images.`
    ),
    L('삭제한 프로젝트는 복구할 수 없습니다. 계속할까요?', 'A deleted project cannot be recovered. Continue?'),
  ].join('\n\n');
  if (!window.confirm(confirmation)) return;

  state.error = '';
  state.notice = '';
  processingProjectIds.add(id);
  state.processingMessage = L('프로젝트와 관련 파일을 삭제하는 중…', 'Deleting the project and its files…');
  render();
  try {
    await removeProjectDirectory(project.name);
    releaseProjectImages(project.name);
    await deleteProjectRecord(project.id);
    state.projects = state.projects.filter((item) => item.id !== id);
    state.projectOrder = state.projects.map((item) => item.id);
    try {
      await saveProjectOrder(state.projectOrder);
    } catch (error) {
      state.error = `${L('프로젝트는 삭제했지만 목록 순서를 저장하지 못했어요', 'The project was deleted, but the list order could not be saved')}: ${error instanceof Error ? error.message : String(error)}`;
    }
    if (state.project?.id === id) {
      state.project = null;
      state.sourceFile = null;
      state.sourceName = '';
      state.pairs = [];
      state.fsrsData = null;
      state.pendingReview = null;
      state.study = null;
    }
    if (!state.error) {
      state.notice = L(
        `"${project.name}" 프로젝트 폴더와 연결된 파일 및 이미지를 삭제했어요.`,
        `Deleted the "${project.name}" project folder, including its files and images.`
      );
    }
  } catch (error) {
    state.error = `${L('프로젝트를 삭제하지 못했어요', 'Could not delete the project')}: ${error instanceof Error ? error.message : String(error)}`;
  } finally {
    processingProjectIds.delete(id);
    state.processingMessage = '';
  }
  render();
}

async function moveProject(draggedId, targetId, insertAfter) {
  const sourceIndex = state.projects.findIndex((project) => project.id === draggedId);
  const targetIndex = state.projects.findIndex((project) => project.id === targetId);
  if (sourceIndex < 0 || targetIndex < 0 || sourceIndex === targetIndex) return;

  const previousProjects = state.projects;
  const reordered = [...state.projects];
  const [project] = reordered.splice(sourceIndex, 1);
  const adjustedTargetIndex = reordered.findIndex((item) => item.id === targetId);
  reordered.splice(adjustedTargetIndex + (insertAfter ? 1 : 0), 0, project);
  state.projects = reordered;
  state.projectOrder = reordered.map((item) => item.id);
  state.error = '';
  render();

  try {
    await saveProjectOrder(state.projectOrder);
  } catch (error) {
    state.projects = previousProjects;
    state.projectOrder = previousProjects.map((item) => item.id);
    state.error = `${L('프로젝트 순서를 저장하지 못했어요', 'Could not save the project order')}: ${error instanceof Error ? error.message : String(error)}`;
    render();
  }
}

/* ---------------- HTML ---------------- */

function languageSelectHTML(id = 'languageSelect') {
  return `
    <label class="language-switch" for="${id}">
      <span class="language-switch-label">${L('언어', 'Language')}</span>
      <select id="${id}" class="language-select" data-language-select aria-label="${L('언어 설정', 'Language setting')}">
        <option value="ko"${language === 'ko' ? ' selected' : ''}>한국어</option>
        <option value="en"${language === 'en' ? ' selected' : ''}>English</option>
      </select>
    </label>`;
}

function languageChoiceHTML() {
  return `
    <div class="language-choice-modal" role="dialog" aria-modal="true" aria-labelledby="languageChoiceTitle">
      <div class="language-choice-backdrop"></div>
      <div class="language-choice-panel">
        <p class="language-choice-brand">ESPER</p>
        <h2 id="languageChoiceTitle">Choose your language · 언어를 선택하세요</h2>
        <p class="language-choice-hint">You can change this anytime in the settings. · 설정에서 언제든 바꿀 수 있어요.</p>
        <div class="language-choice-actions">
          <button type="button" class="btn btn-primary" data-choose-language="en">English user</button>
          <button type="button" class="btn btn-primary" data-choose-language="ko">한국어 사용자</button>
        </div>
      </div>
    </div>`;
}

const TUTORIAL_PAGE_URL = 'tutorial.html';
const SHARED_DECKS_URL = 'https://github.com/sunghokim128/ESPER/discussions/categories/shared-decks';

function headerHTML() {
  const projectsActive = ['home', 'workspace', 'edit'].includes(state.page);
  const canNavigateProject = Boolean(state.project);
  const activePage = state.page === 'study-records' ? 'progress' : state.page === 'learning-status' ? 'status' : state.page;
  const showGlobalNav = !['workspace', 'home'].includes(state.page);
  const workspaceActions = ['workspace', 'home'].includes(state.page)
    ? folderConnectionHTML(true)
    : '';
  return `
    <header class="site-head${showGlobalNav ? '' : ' projects-header'}">
      <button type="button" id="homeBtn" class="brand${projectsActive ? ' is-active' : ''}" aria-label="${L('프로젝트로 이동', 'Go to projects')}" title="${L('프로젝트로 이동합니다.', 'Go to projects.')}"${projectsActive ? ' aria-current="page"' : ''}>
        <p class="brand-name">ESPER</p>
        <span class="brand-acronym">Excel-based Spaced-repetition Program for Efficient Retention</span>
      </button>
      ${showGlobalNav ? `<nav class="global-nav" aria-label="${L('주요 페이지','Main pages')}">
        <button type="button" data-global-nav="study"${activePage === 'study' ? ' class="is-active" aria-current="page"' : ''}${canNavigateProject ? '' : ' disabled'}>${L('학습하기','Study')}</button>
        <button type="button" data-global-nav="status"${activePage === 'status' ? ' class="is-active" aria-current="page"' : ''}${canNavigateProject ? '' : ' disabled'}>${L('학습 상태','Learning status')}</button>
        <button type="button" data-global-nav="progress"${activePage === 'progress' ? ' class="is-active" aria-current="page"' : ''}${canNavigateProject ? '' : ' disabled'}>${L('나의 기록','My Log')}</button>
        <button type="button" data-global-nav="a4"${activePage === 'a4' ? ' class="is-active" aria-current="page"' : ''}${canNavigateProject ? '' : ' disabled'}>A4 view</button>
      </nav>` : ''}
      <div class="site-head-right">
        <a class="how-to-use-link" href="${TUTORIAL_PAGE_URL}" target="_blank" rel="noopener noreferrer">${L('튜토리얼', 'Tutorial')}</a>
        <a class="how-to-use-link" href="${SHARED_DECKS_URL}" target="_blank" rel="noopener noreferrer">Shared Decks</a>
        ${workspaceActions}
        ${languageSelectHTML()}
      </div>
    </header>`;
}

function startHTML() {
  return `
    ${headerHTML()}
    <main class="shell home-screen">
      ${formatGuideHTML()}
      <section class="home-upload-section" aria-label="${L('엑셀 학습 목록 업로드', 'Upload an Excel study set')}">
        <div class="guide-heading upload-section-heading">
          <h2>${L('새 학습 프로젝트 만들기', 'Create a new study project')}</h2>
          <p>${L('A4 view에서 항목을 한눈에 살펴보고, 아직 익숙하지 않은 항목은 학습하기 페이지의 플래시카드로 복습해 보세요.', 'Review items at a glance in A4 view, then practice unfamiliar ones with flashcards on the Study page.')}</p>
          <div class="workspace-page-actions">
            <button type="button" class="create-project-button" data-open-format-guide aria-haspopup="dialog">
              <span class="create-project-icon" aria-hidden="true">＋</span>
              <span>${L('학습 프로젝트 만들기', 'Create Learning Project')}</span>
            </button>
          </div>
        </div>
        ${uploadAreaHTML()}
        ${state.notice ? `<p class="notice">${esc(state.notice)}</p>` : ''}
        ${state.error ? `<p class="error">${esc(state.error)}</p>` : ''}
      </section>
    </main>`;
}

function chatGptPromptUrl() {
  if (!state.generationPrompt) return '';
  const url = new URL('https://chatgpt.com/');
  url.searchParams.set('q', state.generationPrompt.trim());
  return url.toString();
}

function formatExamplesHTML() {
  const examples = isEnglish()
    ? [
        {
          title: 'Vocabulary',
          sheet: 'Vocabulary',
          rows: [
              ['impair', 'to make something worse or weaker', 'Noise can impair your concentration.'],
              ['reluctant', 'unwilling and hesitant to do something', ''],
              ['mitigate', 'to make something less severe or harmful', 'Planting trees can mitigate flooding.'],
          ],
        },
        {
          title: 'Medical concept',
          sheet: 'Medical',
          rows: [
              ['Hypertension', 'persistently higher than normal blood pressure', 'Diagnosed when repeated readings reach 140/90 mmHg or above.'],
              ['Tachycardia', 'a resting heart rate above 100 beats per minute', ''],
              ['Anemia', 'a lower than normal amount of red blood cells or hemoglobin', 'Common signs: fatigue and pale skin.'],
          ],
        },
      ]
    : [
        {
          title: '단어 암기',
          sheet: '단어',
          rows: [
              ['impair', '손상시키다', 'Noise can impair your concentration.'],
              ['reluctant', '꺼리는, 마지못해 하는', ''],
              ['mitigate', '완화하다', 'Planting trees can mitigate flooding.'],
          ],
        },
        {
          title: '의학 개념 암기',
          sheet: '의학',
          rows: [
              ['Hypertension', '지속적으로 혈압이 정상보다 높은 상태', '수축기 140mmHg 이상 또는 이완기 90mmHg 이상이 반복 측정될 때 진단'],
              ['Tachycardia', '안정 시 심박수가 분당 100회를 넘는 상태', ''],
              ['Anemia', '적혈구 또는 헤모글로빈이 정상보다 적은 상태', '피로감, 창백한 피부가 흔한 증상'],
          ],
        },
      ];
  const headerLabels = ['key', 'value', 'hint'];
  const optionalMark = L('(선택)', '(optional)');
  const columnInfo = [
    {
      name: 'key',
      required: L('필수', 'Required'),
      description: L('암기할 대상입니다. 예: 영단어, 의학 용어, 질문.', 'The item to memorize, such as a word, medical term, or question.'),
    },
    {
      name: 'value',
      required: L('필수', 'Required'),
      description: L('key에 대한 답이나 정의입니다.', 'The answer or definition for the key.'),
    },
    {
      name: 'hint',
      required: L('선택', 'Optional'),
      description: L('기억을 돕는 단서입니다. 여러 줄을 입력하면 셀 안의 줄바꿈이 그대로 표시됩니다.', 'An optional memory aid. Line breaks entered in the cell appear in the hint as written.'),
    },
  ];
  const cellHTML = (ci, content) => {
    const column = columnInfo[ci];
    return `<td class="sheet-cell ${ci < 2 ? 'is-required' : 'is-optional'}" data-col="${ci}" data-guide-title="${String.fromCharCode(65 + ci)} · ${esc(column.name)}" data-guide-required="${esc(column.required)}" data-guide-description="${esc(column.description)}">${content}</td>`;
  };
  return `
    <div class="format-examples">
      ${examples.map((example, index) => `
        <article class="format-example-card" id="formatExamplePanel${index}" role="tabpanel" aria-labelledby="formatExampleSheet${index}" data-format-example-panel="${index}"${index === 0 ? '' : ' hidden'}>
          <table class="format-example-table">
            <thead>
              <tr class="sheet-letters"><th class="sheet-corner"></th>${columnInfo.map((column, ci) => `
                <th class="sheet-cell ${ci < 2 ? 'is-required' : 'is-optional'}" data-col="${ci}" data-guide-title="${String.fromCharCode(65 + ci)} · ${esc(column.name)}" data-guide-required="${esc(column.required)}" data-guide-description="${esc(column.description)}">${String.fromCharCode(65 + ci)}</th>`).join('')}</tr>
            </thead>
            <tbody>
              <tr class="sheet-header-row"><th class="sheet-rownum">1</th>${headerLabels.map((label, ci) => cellHTML(ci, esc(label))).join('')}</tr>
              ${example.rows.map((row, ri) => `<tr><th class="sheet-rownum">${ri + 2}</th>${row.map((value, ci) => cellHTML(ci, value ? esc(value) : (ci === 2 ? `<span class="sheet-optional-mark">${optionalMark}</span>` : ''))).join('')}</tr>`).join('')}
            </tbody>
          </table>
        </article>`).join('')}
      <div class="sheet-tabs" role="tablist" aria-label="${L('예시 엑셀 시트', 'Example spreadsheet sheets')}">
        ${examples.map((example, index) => `
          <button type="button" id="formatExampleSheet${index}" class="sheet-tab${index === 0 ? ' is-active' : ''}" role="tab" aria-selected="${index === 0}" aria-controls="formatExamplePanel${index}" tabindex="${index === 0 ? '0' : '-1'}" data-format-example="${index}">${esc(example.sheet)}</button>`).join('')}
      </div>
      <p class="format-example-note">${L('key / value / hint에 무엇을 넣을지는 자유예요.', 'You are free to choose what goes in key / value / hint.')}</p>
    </div>`;
}

function formatGuideHTML() {
  return `
    <div class="format-guide-modal" id="formatGuideModal" role="dialog" aria-modal="true" aria-labelledby="formatGuideTitle" hidden>
      <div class="format-guide-backdrop" data-close-format-guide></div>
      <div class="format-guide-panel">
        <div class="format-guide-header">
          <h2 id="formatGuideTitle">${L('학습 프로젝트 만들기', 'Create Learning Project')}</h2>
          <button type="button" class="format-guide-close" data-close-format-guide aria-label="${L('닫기', 'Close')}">×</button>
        </div>
      <div class="workspace-format-content">
        <section class="start-guide" aria-label="${L('엑셀 형식 안내', 'Excel format guide')}">
          <div class="guide-heading">
            <h2>${L('엑셀 파일 하나로 시작하세요', 'Start with one spreadsheet')}</h2>
            <p class="guide-lead">${L('항목(key)과 답(value), 힌트(hint)를 표로 정리하면 됩니다.', 'List each item (key), its answer (value), and an optional hint.')}</p>
          </div>
          <figure class="spreadsheet-example">
            <div class="example-visual">
              ${formatExamplesHTML()}
            </div>
            <figcaption class="guide-column-detail" aria-live="polite" aria-atomic="true">
              <p class="guide-default-message">${L('표의 열에 마우스를 올리면 해당 열의 입력 규칙을 확인할 수 있어요.', 'Hover over a column in the table to see its rules.')}</p>
              <div class="guide-column-selected" hidden>
                <strong class="guide-column-title"><span id="guideColumnNumber"></span><span id="guideColumnNameSeparator" class="guide-column-separator" aria-hidden="true">·</span><span id="guideColumnName"></span><span id="guideColumnRequiredSeparator" class="guide-column-separator" aria-hidden="true">·</span><span id="guideColumnRequired"></span></strong>
                <p id="guideColumnDescription"></p>
              </div>
            </figcaption>
          </figure>
        </section>
        <div class="workspace-guide-resources">
          <section class="workspace-guide-resource">
            <div class="guide-resource-actions">
              ${state.generationPrompt
                ? `<a class="guide-chatgpt-link" href="${esc(chatGptPromptUrl())}" target="_blank" rel="noopener noreferrer">${L('AI로 단어장 엑셀파일 생성', 'Create a vocabulary spreadsheet with AI')}</a>`
                : `<button type="button" class="guide-chatgpt-link" disabled title="${esc(state.generationPromptError || L('프롬프트를 불러오는 중이에요.', 'Loading the prompt…'))}">${state.generationPromptError ? L('바로가기를 준비하지 못했어요', 'Could not prepare the link') : L('프롬프트 불러오는 중…', 'Loading the prompt…')}</button>`}
              <a class="anki-shared-decks-link" href="https://ankiweb.net/shared/decks" target="_blank" rel="noopener noreferrer"><svg class="anki-star" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 1.8l2.6 6.2 6.7.6-5.1 4.4 1.6 6.6L12 16l-5.8 3.6 1.6-6.6L2.7 8.6l6.7-.6z"/></svg>${L('Anki Shared Deck 둘러보기', 'Browse Anki Shared Decks')}</a>
              <button type="button" class="guide-tutorial-deck-link" data-start-tutorial-deck>${L('튜토리얼 덱 체험', 'Try the tutorial deck')}</button>
              <button type="button" class="guide-create-blank-link" data-create-blank-project>${L('직접 만들기', 'Create from scratch')}</button>
            </div>
          </section>
        </div>
      </div>
      </div>
    </div>`;
}

function folderConnectionHTML(compact = false) {
  const issue = state.directoryAccessIssue
    ? `<span class="folder-location-issue" role="status">${esc(state.directoryAccessIssue)}</span>`
    : '';
  const hasLocation = Boolean(state.directoryName && !state.directoryAccessIssue);
  const buttonLabel = hasLocation
    ? `${L('저장 위치', 'Save location')}: ${state.directoryName} · ${L('변경', 'Change')}`
    : L('저장 위치 선택', 'Choose save location');
  const warningText = L('이전에 저장한 내용을 불러오려면 같은 저장 위치를 선택해야 해요.', 'To load previously saved content, choose the same save location as before.');
  const buttonTitle = hasLocation ? `${buttonLabel}. ${warningText}` : buttonLabel;
  const selectionButton = `<button type="button" id="connectFolderBtn" class="btn folder-location-button" title="${esc(buttonTitle)}">${esc(buttonLabel)}</button>`;
  const warning = state.directoryName && !compact
    ? `<span class="folder-location-warning">${warningText}</span>`
    : '';

  return `
    <div class="folder-connection${compact ? ' folder-connection-compact' : ''}${hasLocation ? ' has-location' : ''}${state.directoryAccessIssue ? ' has-issue' : ''}">
      ${issue}
      <div class="folder-connection-action">
        ${selectionButton}
        ${warning}
      </div>
    </div>`;
}

function workspaceHTML() {
  const palette = ['#B5573A', '#2F6F5E', '#4F5BD5', '#6B5B4B'];
  const projectCards = state.projects.map((project, index) => {
    const itemCount = project.pairs.length;
    const extension = (project.sourceName || project.fileName).match(/\.[^.]+$/)?.[0] || '.xlsx';
    return `
      <article class="project-card" data-project-card="${esc(project.id)}" draggable="true" style="--c:${palette[(project.cartridgeColorIndex ?? index) % palette.length]}">
        <div class="project-menu">
          <button type="button" class="project-menu-toggle" data-project-menu-toggle aria-haspopup="menu" aria-expanded="false" aria-label="${esc(L(`${project.name} 설정`, `${project.name} options`))}">
            <span aria-hidden="true">⋯</span>
          </button>
          <div class="project-menu-list" role="menu" hidden>
            <button type="button" role="menuitem" class="project-menu-item" data-edit-project="${esc(project.id)}"${processingProjectIds.has(project.id) ? ' disabled' : ''}>
              ${L('프로젝트 편집', 'Edit project')}
            </button>
            <button type="button" role="menuitem" class="project-menu-item" data-share-project="${esc(project.id)}"${processingProjectIds.has(project.id) ? ' disabled' : ''}>
              ${L('프로젝트 공유', 'Share project')}
            </button>
            <button type="button" role="menuitem" class="project-menu-item is-danger" data-delete-project="${esc(project.id)}"${processingProjectIds.has(project.id) ? ' disabled' : ''}>
              ${L('프로젝트 삭제하기', 'Delete project')}
            </button>
          </div>
        </div>
        <button type="button" class="project-open" data-open-project="${esc(project.id)}" aria-label="${esc(L(`${project.name} 학습하기, ${itemCount}개 항목`, `Study ${project.name}, ${itemCount} items`))}"${processingProjectIds.has(project.id) ? ' disabled' : ''}>
          <span class="cartridge-stage" aria-hidden="true">
            <span class="cartridge">
              <span class="cartridge-shell">
                <span class="cartridge-ridge"></span>
                <span class="cartridge-label">
                  <span class="cartridge-label-head"><span>EXCEL</span><b>${esc(extension)}</b></span>
                  <span class="cartridge-title">${esc(project.name)}</span>
                  <span class="cartridge-label-foot"><span>${itemCount} ITEMS</span><span>▶ PLAY</span></span>
                </span>
              </span>
              <span class="cartridge-pins"></span>
            </span>
            <span class="console-lip"></span>
          </span>
          <span class="project-info">
            <strong class="project-name">${esc(project.name)}</strong>
            <span class="project-meta">
              <span class="project-review-count" data-project-review-count="${esc(project.id)}" hidden></span>
            </span>
          </span>
        </button>
        <div class="project-exam-readiness" data-project-exam-readiness="${esc(project.id)}" data-study-hint="${esc(L(
          '색이 있는 왼쪽은 복습한 항목의 시험일 예상 기억률, 회색 오른쪽은 아직 복습하지 않은 항목의 비율이에요.',
          'The colored left side shows reviewed items’ expected exam-day recall; the gray right side shows the share not yet reviewed.'
        ))}">
          <div class="project-exam-readiness-head">
            <span>${L('D-day 예상 평균 기억률', 'Expected D-day average')}</span>
            <strong data-project-exam-value>${projectPlan(project).examDate ? '…' : '—'}</strong>
          </div>
          <div class="project-exam-meter" role="group" aria-label="${L('시험일 예상 기억률 및 미복습 비율', 'Expected exam-day recall and unreviewed share')}" title="${L('색이 있는 왼쪽 영역은 복습한 항목의 시험일 예상 평균 기억률이고, 회색 오른쪽 영역은 아직 복습하지 않은 항목의 비율이에요.', 'The colored left section shows expected exam-day average recall for reviewed items; the gray right section shows the share not yet reviewed.')}">
            <div class="project-exam-reviewed" role="progressbar" aria-label="${L('복습 카드의 시험일 예상 평균 기억률', 'Expected exam-day average for reviewed cards')}" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0" aria-valuetext="${L('복습 카드 비율을 계산 중', 'Calculating reviewed card share')}" data-project-exam-reviewed>
              <span data-project-exam-meter></span>
            </div>
            <div class="project-exam-new" role="img" aria-label="${L('미복습 카드 비율을 계산 중', 'Calculating unreviewed card share')}" data-project-exam-new></div>
          </div>
        </div>
      </article>`;
  }).join('');

  return `
    ${headerHTML()}
    <main class="shell workspace-screen">
      ${formatGuideHTML()}
      ${state.notice ? `<p class="notice" role="status">${esc(state.notice)}</p>` : ''}
      ${state.error ? `<p class="error" role="alert">${esc(state.error)}</p>` : ''}
      <div class="workspace-heading">
        <div class="workspace-heading-copy">
          <p class="workspace-eyebrow">${L('나의 학습 공간','YOUR STUDY SPACE')}</p>
          <h1>${L('프로젝트','Projects')}</h1>
          <p>${L('프로젝트를 선택해 학습을 시작하거나 A4 view와 나의 기록을 확인하세요.','Choose a project to study, open its A4 view, or check your log.')}</p>
        </div>
        <div class="workspace-page-actions">
          <button type="button" class="create-project-button" data-open-format-guide aria-haspopup="dialog">
            <span class="create-project-icon" aria-hidden="true">＋</span>
            <span>${L('학습 프로젝트 만들기', 'Create Learning Project')}</span>
          </button>
        </div>
      </div>
      <section class="project-list" aria-label="${L('엑셀 카트리지 목록', 'Excel cartridges')}">
        ${projectCards}
        <button type="button" class="cartridge-slot" data-add-project aria-label="${L('규격에 맞는 엑셀, Anki .apkg, ESPER .espr 파일 선택', 'Choose a formatted Excel, Anki .apkg or ESPER .espr file')}">
          <span class="cartridge-slot-ghost" aria-hidden="true">＋</span>
          <strong>${L('규격에 맞는 엑셀, Anki .apkg, ESPER .espr 파일을 업로드하세요', 'Upload a formatted Excel, Anki .apkg or ESPER .espr file')}</strong>
          <span>${L('프로젝트 이름과 저장할 Excel 파일명은 업로드한 파일명으로 고정돼요. 저장 폴더에서 프로젝트 폴더나 파일 이름을 바꾸지 마세요. 같은 이름의 프로젝트는 만들 수 없습니다.', 'The project name and saved workbook filename are fixed to the uploaded filename. Do not rename the project folder or its files. Project names must be unique.')}</span>
        </button>
      </section>
    </main>`;
}

function studyCurrentPair() {
  const session = state.study;
  if (!session?.started) return null;
  const id = session.cards[session.index];
  return state.pairs.find((pair) => pair.id === id) || null;
}

function fsrsRecordForCard(id) {
  return state.fsrsData?.cards.find((record) => record.id === id && record.archivedAt === null) || null;
}

async function refreshWorkspaceExamGauges() {
  const generation = ++workspaceExamGaugeGeneration;
  const gauges = [...appEl.querySelectorAll('[data-project-exam-readiness]')];
  await Promise.all(gauges.map(async (gauge) => {
    const project = state.projects.find((item) => item.id === gauge.dataset.projectExamReadiness);
    if (!project) return;
    const value = gauge.querySelector('[data-project-exam-value]');
    const fill = gauge.querySelector('[data-project-exam-meter]');
    const reviewedSegment = gauge.querySelector('[data-project-exam-reviewed]');
    const newSegment = gauge.querySelector('[data-project-exam-new]');
    const reviewCount = gauge.closest('[data-project-card]')?.querySelector('[data-project-review-count]');
    const { examDate, desiredRetention } = projectPlan(project);
    let reviewedPercent = 0;
    let newPercent = 0;
    const isCurrent = () =>
      generation === workspaceExamGaugeGeneration &&
      state.page === 'workspace' &&
      gauge.isConnected;
    const setUnavailable = () => {
      if (!isCurrent()) return;
      value.textContent = '—';
      reviewedSegment.setAttribute('aria-valuenow', '0');
      reviewedSegment.setAttribute('aria-valuetext', L(
        `복습한 항목 ${reviewedPercent}% · 미복습 항목 ${newPercent}% · 시험일 예상 평균 기억률 계산 불가`,
        `${reviewedPercent}% reviewed · ${newPercent}% unreviewed · exam-day average unavailable`
      ));
      fill.style.width = '0%';
    };

    try {
      const stored = await readFsrsDataFile(project.name, project.fileName);
      let cards = [];
      if (stored !== null) {
        const data = JSON.parse(stored);
        if (
          data?.projectId !== project.id ||
          data?.workbookFileName !== project.fileName ||
          data?.library?.model !== 'FSRS-7' ||
          !Array.isArray(data.cards)
        ) {
          throw new Error('FSRS file does not match this project or model.');
        }
        cards = data.cards.filter((record) => record?.archivedAt === null);
      }
      if (!isCurrent()) return;
      const activeById = new Map(cards.map((record) => [record.id, record]));
      const now = Date.now();
      const reviewedCards = project.pairs
        .map((pair) => activeById.get(pair.id))
        .filter((record) => record && record.card.state !== FsrsState.New && record.card.scheduleStatus !== 'new');
      const totalCount = project.pairs.length;
      const reviewedCount = reviewedCards.length;
      const newCount = Math.max(0, totalCount - reviewedCount);
      reviewedPercent = totalCount ? Math.round((reviewedCount / totalCount) * 100) : 0;
      newPercent = totalCount ? 100 - reviewedPercent : 0;
      reviewedSegment.style.width = `${reviewedPercent}%`;
      newSegment.style.width = `${newPercent}%`;
      reviewedSegment.title = L(
        `복습한 항목 ${reviewedPercent}% · 색이 채워진 정도는 해당 항목의 시험일 예상 평균 기억률`,
        `${reviewedPercent}% reviewed · fill shows their expected average recall on exam day`
      );
      newSegment.title = L(
        `아직 복습하지 않은 항목 ${newPercent}%`,
        `${newPercent}% not yet reviewed`
      );
      newSegment.setAttribute('aria-label', L(
        `미복습 항목 ${newPercent}%`,
        `${newPercent}% unreviewed`
      ));
      const dueCount = project.pairs.filter((pair) => {
        const record = activeById.get(pair.id);
        return !record || fsrsReviewCategory(record.card, now) !== 'scheduled';
      }).length;
      if (reviewCount) {
        reviewCount.textContent = L(`지금 복습할 항목 ${dueCount}개`, `${dueCount} ${dueCount === 1 ? 'item' : 'items'} to review now`);
        reviewCount.hidden = dueCount === 0;
      }

      if (!examDate || new Date(`${examDate}T23:59:59`) < new Date()) {
        setUnavailable();
        return;
      }
      const summary = await summarizeExamReadiness(
        reviewedCards,
        new Date(`${examDate}T23:59:59`),
        desiredRetention
      );
      if (!isCurrent()) return;
      const percent = summary.averageRetention === null
        ? null
        : Math.round(summary.averageRetention * 100);
      value.textContent = percent === null ? '—' : `${percent}%`;
      reviewedSegment.setAttribute('aria-valuenow', String(percent ?? 0));
      reviewedSegment.setAttribute('aria-valuetext', L(
        `복습한 항목 ${reviewedPercent}% · 미복습 항목 ${newPercent}% · 시험일 예상 평균 기억률 ${percent === null ? '계산 불가' : `${percent}%`}`,
        `${reviewedPercent}% reviewed · ${newPercent}% unreviewed · expected exam-day average ${percent === null ? 'unavailable' : `${percent}%`}`
      ));
      fill.style.width = `${percent ?? 0}%`;
      gauge.dataset.recallLevel = percent === null
        ? 'none'
        : percent >= 85 ? 'high' : percent >= 60 ? 'mid' : 'low';
    } catch (error) {
      if (reviewCount && isCurrent()) {
        reviewCount.textContent = L('복습 개수를 불러오지 못했어요.', 'Could not load review count.');
        reviewCount.hidden = false;
        reviewCount.dataset.reviewCount = 'error';
      }
      setUnavailable();
      console.error(`Could not refresh project status for "${project.name}".`, error);
    }
  }));
}

function fsrsOverview() {
  const pairIds = new Set(state.pairs.map((pair) => pair.id));
  const cards = (state.fsrsData?.cards || []).filter(
    (record) => record.archivedAt === null && pairIds.has(record.id)
  );
  const now = Date.now();
  const due = cards.filter(
    (record) => record.card.state === FsrsState.New || Date.parse(record.card.dueAt) <= now
  );
  const upcoming = cards
    .filter((record) => record.card.state !== FsrsState.New && Date.parse(record.card.dueAt) > now)
    .sort((a, b) => Date.parse(a.card.dueAt) - Date.parse(b.card.dueAt))[0];
  return {
    dueCount: due.length,
    newCount: cards.filter((record) => record.card.state === FsrsState.New).length,
    nextDueAt: upcoming?.card.dueAt || null,
  };
}

function formatFsrsDuration(value, baseMs) {
  const ms = new Date(value).getTime() - baseMs;
  if (!Number.isFinite(ms)) return '';
  const minutes = Math.max(1, Math.round(ms / 60_000));
  if (minutes < 60) return L(`${minutes}분 뒤`, `in ${minutes} min`);
  const hours = Math.round(minutes / 60);
  if (hours < 24) return L(`${hours}시간 뒤`, `in ${hours} h`);
  const days = Math.round(ms / 86_400_000);
  if (days < 30) return L(`${days}일 뒤`, `in ${days} d`);
  if (days < 365) {
    const months = Math.round((days / 30.4) * 10) / 10;
    return L(`${months}개월 뒤`, `in ${months} mo`);
  }
  const years = Math.round((days / 365) * 10) / 10;
  return L(`${years}년 뒤`, `in ${years} yr`);
}

function formatFsrsShortDate(value) {
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) return L('날짜 없음', 'No date');
  return new Intl.DateTimeFormat(language === 'ko' ? 'ko-KR' : 'en', { month: 'short', day: 'numeric' }).format(date);
}

function formatFsrsDate(value) {
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) return L('날짜 없음', 'No date');
  return new Intl.DateTimeFormat(language === 'ko' ? 'ko-KR' : 'en', {
    dateStyle: 'medium',
    timeStyle: 'short',
  }).format(date);
}

function fsrsStateLabel(card) {
  const labels = {
    new: L('새 카드', 'New'),
    learning: L('학습 중', 'Learning'),
    review: L('복습 중', 'Review'),
    relearning: L('재학습 중', 'Relearning'),
  };
  return labels[card?.scheduleStatus] || L('FSRS 상태 확인 필요', 'Unknown FSRS state');
}

function formatStudyTime(milliseconds) {
  const seconds = Math.max(0, Math.floor(milliseconds / 1000));
  return `${String(Math.floor(seconds / 60)).padStart(2, '0')}:${String(seconds % 60).padStart(2, '0')}`;
}

const RETENTION_WARNING_THRESHOLD = 95;

// Exam date and target retention belong to each project. Projects created before
// this change fall back to the old global values.
function projectPlan(project = state.project) {
  const plan = project?.plan;
  if (plan) {
    return {
      examDate: /^\d{4}-\d{2}-\d{2}$/.test(plan.examDate || '') ? plan.examDate : '',
      desiredRetention: normalizeDesiredRetention(plan.desiredRetention ?? DEFAULT_DESIRED_RETENTION),
    };
  }
  return { examDate: state.studySettings.examDate, desiredRetention: state.studySettings.desiredRetention };
}

function retentionWarningText() {
  return L('100%에 가까워질수록 복습양이 기하급수적으로 증가합니다. 시작은 90%로 두길 권장합니다', 'Review load grows exponentially as you approach 100%. Starting at 90% is recommended');
}

function retentionWarningHTML(percent, attr) {
  return `<p class="retention-warning" ${attr} role="status" hidden>${retentionWarningText()}</p>`;
}

async function refreshExamSummary() {
  const elements = [...appEl.querySelectorAll('[data-exam-summary]')];
  if (!elements.length) return;
  const { examDate, desiredRetention } = projectPlan();
  const set = (text) => elements.forEach((element) => { element.textContent = text; });
  if (!examDate) {
    set(L('시험 날짜가 설정되지 않았어요.', 'No exam date set.'));
    return;
  }
  const records = (state.fsrsData?.cards || []).filter((record) => record.archivedAt === null);
  const examEnd = new Date(`${examDate}T23:59:59`);
  try {
    const summary = await summarizeExamReadiness(records, examEnd, desiredRetention);
    if (projectPlan().examDate !== examDate) return;
    if (summary.daysLeft < 0) {
      set(L('시험 날짜가 이미 지났어요. 날짜를 다시 지정해 주세요.', 'The exam date has passed. Please choose a new date.'));
      return;
    }
    const average = summary.averageRetention === null ? '—' : `${Math.round(summary.averageRetention * 100)}%`;
    set(L(
      `D-${summary.daysLeft} · 시험일 예상 평균 기억률 ${average}`,
      `D-${summary.daysLeft} · expected exam-day average ${average}`
    ));
  } catch (error) {
    set(error instanceof Error ? error.message : String(error));
  }
}

function studySettingsHTML() {
  const settings = state.studySettings;
  return `
    <div class="study-settings-panel" id="studySettingsPanel" aria-label="${L('학습 설정', 'Study Setting')}">
      <section class="study-setting-card">
        <h2>${L('카드 글자 크기', 'Card text size')}</h2>
        <label class="study-font-setting">
          <span>${L('앞면', 'Front')}</span>
          <input type="range" min="18" max="120" step="1" data-study-font-size="front" value="${settings.fontSizes.front}" aria-label="${L('앞면 글자 크기', 'Front text size')}">
          <output data-study-font-size-value="front">${settings.fontSizes.front}px</output>
        </label>
        <label class="study-font-setting">
          <span>${L('뒷면', 'Back')}</span>
          <input type="range" min="16" max="72" step="1" data-study-font-size="back" value="${settings.fontSizes.back}" aria-label="${L('뒷면 글자 크기', 'Back text size')}">
          <output data-study-font-size-value="back">${settings.fontSizes.back}px</output>
        </label>
        <p class="study-font-setting-note">${L('글이 길면 카드 안에 맞게 자동으로 작아져요.', 'Long text automatically shrinks to fit the card.')}</p>
      </section>
      <section class="study-setting-card">
        <h2>${L('힌트 표시', 'Hint display')}</h2>
        <label class="study-hint-setting"><input type="checkbox" data-study-setting="showHintOnBack"${settings.showHintOnBack ? ' checked' : ''}> ${L('카드 뒷면에 hint 기본 표시', 'Show the hint on the card back by default')}</label>
      </section>
      <section class="study-setting-card">
        <h2>${L('key-value 반전', 'Swap key and value')}</h2>
        <label class="study-hint-setting"><input type="checkbox" data-study-setting="swapKeyValue"${settings.swapKeyValue ? ' checked' : ''}> ${L('value를 앞면, key를 뒷면으로 학습', 'Show value on the front and key on the back')}</label>
        <label class="study-hint-setting"><input type="checkbox" data-study-setting="randomizeKeyValue"${settings.randomizeKeyValue ? ' checked' : ''}> ${L('카드마다 key-value 방향을 무작위로 학습', 'Randomize key-value direction for each card')}</label>
      </section>
      <section class="study-setting-card">
        <h2>${L('이미지 크기', 'Image size')}</h2>
        ${imageSizeControlHTML('study')}
      </section>
      <details class="study-setting-card study-shortcuts"${state.studyShortcutsOpen ? ' open' : ''}>
        <summary>${L('키보드 단축키', 'Keyboard shortcuts')}</summary>
        ${STUDY_ACTIONS.map((entry) => { const [action] = entry; const label = studyActionLabel(entry); return `
          <div class="study-shortcut-row">
            <span>${esc(label)}</span>
            <button type="button" class="study-shortcut-key${state.studyKeyCapture === action ? ' is-capturing' : ''}" data-study-bind="${action}" aria-label="${esc(L(`${label} 단축키 지정`, `Set shortcut for ${label}`))}">
              ${state.studyKeyCapture === action ? L('키를 누르세요…', 'Press a key…') : esc(shortcutKeyLabel(settings.shortcuts[action]))}
            </button>
          </div>`; }).join('')}
        ${state.studyKeyError ? `<p class="study-settings-error" role="alert">${esc(state.studyKeyError)}</p>` : ''}
        <button type="button" class="study-reset-shortcuts" data-study-reset-shortcuts>${L('기본값으로 되돌리기', 'Reset to defaults')}</button>
      </details>
    </div>`;
}

function studyHTML() {
  const session = state.study;
  const pair = session?.started ? studyCurrentPair() : null;
  const shown = pair ? studyShownPair(pair) : null;
  const record = pair ? fsrsRecordForCard(pair.id) : null;
  const overview = fsrsOverview();
  const elapsed = session?.started ? Date.now() - session.startedAt : 0;
  const seenCount = session?.seenIds.size || 0;
  const ratings = [
    { grade: FsrsRating.Again, action: 'again', label: 'Again' },
    { grade: FsrsRating.Hard, action: 'hard', label: 'Hard' },
    { grade: FsrsRating.Good, action: 'good', label: 'Good' },
    { grade: FsrsRating.Easy, action: 'easy', label: 'Easy' },
  ];
  const cardMarkup = session && !session.started
    ? `<div class="study-ready">
        <p class="study-ready-eyebrow">READY WHEN YOU ARE</p>
        <h1>${L('준비되면 시작하세요','Start when you are ready')}</h1>
        <p>${L(
          `복습 예정 ${overview.dueCount - overview.newCount}개 · 새 카드 ${overview.newCount}개를 준비했어요.`,
          `${overview.dueCount - overview.newCount} reviews due · ${overview.newCount} new cards ready.`
        )}</p>
        <button type="button" class="study-start-button" data-study-action="start">Start</button>
      </div>`
    : pair && session
    ? `<article class="study-card${session.flipped ? ' is-flipped' : ''}${session.enterEffect ? ` enter-${session.enterEffect}` : ''}" data-study-flip tabindex="0" aria-label="${L('카드. Space 키로 뒤집기','Card. Press Space to flip')}">
        ${session.flipped
          ? `<div class="study-card-face study-card-back">
              <button type="button" class="study-speak" data-study-action="speak" aria-label="${L('key 발음 듣기','Listen to the key')}" title="${L('발음 듣기','Pronunciation')}">
                <svg aria-hidden="true" viewBox="0 0 24 24" fill="none"><path d="M4 10v4h3l4 3V7l-4 3H4Z" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"/><path d="M15 9a4 4 0 0 1 0 6m2-8a7 7 0 0 1 0 10" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/></svg>
              </button>
              <p class="study-value">${richText(shown.value)}</p>
              ${record ? `<p class="study-fsrs-meta">${esc(fsrsStateLabel(record.card))} · ${record.card.reps} ${L('회 복습','reviews')} · ${L('난이도','Difficulty')} ${record.card.difficulty.toFixed(1)} · ${L('안정성','Stability')} ${record.card.stability.toFixed(1)} ${L('일','days')}<br>${L('예정','Due')}: ${esc(formatFsrsDate(record.card.dueAt))}</p>` : ''}
              <div class="study-back-extras">
                ${session.settings.showHintOnBack || session.showBackHint
                  ? `<p class="study-hint">${pair.hint ? richText(pair.hint) : esc(L('등록된 hint가 없어요.','No hint available.'))}</p>`
                  : `<button type="button" class="study-reveal-extra" data-study-action="back-hint">${L('hint 보기','Show hint')}</button>`}
              </div>
            </div>`
          : `<div class="study-card-face study-card-front">
              <button type="button" class="study-speak" data-study-action="speak" aria-label="${L('key 발음 듣기','Listen to the key')}" title="${L('발음 듣기','Pronunciation')}">
                <svg aria-hidden="true" viewBox="0 0 24 24" fill="none"><path d="M4 10v4h3l4 3V7l-4 3H4Z" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"/><path d="M15 9a4 4 0 0 1 0 6m2-8a7 7 0 0 1 0 10" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/></svg>
              </button>
              <p class="study-key">${richText(shown.key)}</p>
              ${session.hintOpen
                ? `<p class="study-hint">${pair.hint ? richText(pair.hint) : esc(L('등록된 hint가 없어요.','No hint available.'))}</p>`
                : ''}
              <button type="button" class="study-hint-toggle" data-study-action="hint">
                ${session.hintOpen ? L('hint 닫기','Hide hint') : L('hint 보기','Show hint')}
              </button>
            </div>`}
      </article>`
    : `<div class="study-empty">
        <p>${session ? L('이번 학습을 마쳤어요.','You finished this study round.') : L('지금 복습할 카드가 없어요.','No cards are due right now.')}</p>
        <p>${session ? '' : overview.nextDueAt
          ? L(`다음 복습 예정: ${formatFsrsDate(overview.nextDueAt)}`, `Next review: ${formatFsrsDate(overview.nextDueAt)}`)
          : L('모든 카드가 복습되었어요. 새 항목을 추가하면 FSRS가 복습 일정을 시작해요.', 'All cards have been reviewed. Add items to start an FSRS schedule.')}</p>
        <button type="button" class="btn" data-study-action="restart">${L('다시 학습하기','Study again')}</button>
      </div>`;
  const shortcuts = Object.entries(session?.settings.shortcuts || state.studySettings.shortcuts)
    .map(([action, key]) => `${esc(shortcutKeyLabel(key))} ${esc(studyActionLabel(STUDY_ACTIONS.find(([id]) => id === action) || [action, '']))}`)
    .join(' · ');
  return `
    ${headerHTML()}
    <main class="study-screen study-workspace${studySidebarCollapsed ? ' study-sidebar-collapsed' : ''}${state.studyFullscreen && session?.started ? ' is-fullscreen' : ''}">
      <button type="button" id="reopenStudySidebarBtn" class="study-sidebar-reopen" aria-label="${L('학습 설정 펼치기','Expand Study Setting panel')}" title="${L('학습 설정 펼치기','Expand Study Setting panel')}"${studySidebarCollapsed ? '' : ' hidden'}>
        <svg aria-hidden="true" viewBox="0 0 24 24" fill="none">
          <path d="M19.14 12.94c.04-.3.06-.61.06-.94s-.02-.64-.07-.94l2.03-1.58a.5.5 0 0 0 .12-.64l-1.92-3.32a.5.5 0 0 0-.61-.22l-2.39.96a7.1 7.1 0 0 0-1.63-.94l-.36-2.54a.49.49 0 0 0-.49-.42h-3.84a.49.49 0 0 0-.49.42l-.36 2.54c-.59.23-1.14.55-1.63.94l-2.39-.96a.49.49 0 0 0-.61.22L2.63 8.84a.5.5 0 0 0 .12.64l2.03 1.58c-.05.3-.08.62-.08.94s.03.64.08.94l-2.03 1.58a.5.5 0 0 0-.12.64l1.92 3.32c.12.21.37.3.61.22l2.39-.96c.49.39 1.04.71 1.63.94l.36 2.54c.04.24.24.42.49.42h3.84c.25 0 .45-.18.49-.42l.36-2.54c.59-.23 1.14-.55 1.63-.94l2.39.96c.23.09.49-.01.61-.22l1.92-3.32a.5.5 0 0 0-.12-.64l-2.01-1.56ZM12 15.5a3.5 3.5 0 1 1 0-7 3.5 3.5 0 0 1 0 7Z" fill="currentColor"/>
        </svg>
      </button>
      <aside class="study-sidebar" aria-label="${L('학습 설정','Study Setting')}">
        <div class="study-sidebar-topbar">
          <h1 class="study-settings-title">${L('학습 설정','Study Setting')}</h1>
          <button type="button" id="collapseStudySidebarBtn" class="study-sidebar-collapse" aria-label="${L('학습 설정 접기','Collapse Study Setting panel')}" title="${L('학습 설정 접기','Collapse Study Setting panel')}">
            <svg aria-hidden="true" viewBox="0 0 24 24" fill="none">
              <path d="M19.14 12.94c.04-.3.06-.61.06-.94s-.02-.64-.07-.94l2.03-1.58a.5.5 0 0 0 .12-.64l-1.92-3.32a.5.5 0 0 0-.61-.22l-2.39.96a7.1 7.1 0 0 0-1.63-.94l-.36-2.54a.49.49 0 0 0-.49-.42h-3.84a.49.49 0 0 0-.49.42l-.36 2.54c-.59.23-1.14.55-1.63.94l-2.39-.96a.49.49 0 0 0-.61.22L2.63 8.84a.5.5 0 0 0 .12.64l2.03 1.58c-.05.3-.08.62-.08.94s.03.64.08.94l-2.03 1.58a.5.5 0 0 0-.12.64l1.92 3.32c.12.21.37.3.61.22l2.39-.96c.49.39 1.04.71 1.63.94l.36 2.54c.04.24.24.42.49.42h3.84c.25 0 .45-.18.49-.42l.36-2.54c.59-.23 1.14-.55 1.63-.94l2.39.96c.23.09.49-.01.61-.22l1.92-3.32a.5.5 0 0 0-.12-.64l-2.01-1.56ZM12 15.5a3.5 3.5 0 1 1 0-7 3.5 3.5 0 0 1 0 7Z" fill="currentColor"/>
            </svg>
          </button>
        </div>
        <div class="study-sidebar-content">
          ${studySettingsHTML()}
        </div>
      </aside>
      <section class="study-content" aria-label="${L('학습하기 페이지','Study page')}">
        <div class="study-topline">
          <div class="study-session-info">
            <span id="studyTimer" class="study-timer">${formatStudyTime(elapsed)}</span>
            ${session?.started ? `<button type="button" class="btn study-back-button" data-study-action="exit">${L('회차 종료','End round')}</button>` : ''}
            ${session ? `<span class="study-progress">${session.started
              ? `${Math.min(session.index + 1, session.cards.length)} / ${session.cards.length} · ${L(`이번 회차 ${seenCount}개`,`${seenCount} this round`)}`
              : L(`${session.cards.length}개 항목 준비됨`,`${session.cards.length} items ready`)}</span>` : ''}
          </div>
          ${session?.started ? `<button type="button" class="study-fullscreen-button" data-study-action="fullscreen">${state.studyFullscreen ? L('전체화면 종료','Exit full screen') : L('⛶ 전체화면','⛶ Full screen')}</button>` : ''}
        </div>
        <p class="study-error" role="alert"${state.studyError ? '' : ' hidden'}>${esc(state.studyError)}</p>
        <section class="study-stage" aria-label="${L('학습하기 페이지','Study page')}">
          ${cardMarkup}
        </section>
        ${pair ? `<div class="study-card-actions">
          ${!session.flipped
            ? `<span class="study-rating-prompt">${L('답을 확인한 뒤 복습 등급을 선택하세요.', 'Reveal the answer before choosing a review rating.')}</span>`
            : state.pendingReview
              ? `<button type="button" class="btn btn-primary" data-study-action="retry-review"${state.studySaving ? ' disabled' : ''}>${L('복습 저장 다시 시도', 'Retry saving review')}</button>`
              : ratings.map(({ grade, action, label }) => {
                const preview = session.previews?.find((item) => item.grade === grade);
                const dueText = preview ? formatFsrsDuration(preview.dueAt, session.previewBase) : '';
                return `<button type="button" class="btn study-rating-button study-rating-${action}" data-study-rating="${grade}"${state.studySaving ? ' disabled' : ''} title="${esc(dueText ? L(`다음 복습: ${dueText} (${formatFsrsDate(preview.dueAt)})`, `Next review ${dueText} (${formatFsrsDate(preview.dueAt)})`) : L(`${label} 평가`, `${label} rating`))}">
                  <span>${label}</span>${dueText ? `<small>${esc(dueText)}</small>` : ''}
                </button>`;
              }).join('')}
        </div>` : ''}
        <p class="study-key-hint">${shortcuts}</p>
      </section>
    </main>`;
}

function studyCalendarHTML(sessions) {
  const now = new Date();
  const cursor = state.studyCalendarMonth || `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  const [year, month] = cursor.split('-').map(Number);
  const prefix = `${year}-${String(month).padStart(2, '0')}-`;
  const byDay = new Map();
  for (const session of sessions) {
    const date = String(session.date).slice(0, 10);
    if (!date.startsWith(prefix)) continue;
    const day = Number(date.slice(8, 10));
    const entry = byDay.get(day) || { items: 0, count: 0 };
    entry.items += session.itemCount;
    entry.count += 1;
    byDay.set(day, entry);
  }
  const firstWeekday = new Date(year, month - 1, 1).getDay();
  const daysInMonth = new Date(year, month, 0).getDate();
  const today = localDateForExport();
  const cells = [];
  for (let i = 0; i < firstWeekday; i += 1) cells.push('<div class="study-cal-cell is-empty"></div>');
  for (let day = 1; day <= daysInMonth; day += 1) {
    const entry = byDay.get(day);
    const isToday = `${prefix}${String(day).padStart(2, '0')}` === today;
    cells.push(`<div class="study-cal-cell${entry ? ' is-done' : ''}${isToday ? ' is-today' : ''}"${entry ? ` title="${L(`${entry.count}회차 · ${entry.items}항목`,`${entry.count} rounds · ${entry.items} items`)}"` : ''}>
      <span class="study-cal-day">${day}</span>
      ${entry ? `<span class="study-cal-mark">✓</span>      <span class="study-cal-items">${entry.items}</span>` : ''}
    </div>`);
  }
  return `
    <section class="study-calendar" aria-label="${L('월간 학습 출석표','Monthly study attendance')}">
      <div class="study-cal-head">
        <button type="button" data-study-cal-move="-1" aria-label="${L('이전 달','Previous month')}">‹</button>
        <strong>${L(`${year}년 ${month}월 · ${byDay.size}일 출석`,`${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][month - 1]} ${year} · ${byDay.size} day${byDay.size === 1 ? '' : 's'} studied`)}</strong>
        <button type="button" data-study-cal-move="1" aria-label="${L('다음 달','Next month')}">›</button>
      </div>
      <div class="study-cal-grid">
        ${L(['일', '월', '화', '수', '목', '금', '토'], ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']).map((name) => `<div class="study-cal-weekday">${name}</div>`).join('')}
        ${cells.join('')}
      </div>
    </section>`;
}

function studyRecordsHTML() {
  const records = state.studyRecords;
  const sessions = [...(records?.sessions || [])].reverse();
  const keys = [...(records?.keys || [])].sort((a, b) => b.count - a.count || a.key.localeCompare(b.key));
  const totalItems = records?.sessions.reduce((sum, session) => sum + session.itemCount, 0) || 0;
  const totalDurationSeconds = records?.sessions.reduce((sum, session) => {
    const [minutes, seconds] = session.duration.split(':').map(Number);
    return Number.isFinite(minutes) && Number.isFinite(seconds)
      ? sum + minutes * 60 + seconds
      : sum;
  }, 0) || 0;
  return `
    ${headerHTML()}
    <main class="study-records-screen">
      <header class="study-records-heading">
        <p class="study-ready-eyebrow">${L('DAILY STREAK', 'DAILY STREAK')}</p>
        <h1>${L('나의 기록', 'My Log')}</h1>
        <p>${esc(state.project?.name || '')}</p>
      </header>
      ${state.studyRecordsLoading
        ? `<p class="study-records-empty" role="status">${L('기록을 불러오는 중…','Loading records…')}</p>`
        : state.studyRecordsError
          ? `<p class="study-error" role="alert">${esc(state.studyRecordsError)}</p>
             <button type="button" class="btn" data-study-records-retry>${L('다시 불러오기','Reload')}</button>`
          : records
            ? `${studyCalendarHTML(records.sessions)}
              <div class="study-records-summary">
                <article><strong>${sessions.length}</strong><span>${L('회차','Rounds')}</span></article>
                <article><strong>${totalItems}</strong><span>${L('총 학습 항목 수','Total items studied')}</span></article>
                <article><strong>${formatStudyTime(totalDurationSeconds * 1000)}</strong><span>${L('누적 학습 시간','Total study time')}</span></article>
              </div>
              <details class="study-records-section" data-study-records-details="sessions"${state.studyRecordsOpen?.sessions ? ' open' : ''}>
                <summary><h2>${L('회차별 기록','Rounds')} <span class="study-records-count">${sessions.length}</span></h2></summary>
                ${sessions.length
                  ? `<div class="study-records-table-wrap"><table class="study-records-table">
                      <thead><tr><th>${L('날짜','Date')}</th><th>${L('학습 항목 수','Items studied')}</th><th>${L('소요 시간','Duration')}</th></tr></thead>
                      <tbody>${sessions.map((session) => `<tr>
                        <td>${esc(session.date)}</td><td>${session.itemCount}</td><td>${esc(session.duration)}</td>
                      </tr>`).join('')}</tbody>
                    </table></div>`
                  : `<p class="study-records-empty">${L('아직 저장된 회차 기록이 없어요.','No saved rounds yet.')}</p>`}
              </details>
              <details class="study-records-section" data-study-records-details="keys"${state.studyRecordsOpen?.keys ? ' open' : ''}>
                <summary><h2>${L('key별 학습 횟수','Study count by key')} <span class="study-records-count">${keys.length}</span></h2></summary>
                ${keys.length
                  ? `<div class="study-records-table-wrap"><table class="study-records-table">
                      <thead><tr><th>Key</th><th>${L('학습 횟수','Study count')}</th></tr></thead>
                      <tbody>${keys.map((item) => `<tr><td>${richText(item.key)}</td><td>${item.count}</td></tr>`).join('')}</tbody>
                    </table></div>`
                  : `<p class="study-records-empty">${L('아직 저장된 key 기록이 없어요.','No saved key records yet.')}</p>`}
              </details>`
            : `<p class="study-records-empty">${L('저장된 학습 기록 파일이 없어요.','No study record file saved yet.')}</p>`}
    </main>`;
}

function learningStatusHTML() {
  const now = Date.now();
  const { byId: pairs, invalidCount: invalidPairCount } = indexStudyPairs(state.pairs);
  const records = (state.fsrsData?.cards || [])
    .filter((record) => record.archivedAt === null && pairs.has(record.id))
    .map((record) => ({
      record,
      pair: pairs.get(record.id),
      category: fsrsReviewCategory(record.card, now),
    }));
  const counts = {
    all: records.length,
    due: records.filter((item) => item.category === 'due').length,
    new: records.filter((item) => item.category === 'new').length,
    scheduled: records.filter((item) => item.category === 'scheduled').length,
  };
  const filters = [
    ['all', L('전체', 'All'), counts.all],
    ['due', L('복습 대상', 'Due'), counts.due],
    ['new', L('새 카드', 'New'), counts.new],
    ['scheduled', L('복습 예정', 'Scheduled'), counts.scheduled],
  ];
  const sortedRecords = records.sort((a, b) => {
    const order = { due: 0, new: 1, scheduled: 2 };
    return order[a.category] - order[b.category] ||
      Date.parse(a.record.card.dueAt) - Date.parse(b.record.card.dueAt);
  });
  const statusLabels = {
    due: L('복습할 때예요', 'Review due'),
    new: L('새 카드', 'New'),
    scheduled: L('복습 예정', 'Scheduled'),
  };
  return `
    ${headerHTML()}
    <main class="learning-status-screen">
      <header class="learning-status-heading">
        <p class="study-ready-eyebrow">${L('FSRS 학습 상태', 'FSRS LEARNING STATUS')}</p>
        ${invalidPairCount ? `<p class="error" role="alert">${L(
          `학습 항목 ${invalidPairCount}개에서 카드 ID를 찾지 못해 상태 목록에서 제외했어요.`,
          `${invalidPairCount} study item(s) have no card ID and were excluded from the status list.`
        )}</p>` : ''}
      </header>
      <div class="learning-status-top">
      <details class="exam-plan-panel"${state.planPanelOpen ? ' open' : ''}>
        <summary class="exam-plan-summary"><strong>${L('핵심 설정', 'Key settings')}</strong><span>${L('시험 날짜', 'Exam date')} ${projectPlan().examDate ? esc(formatExamDate(projectPlan().examDate)) : L('미설정', 'not set')}</span><span>${L('목표 기억률', 'Target retention')} ${Math.round(projectPlan().desiredRetention * 100)}%</span></summary>
        <div class="exam-plan-controls">
          <div class="exam-plan-field">
            <span>${L('시험 날짜', 'Exam date')}</span>
            ${datePickerHTML('data-plan-setting="examDate"', projectPlan().examDate, '', true)}
          </div>
          <label class="exam-plan-field">
            <span>${L('목표 기억률', 'Target retention')} <strong data-retention-label>${Math.round(projectPlan().desiredRetention * 100)}%</strong></span>
            <input type="range" min="${Math.round(MIN_DESIRED_RETENTION * 100)}" max="${Math.round(MAX_DESIRED_RETENTION * 100)}" step="1" data-plan-setting="desiredRetention" value="${Math.round(projectPlan().desiredRetention * 100)}" disabled>
          </label>
        </div>
        ${retentionWarningHTML(Math.round(projectPlan().desiredRetention * 100), 'data-retention-note')}
        <p class="exam-plan-line" data-exam-summary></p>
        <div class="exam-plan-footer">
          <p class="exam-plan-hint">${L('저장하면 복습 중인 모든 카드의 다음 복습일이 새 기준으로 변경돼요.', 'Saving updates the next review date of all cards in review.')}</p>
          <div class="exam-plan-actions">
            <button type="button" class="btn" data-plan-edit>${L('변경하기', 'Edit')}</button>
            <button type="button" class="btn btn-primary" data-plan-save hidden>${L('저장하기', 'Save')}</button>
            <button type="button" class="btn" data-plan-cancel hidden>${L('취소', 'Cancel')}</button>
          </div>
        </div>
        <p class="exam-plan-line" data-plan-status role="status"></p>
      </details>
      <details class="fsrs-guide">
        <summary>${L('FSRS-7이란?', 'What is FSRS-7?')}</summary>
        <p>${L('복습 기록으로 카드마다 기억 상태를 추적하고, 망각 곡선으로 다음 복습 시점을 정하는 간격 반복 알고리즘이에요.', 'A spaced-repetition algorithm that tracks each card’s memory state and uses a forgetting curve to schedule the next review.')}</p>
        <dl>
          <dt>${L('기억 확률', 'Recall probability')}</dt>
          <dd>${L('지금 이 카드를 떠올릴 수 있을 확률 추정치. 시간이 지날수록 떨어져요.', 'Estimated chance you can recall the card now. It falls as time passes.')}</dd>
          <dt>${L('난이도', 'Difficulty')}</dt>
          <dd>${L('1~10. 높을수록 외우기 어려운 카드예요.', '1–10. Higher means harder to remember.')}</dd>
          <dt>${L('안정성', 'Stability')}</dt>
          <dd>${L('기억이 유지되는 기간(일). 클수록 천천히 잊어요.', 'How long the memory lasts, in days. Larger means slower forgetting.')}</dd>
        </dl>
        <p class="fsrs-guide-source">${L('출처', 'Source')}: <a href="https://github.com/open-spaced-repetition/ts-fsrs" target="_blank" rel="noopener noreferrer">ts-fsrs</a> · <a href="https://github.com/open-spaced-repetition/free-spaced-repetition-scheduler" target="_blank" rel="noopener noreferrer">Free Spaced Repetition Scheduler</a></p>
      </details>
      </div>
      <section class="learning-status-toolbar" aria-label="${L('카드 검색 및 필터', 'Search and filter cards')}">
        <label class="learning-status-search">
          <span>${L('항목 검색', 'Search items')}</span>
          <input type="search" data-learning-status-search value="${esc(state.learningStatusQuery)}" placeholder="${L('key, value, hint 검색', 'Search key, value, or hint')}" autocomplete="off">
        </label>
        ${imageSizeControlHTML('status')}
        <div class="learning-status-filters" role="group" aria-label="${L('학습 상태 필터', 'Learning status filter')}">
          ${filters.map(([id, label, count]) => `<button type="button" class="learning-status-filter${state.learningStatusFilter === id ? ' is-active' : ''}" data-learning-status-filter="${id}" aria-pressed="${state.learningStatusFilter === id}">${label}<span>${count}</span></button>`).join('')}
        </div>
        <button type="button" class="learning-status-swap" data-learning-status-swap aria-pressed="${state.learningStatusShowValue}" title="${L('카드에 표시할 항목을 key와 value 사이에서 바꿉니다.', 'Switch the card tiles between key and value.')}">
          ${state.learningStatusShowValue ? L('Value 표시 중 · Key로 전환', 'Showing value · Switch to key') : L('Key 표시 중 · Value로 전환', 'Showing key · Switch to value')}
        </button>
      </section>
      <p class="learning-status-footnote">${L(
        '카드가 진할수록 현재 기억 확률이 높아요. 확률은 현재 시각 기준 추정치이며 매시간 갱신돼요. 카드를 눌러 자세한 정보를 확인하세요.',
        'More opaque cards indicate a higher current recall probability. Estimates update hourly and may differ from actual recall. Select a card for details.'
      )}</p>
      <section class="learning-status-list" aria-label="${L('카드별 상태 목록', 'Card status list')}">
        ${sortedRecords.length
          ? sortedRecords.map(({ pair, record, category }) => `
            <button type="button" class="learning-status-card" data-learning-status-card="${category}" data-learning-status-open="${esc(record.id)}" data-search="${esc(`${pair.key} ${pair.value} ${pair.hint}`.toLocaleLowerCase())}" aria-haspopup="dialog" title="${L('눌러서 카드 정보 보기', 'Select to view card details')}">
              <span class="learning-status-card-key">${richText(state.learningStatusShowValue ? pair.value : pair.key)}</span>
              <span class="learning-status-card-recall" data-fsrs-retrievability="${esc(record.id)}">${category === 'new' ? L('학습 전', 'Not studied') : '…'}</span>
            </button>
            <template data-learning-status-detail="${esc(record.id)}">
              <div class="learning-status-dialog-card">
                <button type="button" class="learning-status-dialog-close" data-learning-status-close aria-label="${L('닫기', 'Close')}">×</button>
                <span class="learning-status-badge is-${category}">${statusLabels[category]}</span>
                <div class="learning-status-dialog-content">
                  <section>
                    <h3>${L('값', 'Value')}</h3>
                    <div class="learning-status-value">${richText(pair.value)}</div>
                    ${pair.hint ? `<h3>${L('힌트', 'Hint')}</h3><div class="learning-status-hint">${richText(pair.hint)}</div>` : ''}
                  </section>
                  <section class="learning-status-dialog-recall">
                    <span>${L('현재 기억 확률', 'Current recall probability')}</span>
                    <strong data-learning-status-dialog-recall>${category === 'new' ? L('학습 전', 'Not studied') : '…'}</strong>
                  </section>
                  <dl class="learning-status-stats">
                    <div><dt>${L('다음 복습', 'Next review')}</dt><dd>${category === 'new' ? '—' : esc(formatFsrsDate(record.card.dueAt))}</dd></div>
                    <div><dt>${L('난이도', 'Difficulty')}</dt><dd>${category === 'new' ? '—' : record.card.difficulty.toFixed(1)}</dd></div>
                    <div><dt>${L('안정성', 'Stability')}</dt><dd>${category === 'new' ? '—' : `${record.card.stability.toFixed(1)}${L('일', 'd')}`}</dd></div>
                  </dl>
                  <p class="learning-status-last-rating">${L('마지막 평가', 'Last rating')}: <strong>${
                  record.reviewLogs.length
                    ? ({
                      [FsrsRating.Again]: L('다시', 'Again'),
                      [FsrsRating.Hard]: L('어려움', 'Hard'),
                      [FsrsRating.Good]: L('알맞음', 'Good'),
                      [FsrsRating.Easy]: L('쉬움', 'Easy'),
                    }[record.reviewLogs.at(-1).rating] || L('알 수 없음', 'Unknown'))
                    : L('없음', 'None')
                }</strong></p>
                </div>
              </div>
            </template>`).join('')
          : `<p class="learning-status-empty">${L('표시할 학습 항목이 없어요.', 'There are no learning items to display.')}</p>`}
        ${sortedRecords.length ? `<p class="learning-status-empty" data-learning-status-empty hidden>${L('검색 조건에 맞는 항목이 없어요.', 'No cards match your search.')}</p>` : ''}
      </section>
      <dialog class="learning-status-dialog" aria-label="${L('학습 항목 상세 정보', 'Learning item details')}"></dialog>
    </main>`;
}

function a4ViewHTML() {
  const itemCounts = a4ItemCounts();
  return `
    ${headerHTML()}
    <main class="shell a4-view-screen${settingsCollapsed ? ' settings-collapsed' : ''}" style="--study-panel-width:${studyPanelWidth}px">
      <button type="button" id="reopenSettingsBtn" class="a4-settings-reopen" aria-label="${L('설정 펼치기','Expand settings')}" title="${L('설정 펼치기','Expand settings')}"${settingsCollapsed ? '' : ' hidden'}>
        <svg aria-hidden="true" viewBox="0 0 24 24" fill="none">
          <path d="M19.14 12.94c.04-.3.06-.61.06-.94s-.02-.64-.07-.94l2.03-1.58a.5.5 0 0 0 .12-.64l-1.92-3.32a.5.5 0 0 0-.61-.22l-2.39.96a7.1 7.1 0 0 0-1.63-.94l-.36-2.54a.49.49 0 0 0-.49-.42h-3.84a.49.49 0 0 0-.49.42l-.36 2.54c-.59.23-1.14.55-1.63.94l-2.39-.96a.49.49 0 0 0-.61.22L2.63 8.84a.5.5 0 0 0 .12.64l2.03 1.58c-.05.3-.08.62-.08.94s.03.64.08.94l-2.03 1.58a.5.5 0 0 0-.12.64l1.92 3.32c.12.21.37.3.61.22l2.39-.96c.49.39 1.04.71 1.63.94l.36 2.54c.04.24.24.42.49.42h3.84c.25 0 .45-.18.49-.42l.36-2.54c.59-.23 1.14-.55 1.63-.94l2.39.96c.23.09.49-.01.61-.22l1.92-3.32a.5.5 0 0 0-.12-.64l-2.01-1.56ZM12 15.5a3.5 3.5 0 1 1 0-7 3.5 3.5 0 0 1 0 7Z" fill="currentColor"/>
        </svg>
      </button>
      <section class="panel a4-settings" aria-label="${L('설정','Settings')}">
        <div class="a4-settings-topbar">
          <button type="button" id="collapseSettingsBtn" class="a4-settings-collapse" aria-label="${L('설정 접기','Collapse settings')}" title="${L('설정 접기','Collapse settings')}">
            <svg aria-hidden="true" viewBox="0 0 24 24" fill="none">
              <path d="M19.14 12.94c.04-.3.06-.61.06-.94s-.02-.64-.07-.94l2.03-1.58a.5.5 0 0 0 .12-.64l-1.92-3.32a.5.5 0 0 0-.61-.22l-2.39.96a7.1 7.1 0 0 0-1.63-.94l-.36-2.54a.49.49 0 0 0-.49-.42h-3.84a.49.49 0 0 0-.49.42l-.36 2.54c-.59.23-1.14.55-1.63.94l-2.39-.96a.49.49 0 0 0-.61.22L2.63 8.84a.5.5 0 0 0 .12.64l2.03 1.58c-.05.3-.08.62-.08.94s.03.64.08.94l-2.03 1.58a.5.5 0 0 0-.12.64l1.92 3.32c.12.21.37.3.61.22l2.39-.96c.49.39 1.04.71 1.63.94l.36 2.54c.04.24.24.42.49.42h3.84c.25 0 .45-.18.49-.42l.36-2.54c.59-.23 1.14-.55 1.63-.94l2.39.96c.23.09.49-.01.61-.22l1.92-3.32a.5.5 0 0 0-.12-.64l-2.01-1.56ZM12 15.5a3.5 3.5 0 1 1 0-7 3.5 3.5 0 0 1 0 7Z" fill="currentColor"/>
            </svg>
          </button>
        </div>
        <div class="a4-settings-content">
          <p class="big a4-item-count"><strong>${itemCounts.totalCount}</strong><span class="big-sub">${L('개 항목',' items')}</span></p>
          <p class="selection-count-hint">${L('FSRS 복습 일정은 학습 페이지에서 확인할 수 있어요.','See FSRS due cards on the Study page.')}</p>
          ${state.notice ? `<p class="notice a4-notice" role="status">${esc(state.notice)}</p>` : ''}
          <p id="a4StorageError" class="error" role="alert"${state.error ? '' : ' hidden'}>${esc(state.error)}</p>
          <label class="a4-search-label" for="studySearch">${L('key 또는 value 검색','Search by key or value')}</label>
          <input id="studySearch" class="search" type="search" placeholder="${L('key 또는 value 검색','Search by key or value')}" value="${esc(state.a4Query)}" />
          <section class="a4-settings-group" aria-labelledby="a4ViewSettingsTitle">
            <h2 id="a4ViewSettingsTitle">${L('표시 설정', 'Display')}</h2>
            <button type="button" id="toggleAllValuesBtn" class="btn a4-settings-control a4-values-toggle" aria-pressed="${state.showAllA4Values}" data-study-hint="${L('전체 value 보기','Show all values')}">${state.showAllA4Values ? L('전체 value 가리기','Hide all values') : L('전체 value 보기','Show all values')}</button>
            <button type="button" id="toggleAllHintsBtn" class="btn a4-settings-control" aria-pressed="false" data-study-hint="${L('전체 hint 보기','Show all hints')}">${L('전체 hint 보기','Show all hints')}</button>
            <button type="button" id="a4SwapBtn" class="btn a4-settings-control" aria-pressed="${state.studySettings.swapKeyValue}" data-study-hint="${L('key와 value의 앞뒤를 바꿔서 표시합니다.','Swap which side is shown as key and value.')}">${L('key-value 반전','Swap key-value')}</button>
            ${imageSizeControlHTML('a4')}
          </section>
          <section class="a4-settings-group" aria-labelledby="a4OrderSettingsTitle">
            <h2 id="a4OrderSettingsTitle">${L('학습 순서', 'Study order')}</h2>
            <label class="a4-sort-control">
              <span>${L('정렬 기준', 'Sort by')}</span>
              <select id="a4SortSelect">
                <option value="order"${state.a4Sort === 'order' ? ' selected' : ''}>${L('기본 순서', 'Original order')}</option>
                <option value="priority"${state.a4Sort === 'priority' ? ' selected' : ''}>${L('학습 우선순위', 'Learning priority')}</option>
              </select>
            </label>
            <p class="a4-priority-status" role="status"${state.a4Sort === 'priority' && a4PriorityStatusText() ? '' : ' hidden'}>${esc(a4PriorityStatusText())}</p>
            <button type="button" id="a4ReshuffleBtn" class="btn btn-shuffle a4-settings-control" data-study-hint="${L('순서 섞기','Shuffle order')}">
              <svg aria-hidden="true" viewBox="0 0 24 24" fill="none">
                <path d="M16 3h5v5M4 7h2.5c4.2 0 6.8 10 11 10H21M17 14l4 3-4 3M4 17h2.5c1.8 0 3.2-1.8 4.4-3.8M14 8.8c1.1-1.1 2.2-1.8 3.5-1.8H21M17 4l4 3-4 3" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
              ${L('순서 섞기', 'Shuffle')}
            </button>
          </section>
          <div class="a4-output-actions">
            <button type="button" id="a4PrintBtn" class="btn a4-settings-control a4-print-button" data-study-hint="${L('PDF 인쇄','Print to PDF')}">${L('PDF로 내보내기','Export to PDF')}</button>
          </div>
        </div>
      </section>
      <div id="a4PanelResize" class="resize-handle a4-resize-handle" role="separator" aria-label="${L('설정 영역 너비 조절','Resize settings panel')}" aria-orientation="vertical" aria-valuemin="240" aria-valuemax="900" aria-valuenow="${studyPanelWidth}" title="${L('드래그하거나 방향키로 너비 조절','Drag or use arrow keys to resize')}" tabindex="0"></div>
      <section class="panel a4-view-panel" aria-label="${L('A4 미리보기','A4 preview')}">
        <div id="a4PreviewBox" class="preview-box"></div>
      </section>
    </main>`;
}

function uploadAreaHTML() {
  return `
    <div id="drop" class="panel-upload upload-dropzone home-upload${state.busy ? ' is-busy' : ''}" role="button" tabindex="0" aria-label="${L('엑셀 학습 목록 파일 선택 또는 드래그 앤 드롭', 'Choose an Excel study set or drag and drop it')}">
      <div class="upload-mark" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none">
          <path d="M12 16V4m0 0L7.5 8.5M12 4l4.5 4.5M5 14v4.5A1.5 1.5 0 0 0 6.5 20h11a1.5 1.5 0 0 0 1.5-1.5V14" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </div>
      <div class="upload-copy">
        <p class="upload-title">${state.busy ? L('학습 항목을 읽고 있어요…', 'Reading study items…') : L('엑셀 학습 목록을 여기에 놓거나 클릭해 선택하세요', 'Drop an Excel study set here or click to choose one')}</p>
        <p class="upload-sub">${L('.xlsx · .xlsm · .csv · .txt · .tsv · .apkg(Anki) · .espr(ESPER 공유) 지원 · 프로젝트 이름과 Excel 파일명은 업로드한 파일명으로 고정돼요. 저장 폴더에서 프로젝트 폴더나 파일 이름을 바꾸지 마세요. 같은 이름은 사용할 수 없어요', 'Supports .xlsx · .xlsm · .csv · .txt · .tsv · .apkg (Anki) · .espr (ESPER share) · The project name and workbook filename are fixed to the uploaded filename. Do not rename the project folder or its files; project names must be unique.')}</p>
      </div>
    </div>`;
}

function a4SectionLabel(section, sectionStart) {
  if (!sectionStart) return '';
  return section === 'reviewed'
    ? L('섹션 1 · 복습한 카드 · 기억률 낮은 순', 'Section 1 · Reviewed cards · lowest recall first')
    : section === 'unreviewed'
      ? L('섹션 2 · 아직 복습하지 않은 카드', 'Section 2 · Not yet reviewed')
      : '';
}

function a4SheetHTML(rows, pageNumber, pageCount, isLastPage = false, sectionLabel = '') {
  return `<div class="sheet a4-sheet${isLastPage ? ' a4-last-page' : ''}">
      <div class="sheet-head">
        <div>
          <div class="sheet-title" data-study-edit="title" data-study-hint="${L('제목 수정','Edit title')}" contenteditable="plaintext-only" role="textbox" aria-label="${L('학습 시트 제목 편집','Edit study sheet title')}" spellcheck="false">${esc(a4SheetTitle())}</div>
          <div class="sheet-meta" data-study-edit="info" data-study-hint="${L('정보 수정','Edit info')}" contenteditable="plaintext-only" role="textbox" aria-label="${L('학습 시트 정보 편집','Edit study sheet info')}" spellcheck="false">${esc(a4SheetInfo())}</div>
          ${sectionLabel ? `<div class="sheet-section-label">${esc(sectionLabel)}</div>` : ''}
        </div>
        <div class="sheet-pageno">${pageNumber} / ${pageCount}</div>
      </div>
      <div class="sheet-cols">
        ${[0, 1]
          .map(
            () => `<div class="sheet-column-head">
              <span class="col-no">No.</span>
              <span class="col-key">Key</span>
              <span class="col-value">Value</span>
            </div>`
          )
          .join('')}
      </div>
      <div class="sheet-rows">
        ${[0, 1]
          .map(
            (column) => `<div class="sheet-column">
              ${rows
                .filter((row) => row.column === column)
                .map((row) => {
                  const valueVisible = isA4ValueVisible(row.id);
                  const hintVisible = state.expandedA4Hints.has(row.id);
                  const hasHint = row.hasStudyHint;
                  return `<div class="sheet-row study-sheet-row${hintVisible ? ' study-hint-visible' : ''}" data-study-id="${esc(row.id)}" style="flex-grow:${isLastPage || hintVisible ? 0 : 1}">
                    <span class="col-no">${String(row.index).padStart(3, '0')}</span>
                    <span class="col-key">
                      ${
                        hasHint
                          ? `<button type="button" class="study-key-trigger" data-study-action="hint" data-id="${esc(row.id)}" aria-expanded="${hintVisible}" aria-label="${esc(L(`${stripImageTags(row.key)} hint ${hintVisible ? '숨기기' : '보기'}`, `${hintVisible ? 'Hide' : 'Show'} hint for ${stripImageTags(row.key)}`))}"><span class="sheet-main-key">${richText(row.key)}</span></button>`
                          : `<span class="sheet-main-key">${richText(row.key)}</span>`
                      }
                    </span>
                    <span class="col-value">
                      <button type="button" class="study-value-toggle${valueVisible ? ' is-revealed' : ''}" data-study-action="value" data-id="${esc(row.id)}" aria-pressed="${valueVisible}" aria-label="${esc(L(`${stripImageTags(row.key)} value ${valueVisible ? '가리기' : '보기'}`, `${valueVisible ? 'Hide' : 'Show'} value for ${stripImageTags(row.key)}`))}"><span class="sheet-main-value study-value-text" aria-hidden="${!valueVisible}">${richText(row.value)}</span></button>
                    </span>
                    ${
                      row.hint
                        ? `<div class="study-hint-content">${richText(row.hint)}</div>`
                        : ''
                    }
                  </div>`;
                })
                .join('')}
            </div>`
          )
          .join('')}
      </div>
      <div class="sheet-foot">
        <span>ESPER</span>
        <span>${pageNumber} / ${pageCount}</span>
      </div>
    </div>`;
}

function a4PrintSheetHTML(rows, pageNumber, pageCount, isLastPage = false, sectionLabel = '') {
  return `<div class="sheet a4-sheet a4-print-sheet${isLastPage ? ' a4-last-page' : ''}">
      <div class="sheet-head">
        <div>
          <div class="sheet-title">${esc(a4SheetTitle())}</div>
          <div class="sheet-meta">${esc(a4SheetInfo())}</div>
          ${sectionLabel ? `<div class="sheet-section-label">${esc(sectionLabel)}</div>` : ''}
        </div>
        <div class="sheet-pageno">${pageNumber} / ${pageCount}</div>
      </div>
      <div class="sheet-cols">
        ${[0, 1]
          .map(
            () => `<div class="sheet-column-head">
              <span class="col-no">No.</span>
              <span class="col-key">Key</span>
              <span class="col-value">Value</span>
            </div>`
          )
          .join('')}
      </div>
      <div class="sheet-rows">
        ${[0, 1]
          .map(
            (column) => `<div class="sheet-column">
              ${rows
                .filter((row) => row.column === column)
                .map(
                  (row) => `<div class="sheet-row" style="flex-grow:${isLastPage ? 0 : row.estimatedHeightMm}">
                    <span class="col-no">${String(row.index).padStart(3, '0')}</span>
                    <span class="col-key"><span class="sheet-main-key">${richText(row.key)}</span></span>
                    <span class="col-value">
                      <span class="sheet-main-value${isA4ValueVisible(row.id) ? '' : ' study-print-value-hidden'}">${richText(row.value)}</span>
                      ${
                        row.hint
                          ? `<span class="sheet-hint">${richText(row.hint)}</span>`
                          : ''
                      }
                    </span>
                  </div>`
                )
                .join('')}
            </div>`
          )
          .join('')}
      </div>
      <div class="sheet-foot">
        <span>ESPER</span>
        <span>${pageNumber} / ${pageCount}</span>
      </div>
    </div>`;
}

function refreshLearningStatusRetrievability() {
  if (learningStatusRefreshPromise) {
    learningStatusRefreshRequested = true;
    return learningStatusRefreshPromise;
  }
  const data = state.fsrsData;
  const now = new Date();
  const cards = [...appEl.querySelectorAll('[data-fsrs-retrievability]')].map((element) => ({
    element,
    record: data?.cards.find((item) => item.id === element.dataset.fsrsRetrievability),
  }));
  const dateTimeFormat = new Intl.DateTimeFormat(language === 'ko' ? 'ko-KR' : 'en', {
    dateStyle: 'medium',
    timeStyle: 'short',
  });
  learningStatusRefreshPromise = Promise.all(cards.map(async ({ element, record }) => {
    if (!record || fsrsReviewCategory(record.card) === 'new') return;
    try {
      const probability = await getFsrsRetrievability(record.card, now);
      if (state.page !== 'learning-status' || state.fsrsData !== data || !element.isConnected) return;
      const percent = Math.round(probability * 100);
      element.textContent = `${percent}%`;
      const dialog = appEl.querySelector('.learning-status-dialog');
      if (dialog?.open && dialog.dataset.learningStatusCardId === record.id) {
        dialog.querySelector('[data-learning-status-dialog-recall]').textContent = `${percent}%`;
      }
      const card = element.closest('.learning-status-card');
      if (card) {
        const intensity = Math.max(0, Math.min(1, (probability - 0.5) / 0.5));
        card.style.setProperty('--retention-opacity', (0.08 + intensity * 0.32).toFixed(2));
        card.style.setProperty('--retention-content-opacity', (0.3 + intensity * 0.7).toFixed(2));
        card.dataset.recallLevel = percent >= 85 ? 'high' : percent >= 60 ? 'mid' : 'low';
      }
      element.dataset.lastUpdated = dateTimeFormat.format(now);
      element.title = L(
        `FSRS-7 추정값 · ${dateTimeFormat.format(now)} 기준`,
        `FSRS-7 estimate · as of ${dateTimeFormat.format(now)}`
      );
    } catch (error) {
      if (state.page !== 'learning-status' || !element.isConnected) return;
      element.textContent = L('계산 실패', 'Calculation failed');
      element.title = error instanceof Error ? error.message : String(error);
      const dialog = appEl.querySelector('.learning-status-dialog');
      if (dialog?.open && dialog.dataset.learningStatusCardId === record.id) {
        dialog.querySelector('[data-learning-status-dialog-recall]').textContent = element.textContent;
      }
      console.error('Could not calculate FSRS retrievability.', error);
    }
  })).finally(() => {
    learningStatusRefreshPromise = null;
    if (learningStatusRefreshRequested) {
      learningStatusRefreshRequested = false;
      refreshLearningStatusRetrievability();
    }
  });
  return learningStatusRefreshPromise;
}

/* ---------------- 렌더 ---------------- */
let currentScale = 0;

function renderA4View() {
  const box = document.getElementById('a4PreviewBox');
  if (!box) return;
  const priorityStatus = document.querySelector('.a4-priority-status');
  if (state.a4Sort === 'priority') {
    const message = a4PriorityStatusText();
    if (priorityStatus) {
      priorityStatus.textContent = message;
      priorityStatus.hidden = !message;
    }
    if (state.a4PriorityData !== state.fsrsData || state.a4PriorityPairs !== state.pairs) {
      refreshA4Priority();
    }
  } else if (priorityStatus) {
    priorityStatus.hidden = true;
  }
  const list = a4Pages();
  const totalPages = list.length;
  const priorityPending = state.a4Sort === 'priority' &&
    (state.a4PriorityData !== state.fsrsData || state.a4PriorityPairs !== state.pairs);
  const shuffleButton = document.getElementById('a4ReshuffleBtn');
  const printButton = document.getElementById('a4PrintBtn');
  const allValuesButton = document.getElementById('toggleAllValuesBtn');
  const allHintsButton = document.getElementById('toggleAllHintsBtn');
  if (shuffleButton) shuffleButton.disabled = !list.length;
  if (printButton) {
    printButton.disabled = !totalPages ||
      (state.a4Sort === 'priority' && (priorityPending || Boolean(state.a4PriorityError)));
  }
  if (allValuesButton) {
    allValuesButton.disabled = !totalPages;
    allValuesButton.textContent = state.showAllA4Values ? L('전체 value 가리기', 'Hide all values') : L('전체 value 보기', 'Show all values');
    allValuesButton.setAttribute('aria-pressed', String(state.showAllA4Values));
    allValuesButton.dataset.studyHint = state.showAllA4Values
      ? L('모든 value를 다시 가립니다.', 'Hide all values again.')
      : L('모든 value를 표시합니다.', 'Show all values.');
  }
  if (allHintsButton) {
    syncA4HintControls(allHintsButton);
  }
  if (!totalPages) {
    box.innerHTML = `<div class="empty-box">${state.a4Query.trim() ? L('검색 결과가 없어요.', 'No results found.') : L('학습할 항목이 없어요. 홈에서 프로젝트를 확인해 주세요.', 'There are no items to study. Check your project on the home page.')}</div>`;
    return;
  }

  const width = box.clientWidth || SHEET_W_PX;
  const scale = Math.min(1, Math.max(0.2, width / SHEET_W_PX));
  currentScale = scale;
  const scrollTop = box.scrollTop;
  box.innerHTML = list
    .map(({ rows, section, sectionStart }, i) => {
      const isLastPage = i === list.length - 1;
      const sectionLabel = a4SectionLabel(section, sectionStart);
      return `<div class="frame" style="width:${SHEET_W_PX * scale}px;height:${SHEET_H_PX * scale}px">
        <div class="sheet-scaler" style="transform:scale(${scale});transform-origin:top left">
          ${a4SheetHTML(rows, i + 1, totalPages, isLastPage, sectionLabel)}
        </div>
      </div>`;
    })
    .join('');
  box.scrollTop = scrollTop;
}

function syncStudyValueControls(box) {
  box.querySelectorAll('.study-value-toggle').forEach((button) => {
    const visible = isA4ValueVisible(button.dataset.id);
    const text = button.querySelector('.study-value-text');
    button.classList.toggle('is-revealed', visible);
    button.setAttribute('aria-pressed', String(visible));
    text.setAttribute('aria-hidden', String(!visible));
    button.title = visible
      ? L('이 key의 value를 다시 가립니다.', 'Hide the value for this key again.')
      : L('이 key의 value를 표시합니다.', 'Show the value for this key.');
    button.setAttribute(
      'aria-label',
      (() => { const key = button.closest('.study-sheet-row').querySelector('.sheet-main-key').textContent; return L(`${key} value ${visible ? `가리기: ${text.textContent.trim()}` : '보기'}`, `${visible ? 'Hide' : 'Show'} value for ${key}${visible ? `: ${text.textContent.trim()}` : ''}`); })()
    );
  });
}

function render() {
  if (!['study', 'study-records'].includes(state.page) && studyTimerInterval) {
    window.clearInterval(studyTimerInterval);
    studyTimerInterval = null;
  }
  if (state.page !== 'learning-status' && learningStatusInterval) {
    window.clearInterval(learningStatusInterval);
    learningStatusInterval = null;
  }
  if (state.page !== 'a4' && a4PriorityInterval) {
    window.clearInterval(a4PriorityInterval);
    a4PriorityInterval = null;
  }
  if (state.page !== 'workspace' && workspaceExamGaugeInterval) {
    window.clearInterval(workspaceExamGaugeInterval);
    workspaceExamGaugeInterval = null;
  }
  const pageHTML =
    `${state.page === 'home'
      ? startHTML()
      : state.page === 'workspace'
        ? workspaceHTML()
        : state.page === 'edit'
          ? editHTML()
          : state.page === 'study'
            ? studyHTML()
            : state.page === 'study-records'
              ? studyRecordsHTML()
              : state.page === 'learning-status'
                ? learningStatusHTML()
                : state.page === 'project-setup'
                  ? projectSetupHTML()
            : a4ViewHTML()}${state.languageChoicePending ? languageChoiceHTML() : ''}${state.anki ? ankiImportHTML() : ''}`;
  appEl.innerHTML = pageHTML;
  renderProcessingIndicator();
  applyImageSizes();
  appEl.classList.toggle('home-page', state.page === 'home');
  appEl.classList.toggle('workspace-page', state.page === 'workspace');
  appEl.classList.toggle('a4-view-page', state.page === 'a4');
  appEl.classList.toggle('study-page', state.page === 'study');
  appEl.classList.toggle('study-records-page', state.page === 'study-records');
  appEl.classList.toggle('learning-status-page', state.page === 'learning-status');
  appEl.classList.toggle('edit-page', state.page === 'edit');
  attach();
  refreshExamSummary();
  if (state.page === 'workspace') {
    refreshWorkspaceExamGauges();
    if (!workspaceExamGaugeInterval) {
      workspaceExamGaugeInterval = window.setInterval(refreshWorkspaceExamGauges, 3_600_000);
    }

  }
  attachAnkiDialog();
  hydrateRenderedImages();
  if (state.page === 'study') requestAnimationFrame(fitStudyCardText);
  if (state.page === 'learning-status') {
    refreshLearningStatusRetrievability();
    if (!learningStatusInterval) {
      learningStatusInterval = window.setInterval(refreshLearningStatusRetrievability, 3_600_000);
    }
  }
  if (state.page === 'a4') {
    renderA4View();
    if (!a4PriorityInterval) {
      a4PriorityInterval = window.setInterval(() => {
        if (state.a4Sort !== 'priority') return;
        state.a4PriorityData = null;
        state.a4PriorityPairs = null;
        refreshA4Priority();
      }, 3_600_000);
    }
  }
  const a4PreviewBox = document.getElementById('a4PreviewBox');
  if (a4PreviewBox) resizeObserver.observe(a4PreviewBox);
}

function processingIndicatorMessage() {
  return state.anki?.busy
    ? state.anki.progress || L('가져오는 중…', 'Processing…')
    : state.processingMessage ||
      (state.studySaving
        ? L('저장 중…', 'Saving…')
        : processingProjectIds.size
          ? L('프로젝트 처리 중…', 'Processing project…')
          : state.studyRecordsLoading
            ? L('학습 기록을 불러오는 중…', 'Loading study records…')
            : state.busy
              ? L('처리 중…', 'Processing…')
              : '');
}

function renderProcessingIndicator() {
  const message = processingIndicatorMessage();
  let indicator = document.getElementById('processingIndicator');
  if (!message) {
    indicator?.remove();
    appEl.inert = false;
    return;
  }
  appEl.inert = true;
  if (!indicator) {
    document.body.insertAdjacentHTML('beforeend', `<div class="processing-indicator" id="processingIndicator" role="status" aria-live="polite" aria-atomic="true">
      <div class="processing-indicator-content">
        <p class="processing-indicator-message"><span data-processing-message></span><span class="processing-indicator-dots" aria-hidden="true"><i></i><i></i><i></i></span></p>
        <p class="processing-indicator-warning">${L('작업이 완료될 때까지 새로고침하거나 페이지를 나가지 마세요.', 'Do not refresh or leave this page until the operation is complete.')}</p>
      </div>
    </div>`);
    indicator = document.getElementById('processingIndicator');
  }
  indicator.querySelector('[data-processing-message]').textContent = message;
}

function mountSupportWidget() {
  document.querySelector('.support-widget')?.remove();
  document.body.insertAdjacentHTML('beforeend', supportWidgetHTML());
  const supportDialog = document.getElementById('supportDialog');
  document.getElementById('supportOpenBtn').addEventListener('click', () => supportDialog.showModal());
  document.getElementById('supportCloseBtn').addEventListener('click', () => supportDialog.close());
  supportDialog.addEventListener('click', (event) => {
    if (event.target === supportDialog) supportDialog.close();
  });
}

function supportWidgetHTML() {
  return `
    <div class="support-widget">
      <button type="button" class="support-trigger" id="supportOpenBtn" aria-haspopup="dialog" aria-controls="supportDialog">
        <span aria-hidden="true">☕</span>
        <span>${L('도움이 되었다면', 'Support ESPER')}</span>
      </button>
      <dialog class="support-dialog" id="supportDialog" aria-labelledby="supportTitle" aria-describedby="supportDescription">
        <div class="support-dialog-card">
          <button type="button" class="support-close" id="supportCloseBtn" aria-label="${L('후원 안내 닫기', 'Close support dialog')}">×</button>
          <p class="support-eyebrow">ESPER SUPPORT</p>
          <h2 id="supportTitle">${L('ESPER가 도움이 되었나요?', 'Is ESPER helping you?')}</h2>
          <p></p>
          <img class="support-qr" src="${SUPPORT_QR_SRC[language]}" alt="${L('자발적 후원을 위한 QR 코드', 'QR code for voluntary support')}" />
          <p id="supportDescription">${L('후원하신 금액은 전액 개발자의 식비와 Copilot 구독비로 탕진됩니다.', 'Every donation goes directly toward keeping the developer fed and his Copilot running.')}</p>
        </div>
      </dialog>
    </div>`;
}

/* ---------------- 동작 ---------------- */

async function handleA4Print() {
  hideStudyTooltip();
  if (state.a4Sort === 'priority' &&
    (state.a4PriorityData !== state.fsrsData || state.a4PriorityPairs !== state.pairs)) {
    await refreshA4Priority();
  }
  if (state.a4Sort === 'priority' && state.a4PriorityError) return;
  const list = a4Pages();
  const totalPages = list.length;
  if (!totalPages) return;
  printEl.innerHTML = list
    .map(({ rows, section, sectionStart }, i) => {
      const isLastPage = i === list.length - 1;
      return a4PrintSheetHTML(rows, i + 1, totalPages, isLastPage, a4SectionLabel(section, sectionStart));
    })
    .join('');
  hydrateRenderedImages();
  await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
  await waitForImages(printEl);
  window.addEventListener('afterprint', () => {
    printEl.replaceChildren();
  }, { once: true });
  window.print();
}

function localDateForExport() {
  const now = new Date();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${now.getFullYear()}-${month}-${day}`;
}

function markStudyCardSeen(session) {
  const pair = state.pairs.find((item) => item.id === session.cards[session.index]);
  if (pair) session.seenIds.add(pair.id);
}

function enterStudyCard(session) {
  session.flipped = false;
  session.hintOpen = false;
  session.enterEffect = session.enterEffect || '';
  session.showBackHint = session.settings.showHintOnBack;
  markStudyCardSeen(session);
  session.previews = [];
  session.previewBase = Date.now();
  const cardId = session.cards[session.index];
  const record = fsrsRecordForCard(cardId);
  if (record) {
    previewFsrsCard(record.card, new Date(session.previewBase), projectPlan().desiredRetention, projectPlan().examDate).then((previews) => {
      if (state.study !== session || session.cards[session.index] !== cardId) return;
      session.previews = previews;
      render();
    }).catch((error) => {
      if (state.study !== session || session.cards[session.index] !== cardId) return;
      state.studyError = L(
        `FSRS 예상 일정을 계산하지 못했어요: ${error instanceof Error ? error.message : String(error)}`,
        `Could not calculate FSRS schedule previews: ${error instanceof Error ? error.message : String(error)}`
      );
      render();
    });
  }
}

function startStudy() {
  const pairIds = new Set(state.pairs.map((pair) => pair.id));
  const now = Date.now();
  const dueCards = (state.fsrsData?.cards || [])
    .filter((record) =>
      record.archivedAt === null &&
      pairIds.has(record.id) &&
      (record.card.state === FsrsState.New || Date.parse(record.card.dueAt) <= now)
    )
    .map((record) => record.id);
  const cards = shuffle(dueCards);
  state.studyError = '';
  state.pendingReview = null;
  state.studyKeyCapture = '';
  state.studyKeyError = '';
  state.study = cards.length
    ? {
        cards,
        index: 0,
        started: false,
        startedAt: 0,
        history: [],
        seenIds: new Set(),
        flipped: false,
        hintOpen: false,
        showBackHint: state.studySettings.showHintOnBack,
        settings: normalizeStudySettings(state.studySettings),
      }
    : null;
  if (state.study) randomizeStudyPairDirections(state.study);
  state.page = 'study';
  render();
}

function beginStudySession() {
  const session = state.study;
  if (!session || session.started || state.studySaving) return;
  session.started = true;
  session.startedAt = Date.now();
  enterStudyCard(session);
  render();
}

function advanceStudy(rating) {
  const session = state.study;
  if (!session || state.studySaving) return;
  const pair = studyCurrentPair();
  if (!pair) return;
  session.history.push({ id: pair.id, rating });
  session.index += 1;
  session.enterEffect = 'review';
  if (session.index >= session.cards.length) {
    finishStudy(false);
    return;
  }
  enterStudyCard(session);
  render();
}

async function rateStudyCard(rating) {
  const session = state.study;
  const pair = studyCurrentPair();
  if (!session || !pair || state.studySaving || state.pendingReview) return;
  state.studySaving = true;
  state.studyError = '';
  render();
  try {
    const result = await reviewFsrsCard(state.project, state.fsrsData, pair.id, rating, new Date(), projectPlan().desiredRetention, projectPlan().examDate);
    state.pendingReview = { data: result.data, rating };
    await saveFsrsData(state.project, result.data);
    state.fsrsData = result.data;
    state.pendingReview = null;
    state.studySaving = false;
    advanceStudy(rating);
  } catch (error) {
    state.studySaving = false;
    state.studyError = L(
      `복습 결과를 저장하지 못했어요. 저장을 다시 시도할 때까지 이 화면을 닫지 마세요. ${error instanceof Error ? error.message : String(error)}`,
      `The review could not be saved. Keep this page open until you retry the save. ${error instanceof Error ? error.message : String(error)}`
    );
    render();
  }
}

async function retryPendingReviewSave() {
  if (!state.pendingReview || !state.project || state.studySaving) return;
  state.studySaving = true;
  state.studyError = '';
  render();
  try {
    await saveFsrsData(state.project, state.pendingReview.data);
    const rating = state.pendingReview.rating;
    state.fsrsData = state.pendingReview.data;
    state.pendingReview = null;
    state.studySaving = false;
    advanceStudy(rating);
  } catch (error) {
    state.studySaving = false;
    state.studyError = L(
      `복습 결과를 아직 저장하지 못했어요. 폴더 권한과 여유 공간을 확인한 뒤 다시 시도해 주세요. ${error instanceof Error ? error.message : String(error)}`,
      `The review is still unsaved. Check folder access and free space, then retry. ${error instanceof Error ? error.message : String(error)}`
    );
    render();
  }
}

async function goToPreviousStudyCard() {
  const session = state.study;
  if (!session || !session.history.length || state.studySaving) return;
  const previous = session.history.pop();
  const pair = state.pairs.find((item) => item.id === previous.id);
  if (!pair) return;
  session.index = session.cards.indexOf(previous.id);
  session.enterEffect = 'back';
  enterStudyCard(session);
  render();
}

async function showStudySpeech() {
  const current = studyCurrentPair();
  const pair = current && studyShownPair(current);
  if (!pair || !stripImageTags(pair.key).trim()) return;
  if (!('speechSynthesis' in window) || !('SpeechSynthesisUtterance' in window)) {
    state.studyError = L('이 브라우저에서는 음성 읽기를 지원하지 않아요.', 'This browser does not support speech.');
    render();
    return;
  }
  let availableVoices = window.speechSynthesis.getVoices();
  if (!availableVoices.length) {
    availableVoices = await new Promise((resolve) => {
      const finish = () => {
        window.speechSynthesis.removeEventListener('voiceschanged', onVoicesChanged);
        window.clearTimeout(timeout);
        resolve(window.speechSynthesis.getVoices());
      };
      const onVoicesChanged = () => {
        if (window.speechSynthesis.getVoices().length) finish();
      };
      const timeout = window.setTimeout(finish, 1200);
      window.speechSynthesis.addEventListener('voiceschanged', onVoicesChanged);
      window.speechSynthesis.getVoices();
    });
  }
  if (state.page !== 'study' || studyCurrentPair()?.id !== pair.id) return;
  const preferredVoice = availableVoices.find((voice) =>
    voice.name.trim().toLowerCase() === 'samantha' &&
    voice.lang.toLowerCase() === 'en-us'
  );
  if (!preferredVoice) {
    state.studyError = L('고정 음성인 Samantha (영어·미국)를 사용할 수 없어요. 기기의 영어 음성 설정을 확인해 주세요.', 'The fixed voice Samantha (English, US) is not available. Check your device\'s English voice settings.');
    render();
    return;
  }
  window.speechSynthesis.cancel();
  const utterance = new SpeechSynthesisUtterance(stripImageTags(pair.key));
  utterance.lang = preferredVoice.lang;
  utterance.voice = preferredVoice;
  utterance.rate = 0.9;
  utterance.pitch = 1;
  utterance.volume = 1;
  window.speechSynthesis.speak(utterance);
}

function normalizeStudyPair(pair, index = 0) {
  const normalized = {
    ...pair,
    id: pair.id || `k${index}`,
    key: String(pair.key ?? pair.word ?? ''),
    value: String(pair.value ?? pair.meaning ?? ''),
    hint: String(pair.hint ?? pair.example ?? ''),
  };
  delete normalized.known;
  return normalized;
}

async function saveStudySettings() {
  try {
    await saveStudySettingsToStore(state.studySettings);
    if (state.study) {
      state.study.settings = normalizeStudySettings(state.studySettings);
    }
    state.studyKeyError = '';
    return true;
  } catch (error) {
    state.studyKeyError = L(`설정을 저장하지 못했어요: ${error instanceof Error ? error.message : String(error)}`, `Could not save settings: ${error instanceof Error ? error.message : String(error)}`);
    return false;
  }
}

async function finishStudy(askForConfirmation = false, returnPage = 'study') {
  const session = state.study;
  if (state.pendingReview) {
    state.studyError = L(
      '저장되지 않은 복습 결과가 있어요. 복습 결과를 저장한 뒤 회차를 종료해 주세요.',
      'A review is still unsaved. Save it before ending the round.'
    );
    render();
    return;
  }
  if (state.studyFullscreen) {
    state.studyFullscreen = false;
    if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
  }
  if (!session) {
    state.page = returnPage;
    render();
    return;
  }
  if (!session.started) {
    state.study = null;
    state.page = returnPage;
    render();
    return;
  }
  if (askForConfirmation && !window.confirm(L('이번 학습 회차를 종료하고 기록을 저장할까요?', 'End this study round and save its record?'))) return;
  if (state.studySaving) return;
  state.studySaving = true;
  state.studyError = '';
  render();
  try {
    if (!session.pendingSave) {
      const date = localDateForExport();
      const elapsed = Date.now() - session.startedAt;
      const keys = [...session.seenIds]
        .map((id) => state.pairs.find((pair) => pair.id === id)?.key)
        .filter((key) => Boolean(key));
      const recordFile = await readStudyRecordFile(state.project.name);
      const recordWorkbook = await exportStudyRecordWorkbook(recordFile, keys, {
        date,
        itemCount: keys.length,
        duration: formatStudyTime(elapsed),
      });
      session.pendingSave = { recordWorkbook };
    }
    await writeStudyRecordFile(state.project.name, session.pendingSave.recordWorkbook);
    if (!(await save())) throw new Error(state.error || L('프로젝트 정보를 저장하지 못했어요.', 'Could not save project information.'));
    if (window.speechSynthesis) window.speechSynthesis.cancel();
    state.study = null;
    state.studySaving = false;
    state.studyError = '';
    if (returnPage === 'study') {
      startStudy();
      return;
    }
    state.page = returnPage;
    render();
  } catch (error) {
    state.studySaving = false;
    state.studyError = L(`학습 기록을 저장하지 못했어요. 다시 시도해 주세요. ${error instanceof Error ? error.message : String(error)}`, `Could not save the study record. Please try again. ${error instanceof Error ? error.message : String(error)}`);
    render();
  }
}

async function openStudyRecords() {
  state.page = 'study-records';
  state.studyRecords = null;
  state.studyRecordsLoading = true;
  state.studyRecordsError = '';
  render();
  try {
    const file = await readStudyRecordFile(state.project.name);
    state.studyRecords = file
      ? await readStudyRecordWorkbook(file)
      : { keys: [], sessions: [] };
  } catch (error) {
    state.studyRecordsError = L(`기록을 불러오지 못했어요: ${error instanceof Error ? error.message : String(error)}`, `Could not load records: ${error instanceof Error ? error.message : String(error)}`);
  } finally {
    state.studyRecordsLoading = false;
    render();
  }
}

async function saveProjectPlan(plan) {
  const project = state.project;
  const previous = project.plan;
  project.plan = plan;
  try {
    await saveProject(project);
  } catch (error) {
    project.plan = previous;
    throw error;
  }
  const listed = state.projects.find((item) => item.id === project.id);
  if (listed && listed !== project) listed.plan = plan;
}

function projectSetupHTML() {
  const percent = Math.round((state.project?.plan ? projectPlan().desiredRetention : DEFAULT_DESIRED_RETENTION) * 100);
  return `
    ${headerHTML()}
    <main class="learning-status-screen project-setup-screen">
      <header class="learning-status-heading">
        <p class="study-ready-eyebrow">${L('새 프로젝트', 'NEW PROJECT')}</p>
        <p class="learning-status-explainer">${L('학습을 시작하기 전에 시험 날짜와 목표 기억률을 정해 주세요. 복습 일정이 이 기준에 맞춰 계산돼요.', 'Before you start, set your exam date and target retention. Review dates are scheduled around them.')}</p>
      </header>
      <section class="exam-plan-panel project-setup-panel">
        <div class="exam-plan-controls">
          <div class="exam-plan-field">
            <span>${L('시험 날짜', 'Exam date')}</span>
            ${datePickerHTML(
              'data-setup-date',
              state.project?.plan ? projectPlan().examDate : '',
              localDateValue(new Date())
            )}
          </div>
          <label class="exam-plan-field">
            <span>${L('목표 기억률', 'Target retention')} <strong data-setup-label>${percent}%</strong></span>
            <input type="range" min="${Math.round(MIN_DESIRED_RETENTION * 100)}" max="${Math.round(MAX_DESIRED_RETENTION * 100)}" step="1" data-setup-retention value="${percent}">
          </label>
        </div>
        ${retentionWarningHTML(percent, 'data-setup-note')}
        <p class="exam-plan-line" data-setup-error role="alert"></p>
        <div class="exam-plan-actions">
          <button type="button" class="btn" data-setup-confirm>${L('설정하고 시작하기', 'Save and start')}</button>
        </div>
      </section>
    </main>`;
}

function attachProjectSetupPage() {
  document.getElementById('homeBtn')?.addEventListener('click', () => {
    state.page = 'workspace';
    render();
  });
  const dateInput = appEl.querySelector('[data-setup-date]');
  const slider = appEl.querySelector('[data-setup-retention]');
  const errorEl = appEl.querySelector('[data-setup-error]');
  slider.addEventListener('input', () => {
    const retention = Number(slider.value) / 100;
    appEl.querySelector('[data-setup-label]').textContent = `${Math.round(retention * 100)}%`;
    appEl.querySelector('[data-setup-note]').hidden = Math.round(retention * 100) <= RETENTION_WARNING_THRESHOLD;
  });
  appEl.querySelector('[data-setup-confirm]').addEventListener('click', async () => {
    if (!dateInput.value) {
      errorEl.textContent = L('시험 날짜를 선택해 주세요.', 'Please choose an exam date.');
      appEl.querySelector('[data-date-picker-trigger]')?.focus();
      return;
    }
    try {
      await saveProjectPlan({
        examDate: dateInput.value,
        desiredRetention: normalizeDesiredRetention(Number(slider.value) / 100),
      });
    } catch (error) {
      errorEl.textContent = L('설정을 저장하지 못했어요: ', 'Could not save settings: ') + (error instanceof Error ? error.message : String(error));
      return;
    }
    state.page = 'study';
    render();
  });
}

function attachLearningStatusPage() {
  document.getElementById('homeBtn')?.addEventListener('click', () => {
    state.page = 'workspace';
    state.error = '';
    state.notice = '';
    render();
  });
  appEl.querySelector('.exam-plan-panel')?.addEventListener('toggle', (event) => {
    state.planPanelOpen = event.currentTarget.open;
  });
  const status = appEl.querySelector('[data-plan-status]');
  const slider = appEl.querySelector('[data-plan-setting="desiredRetention"]');
  const dateInput = appEl.querySelector('[data-plan-setting="examDate"]');
  const dateTrigger = appEl.querySelector('[data-date-picker-trigger]');
  const editButton = appEl.querySelector('[data-plan-edit]');
  const saveButton = appEl.querySelector('[data-plan-save]');
  const cancelButton = appEl.querySelector('[data-plan-cancel]');
  slider?.addEventListener('input', () => {
    const retention = Number(slider.value) / 100;
    appEl.querySelector('[data-retention-label]').textContent = `${Math.round(retention * 100)}%`;
    appEl.querySelector('[data-retention-note]').hidden = Math.round(retention * 100) <= RETENTION_WARNING_THRESHOLD;
  });
  editButton?.addEventListener('click', () => {
    slider.disabled = false;
    dateInput.disabled = false;
    dateTrigger.disabled = false;
    editButton.hidden = true;
    saveButton.hidden = false;
    cancelButton.hidden = false;
  });
  cancelButton?.addEventListener('click', () => {
    state.learningStatusPlanMessage = '';
    render();
  });
  saveButton?.addEventListener('click', async () => {
    saveButton.disabled = true;
    const plan = {
      examDate: dateInput.value,
      desiredRetention: normalizeDesiredRetention(Number(slider.value) / 100),
    };
    try {
      const result = await rescheduleFsrsData(state.project, state.fsrsData, plan.desiredRetention, plan.examDate);
      await saveProjectPlan(plan);
      await saveFsrsData(state.project, result.data);
      state.fsrsData = result.data;
      state.learningStatusPlanMessage = L(`저장했어요. 복습 중인 카드 ${result.changed}개의 다음 복습일을 새 기준으로 변경했어요.`, `Saved. Updated the next review date of ${result.changed} card(s) in review to the new settings.`);
    } catch (error) {
      state.learningStatusPlanMessage = L('설정을 적용하지 못했어요: ', 'Could not apply the settings: ') + (error instanceof Error ? error.message : String(error));
      console.error('Could not reschedule FSRS cards.', error);
    }
    render();
  });
  if (status && state.learningStatusPlanMessage) status.textContent = state.learningStatusPlanMessage;

  const search = appEl.querySelector('[data-learning-status-search]');
  const empty = appEl.querySelector('[data-learning-status-empty]');
  const cardList = appEl.querySelector('.learning-status-list');
  const detailDialog = appEl.querySelector('.learning-status-dialog');
  const refreshCardRecall = async (card) => {
    const element = card.querySelector('[data-fsrs-retrievability]');
    if (!element) return;
    if (card.dataset.learningStatusCard === 'new') {
      element.textContent = L('학습 전', 'Not studied');
      return;
    }
    const record = state.fsrsData?.cards.find((item) =>
      item.id === element.dataset.fsrsRetrievability && item.archivedAt === null
    );
    if (!record) return;

    const requestId = String(Number(element.dataset.probabilityRequest || 0) + 1);
    element.dataset.probabilityRequest = requestId;
    element.textContent = '…';
    const now = new Date();
    try {
      const probability = await getFsrsRetrievability(record.card, now);
      if (!element.isConnected || element.dataset.probabilityRequest !== requestId) return;
      const percent = Math.round(probability * 100);
      element.textContent = `${percent}%`;
      const intensity = Math.max(0, Math.min(1, (probability - 0.5) / 0.5));
      card.style.setProperty('--retention-opacity', (0.08 + intensity * 0.32).toFixed(2));
      card.style.setProperty('--retention-content-opacity', (0.3 + intensity * 0.7).toFixed(2));
      const dateTime = new Intl.DateTimeFormat(language === 'ko' ? 'ko-KR' : 'en', {
        dateStyle: 'medium',
        timeStyle: 'short',
      }).format(now);
      element.title = L(`현재 기억 확률 · ${dateTime} 기준`, `Current recall probability · as of ${dateTime}`);
    } catch (error) {
      if (!element.isConnected || element.dataset.probabilityRequest !== requestId) return;
      element.textContent = L('계산 실패', 'Unavailable');
      console.error('Could not calculate hovered card retrievability.', error);
    }
  };
  cardList?.addEventListener('pointerover', (event) => {
    const card = event.target.closest('.learning-status-card');
    if (!card || card.contains(event.relatedTarget)) return;
    refreshCardRecall(card);
  });
  cardList?.addEventListener('focusin', (event) => {
    const card = event.target.closest('.learning-status-card');
    if (card) refreshCardRecall(card);
  });
  cardList?.addEventListener('click', (event) => {
    const trigger = event.target.closest('[data-learning-status-open]');
    if (!trigger) return;
    const cardId = trigger.dataset.learningStatusOpen;
    const template = [...cardList.querySelectorAll('template[data-learning-status-detail]')]
      .find((item) => item.dataset.learningStatusDetail === cardId);
    if (!template || !detailDialog) {
      console.error('Could not find learning status details for card.', cardId);
      return;
    }
    detailDialog.innerHTML = template.innerHTML;
    detailDialog.dataset.learningStatusCardId = cardId;
    const recall = trigger.querySelector('[data-fsrs-retrievability]');
    const dialogRecall = detailDialog.querySelector('[data-learning-status-dialog-recall]');
    if (recall && dialogRecall) dialogRecall.textContent = recall.textContent;
    detailDialog.showModal();
  });
  detailDialog?.addEventListener('click', (event) => {
    if (event.target === detailDialog || event.target.closest('[data-learning-status-close]')) {
      detailDialog.close();
    }
  });
  const applyFilters = () => {
    const query = state.learningStatusQuery.trim().toLocaleLowerCase();
    let visibleCount = 0;
    appEl.querySelectorAll('[data-learning-status-card]').forEach((card) => {
      const matchesFilter = state.learningStatusFilter === 'all' ||
        card.dataset.learningStatusCard === state.learningStatusFilter;
      const matchesQuery = !query || card.dataset.search.includes(query);
      card.hidden = !matchesFilter || !matchesQuery;
      if (!card.hidden) visibleCount += 1;
    });
    if (empty) empty.hidden = visibleCount > 0;
  };
  search?.addEventListener('input', () => {
    state.learningStatusQuery = search.value;
    applyFilters();
  });
  appEl.querySelectorAll('[data-learning-status-filter]').forEach((button) => {
    button.addEventListener('click', () => {
      state.learningStatusFilter = button.dataset.learningStatusFilter;
      render();
    });
  });
  appEl.querySelector('[data-learning-status-swap]')?.addEventListener('click', () => {
    state.learningStatusShowValue = !state.learningStatusShowValue;
    render();
  });
  applyFilters();
}

async function routeToProjectDestination(destination) {
  if (destination === 'a4') {
    state.page = 'a4';
    render();
  } else if (destination === 'progress') {
    await openStudyRecords();
  } else if (destination === 'status') {
    state.page = 'learning-status';
    render();
  } else {
    state.page = 'study';
    render();
  }
}

async function navigateGlobalPage(destination) {
  if (destination === 'workspace') {
    if (state.page === 'edit' && state.editDirty &&
      !window.confirm(L('저장하지 않은 변경사항을 버릴까요?', 'Discard your unsaved changes?'))) return;
    if (state.study?.started) {
      await finishStudy(true, 'workspace');
      return;
    }
    state.page = 'workspace';
    render();
    return;
  }
  if (!state.project) return;
  if (destination === 'study') {
    if (!state.study) startStudy();
    else {
      state.page = 'study';
      render();
    }
  } else if (destination === 'a4') {
    state.page = 'a4';
    render();
  } else if (destination === 'progress') {
    await openStudyRecords();
  } else if (destination === 'status') {
    state.page = 'learning-status';
    render();
  }
}

function attachStudyRecords() {
  if (studyTimerInterval) {
    window.clearInterval(studyTimerInterval);
    studyTimerInterval = null;
  }
  document.getElementById('homeBtn')?.addEventListener('click', () => finishStudy(true, 'workspace'));
  appEl.querySelectorAll('[data-study-records-details]').forEach((details) => {
    details.addEventListener('toggle', () => {
      state.studyRecordsOpen[details.dataset.studyRecordsDetails] = details.open;
    });
  });
  appEl.onclick = (event) => {
    const calMove = event.target.closest('[data-study-cal-move]');
    if (calMove) {
      const now = new Date();
      const [y, m] = (state.studyCalendarMonth || `${now.getFullYear()}-${now.getMonth() + 1}`).split('-').map(Number);
      const next = new Date(y, m - 1 + Number(calMove.dataset.studyCalMove), 1);
      state.studyCalendarMonth = `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, '0')}`;
      render();
    } else if (event.target.closest('[data-study-records-retry]')) {
      openStudyRecords();
    }
  };
}

function studyShortcutKey(event) {
  if (event.key === ' ') return 'Space';
  if (['ArrowDown', 'ArrowRight', 'ArrowLeft', 'Enter', 'Escape'].includes(event.key)) return event.key;
  if (event.key.length === 1 && !event.ctrlKey && !event.altKey && !event.metaKey) return event.key.toLocaleLowerCase();
  return '';
}

async function handleStudyKeydown(event) {
  if (state.page !== 'study' || event.repeat || event.isComposing) return;
  if (state.studySaving) return;
  const key = studyShortcutKey(event);
  if (!key) return;
  if (state.studyKeyCapture) {
    event.preventDefault();
    const action = state.studyKeyCapture;
    const keyId = shortcutKeyId(key);
    const collision = STUDY_ACTIONS.find(([other]) =>
      other !== action && shortcutKeyId(state.studySettings.shortcuts[other]) === keyId
    );
    if (collision) {
      state.studyKeyError = L(`이미 "${collision[1]}" 동작에 지정된 키예요. 다른 키를 선택해 주세요.`, `That key is already assigned to "${studyActionLabel(collision)}". Please choose another key.`);
      render();
      return;
    }
    state.studySettings.shortcuts[action] = key;
    state.studyKeyCapture = '';
    await saveStudySettings();
    render();
    return;
  }
  const target = event.target;
  if (target instanceof HTMLElement && (
    target.isContentEditable ||
    target.matches('input, textarea, select, [role="dialog"]') ||
    target.closest('dialog[open]')
  )) return;
  if (!state.study?.started) {
    if (shortcutKeyId(state.studySettings.shortcuts.exit) === shortcutKeyId(key)) {
      event.preventDefault();
      finishStudy(true);
    }
    return;
  }
  const action = Object.entries(state.studySettings.shortcuts)
    .find(([, assignedKey]) => shortcutKeyId(assignedKey) === shortcutKeyId(key))?.[0];
  if (!action) return;
  event.preventDefault();
  const ratingByAction = {
    again: FsrsRating.Again,
    hard: FsrsRating.Hard,
    good: FsrsRating.Good,
    easy: FsrsRating.Easy,
  };
  if (Object.hasOwn(ratingByAction, action)) {
    if (state.study?.flipped) await rateStudyCard(ratingByAction[action]);
  } else if (action === 'flip') {
    if (state.study) state.study.flipped = !state.study.flipped;
    render();
  } else if (action === 'hint') {
    if (state.study && !state.study.flipped) state.study.hintOpen = !state.study.hintOpen;
    render();
  } else if (action === 'previous') {
    await goToPreviousStudyCard();
  } else if (action === 'speak') {
    await showStudySpeech();
  } else if (action === 'exit') {
    if (state.studyFullscreen) await toggleStudyFullscreen(false);
    else finishStudy(true);
  }
}

async function toggleStudyFullscreen(force) {
  const next = typeof force === 'boolean' ? force : !state.studyFullscreen;
  state.studyFullscreen = next;
  try {
    if (next && !document.fullscreenElement) await document.documentElement.requestFullscreen();
    else if (!next && document.fullscreenElement) await document.exitFullscreen();
  } catch {
    // CSS 전체화면 레이아웃만 사용
  }
  render();
}

document.addEventListener('fullscreenchange', () => {
  if (!document.fullscreenElement && state.studyFullscreen) {
    state.studyFullscreen = false;
    render();
  }
});

function attachStudy() {
  if (state.study?.enterEffect) state.study.enterEffect = '';
  if (studyTimerInterval) window.clearInterval(studyTimerInterval);
  const timer = document.getElementById('studyTimer');
  const updateTimer = () => {
    if (timer && state.study) {
      timer.textContent = formatStudyTime(Date.now() - state.study.startedAt);
    }
  };
  if (state.study?.started) {
    studyTimerInterval = window.setInterval(updateTimer, 1000);
  }

  appEl.onclick = async (event) => {
    const ratingButton = event.target.closest('[data-study-rating]');
    if (ratingButton) {
      event.preventDefault();
      event.stopPropagation();
      await rateStudyCard(Number(ratingButton.dataset.studyRating));
      return;
    }
    const actionButton = event.target.closest('[data-study-action]');
    if (actionButton) {
      event.preventDefault();
      event.stopPropagation();
      const action = actionButton.dataset.studyAction;
      if (action === 'exit') await finishStudy(true);
      else if (action === 'start') beginStudySession();
      else if (action === 'restart') startStudy();
      else if (action === 'fullscreen') await toggleStudyFullscreen();
      else if (action === 'records') {
        await openStudyRecords();
      }
      else if (action === 'retry-review') await retryPendingReviewSave();
      else if (action === 'speak') await showStudySpeech();
      else if (action === 'hint' && state.study) {
        state.study.hintOpen = !state.study.hintOpen;
        render();
      } else if (action === 'back-hint' && state.study) {
        state.study.showBackHint = true;
        render();
      }
      return;
    }
    if (event.target.closest('[data-study-flip]') && state.study) {
      state.study.flipped = !state.study.flipped;
      render();
      return;
    }
    const bindButton = event.target.closest('[data-study-bind]');
    if (bindButton) {
      state.studyKeyCapture = bindButton.dataset.studyBind;
      state.studyKeyError = '';
      render();
      return;
    }
    if (event.target.closest('[data-study-reset-shortcuts]')) {
      state.studySettings.shortcuts = { ...DEFAULT_STUDY_SETTINGS.shortcuts };
      state.studyKeyCapture = '';
      await saveStudySettings();
      render();
    }
  };

  const setStudySidebarCollapsed = (collapsed) => {
    studySidebarCollapsed = collapsed;
    const sidebarLayout = document.querySelector('.study-workspace');
    sidebarLayout?.classList.toggle('study-sidebar-collapsed', collapsed);
    const reopenButton = document.getElementById('reopenStudySidebarBtn');
    if (reopenButton) reopenButton.hidden = !collapsed;
    requestAnimationFrame(fitStudyCardText);
    requestAnimationFrame(() => {
      document
        .getElementById(collapsed ? 'reopenStudySidebarBtn' : 'collapseStudySidebarBtn')
        ?.focus();
    });
  };
  document.getElementById('collapseStudySidebarBtn')?.addEventListener('click', () => {
    setStudySidebarCollapsed(true);
  });
  document.getElementById('reopenStudySidebarBtn')?.addEventListener('click', () => {
    setStudySidebarCollapsed(false);
  });

  appEl.querySelector('.study-shortcuts')?.addEventListener('toggle', (event) => {
    state.studyShortcutsOpen = event.currentTarget.open;
  });

  appEl.querySelectorAll('[data-study-setting]').forEach((input) => {
    input.addEventListener('change', async () => {
      const setting = input.dataset.studySetting;
      state.studySettings[setting] = input.checked;
      if (input.checked && setting === 'swapKeyValue') {
        state.studySettings.randomizeKeyValue = false;
      } else if (input.checked && setting === 'randomizeKeyValue') {
        state.studySettings.swapKeyValue = false;
      }
      if (state.study && (setting === 'swapKeyValue' || setting === 'randomizeKeyValue')) {
        randomizeStudyPairDirections(state.study);
      }
      if (state.study && setting === 'showHintOnBack') {
        state.study.showBackHint = input.checked;
      }
      if (!(await saveStudySettings())) {
        render();
        return;
      }
      render();
    });
  });

  appEl.querySelectorAll('[data-study-font-size]').forEach((input) => {
    const update = () => {
      const side = input.dataset.studyFontSize;
      const size = Number(input.value);
      state.studySettings.fontSizes[side] = size;
      const output = appEl.querySelector(`[data-study-font-size-value="${side}"]`);
      if (output) output.textContent = `${size}px`;
      fitStudyCardText();
    };
    input.addEventListener('input', update);
    input.addEventListener('change', async () => {
      update();
      if (!(await saveStudySettings())) {
        render();
      }
    });
  });

  document.getElementById('homeBtn')?.addEventListener('click', () => finishStudy(true, 'workspace'));
}

document.addEventListener('change', async (event) => {
  const select = event.target.closest?.('[data-image-size]');
  if (!select || !IMAGE_SIZE_IDS.includes(select.value)) return;
  state.studySettings.imageSizes[select.dataset.imageSize] = select.value;
  await saveStudySettings();
  render();
});

document.addEventListener('keydown', handleStudyKeydown);

function attach() {
  attachDatePickers();
  const guide = document.querySelector('.start-guide');
  if (guide) {
    const number = guide.querySelector('#guideColumnNumber');
    const nameSeparator = guide.querySelector('#guideColumnNameSeparator');
    const name = guide.querySelector('#guideColumnName');
    const requiredSeparator = guide.querySelector('#guideColumnRequiredSeparator');
    const required = guide.querySelector('#guideColumnRequired');
    const description = guide.querySelector('#guideColumnDescription');
    const defaultMessage = guide.querySelector('.guide-default-message');
    const selectedDetail = guide.querySelector('.guide-column-selected');

    const updateGuideDetail = (column) => {
      defaultMessage.hidden = Boolean(column);
      selectedDetail.hidden = !column;
      const [columnNumber = '', columnName = ''] = (column?.dataset.guideTitle || '').split(' · ');
      number.textContent = columnNumber;
      name.textContent = columnName;
      nameSeparator.hidden = !column;
      requiredSeparator.hidden = !column;
      required.textContent = column?.dataset.guideRequired || '';
      required.className = `guide-column-required${['필수', 'Required'].includes(column?.dataset.guideRequired) ? ' is-required' : ''}`;
      description.textContent = column?.dataset.guideDescription || '';
    };

    const setHoveredColumn = (table, cell) => {
      const index = cell ? cell.dataset.col : null;
      table.querySelectorAll('[data-col]').forEach((item) => {
        item.classList.toggle('is-col-hover', index !== null && item.dataset.col === index);
      });
      updateGuideDetail(cell);
    };
    guide.querySelectorAll('.format-example-table').forEach((table) => {
      table.addEventListener('mouseover', (event) => {
        setHoveredColumn(table, event.target.closest('[data-col]'));
      });
      table.addEventListener('mouseleave', () => setHoveredColumn(table, null));
    });

    const exampleTabs = [...guide.querySelectorAll('[data-format-example]')];
    const examplePanels = [...guide.querySelectorAll('[data-format-example-panel]')];
    exampleTabs.forEach((tab) => {
      tab.addEventListener('click', () => {
        const selectedIndex = tab.dataset.formatExample;
        exampleTabs.forEach((item) => {
          const selected = item === tab;
          item.classList.toggle('is-active', selected);
          item.setAttribute('aria-selected', String(selected));
          item.tabIndex = selected ? 0 : -1;
        });
        examplePanels.forEach((panel) => {
          panel.hidden = panel.dataset.formatExamplePanel !== selectedIndex;
        });
        updateGuideDetail(null);
      });
      tab.addEventListener('keydown', (event) => {
        if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
        event.preventDefault();
        const currentIndex = exampleTabs.indexOf(tab);
        const nextIndex = event.key === 'Home'
          ? 0
          : event.key === 'End'
            ? exampleTabs.length - 1
            : (currentIndex + (event.key === 'ArrowRight' ? 1 : -1) + exampleTabs.length) % exampleTabs.length;
        exampleTabs[nextIndex].focus();
        exampleTabs[nextIndex].click();
      });
    });
  }

  if (state.page === 'study') {
    attachStudy();
    return;
  }

  if (state.page === 'study-records') {
    attachStudyRecords();
    return;
  }

  if (state.page === 'learning-status') {
    attachLearningStatusPage();
    return;
  }

  if (state.page === 'project-setup') {
    attachProjectSetupPage();
    return;
  }

  if (state.page === 'edit') {
    document.getElementById('homeBtn')?.addEventListener('click', leaveEditPage);
    appEl.querySelector('[data-edit-back]')?.addEventListener('click', leaveEditPage);
    const cells = [...appEl.querySelectorAll('[data-edit-row]')];
    const saveButton = appEl.querySelector('[data-edit-save]');
    cells.forEach(attachEditCell);
    const editSheet = appEl.querySelector('.edit-sheet');
    let columnWidths = [...DEFAULT_EDIT_COLUMN_WIDTHS];
    try {
      columnWidths = loadEditColumnWidths();
    } catch (error) {
      setEditStatus(error instanceof Error ? error.message : String(error), true);
    }
    if (editSheet) {
      applyEditColumnWidths(editSheet, columnWidths);
      const getColumnWidths = () => [...editSheet.querySelectorAll('[data-edit-column]')]
        .map((column) => column.getBoundingClientRect().width);
      const applyWidths = (leftIndex, leftWidth, rightWidth) => {
        const columns = [...editSheet.querySelectorAll('[data-edit-column]')];
        columns[leftIndex].style.width = `${leftWidth}px`;
        columns[leftIndex + 1].style.width = `${rightWidth}px`;
      };
      const saveWidths = () => {
        try {
          columnWidths = saveEditColumnWidths(getColumnWidths());
          setEditStatus(L('열 너비를 저장했어요.', 'Column widths saved.'));
        } catch (error) {
          setEditStatus(`${L('열 너비를 저장하지 못했어요.', 'Could not save column widths.')}${error instanceof Error ? ` ${error.message}` : ''}`, true);
        }
      };
      appEl.querySelectorAll('[data-resize-column]').forEach((handle) => {
        let drag = null;
        const leftIndex = Number(handle.dataset.resizeColumn);
        const columns = editSheet.querySelectorAll('[data-edit-column]');
        const updateHandleValue = () => {
          handle.setAttribute('aria-valuenow', String(Math.round(columns[leftIndex].getBoundingClientRect().width)));
          handle.setAttribute('aria-valuemin', String(MIN_EDIT_COLUMN_WIDTH));
        };
        updateHandleValue();
        handle.addEventListener('pointerdown', (event) => {
          event.preventDefault();
          drag = {
            pointerId: event.pointerId,
            startX: event.clientX,
          };
          handle.setPointerCapture(event.pointerId);
        });
        handle.addEventListener('pointermove', (event) => {
          if (!drag || drag.pointerId !== event.pointerId) return;
          const resized = resizeAdjacentColumns(
            getColumnWidths(),
            leftIndex,
            event.clientX - drag.startX,
            MIN_EDIT_COLUMN_WIDTH
          );
          applyWidths(leftIndex, resized[leftIndex], resized[leftIndex + 1]);
          updateHandleValue();
        });
        const finishResize = (event) => {
          if (!drag || drag.pointerId !== event.pointerId) return;
          drag = null;
          saveWidths();
        };
        handle.addEventListener('pointerup', finishResize);
        handle.addEventListener('pointercancel', finishResize);
        handle.addEventListener('keydown', (event) => {
          if (!['ArrowLeft', 'ArrowRight'].includes(event.key)) return;
          event.preventDefault();
          const widths = getColumnWidths();
          const step = event.shiftKey ? 24 : 8;
          const resized = resizeAdjacentColumns(
            widths,
            leftIndex,
            event.key === 'ArrowRight' ? step : -step,
            MIN_EDIT_COLUMN_WIDTH
          );
          applyWidths(leftIndex, resized[leftIndex], resized[leftIndex + 1]);
          updateHandleValue();
          saveWidths();
        });
      });
    }
    appEl.querySelector('[data-edit-add-image]')?.addEventListener('click', () => {
      const active = state.editActiveCell;
      const cell = active && appEl.querySelector(`[data-edit-row="${active.row}"][data-edit-col="${active.col}"]`);
      if (!cell) {
        setEditStatus(L('이미지를 넣을 셀을 먼저 선택해 주세요.', 'Select a cell to add the image to first.'), true);
        return;
      }
      pickImageFile().then((file) => file && addEditImages(cell, [file]));
    });
    saveButton?.addEventListener('click', saveEditedProject);
    appEl.querySelector('[data-edit-add-row]')?.addEventListener('click', () => {
      state.editRows.push(['', '', '']);
      state.editDirty = true;
      state.editFocus = { row: state.editRows.length - 1, col: 0 };
      render();
    });
    appEl.querySelectorAll('[data-edit-delete-row]').forEach((button) => {
      button.addEventListener('click', () => {
        state.editRows.splice(Number(button.dataset.editDeleteRow), 1);
        state.editDirty = true;
        render();
      });
    });
    if (state.editFocus) {
      const { row, col } = state.editFocus;
      state.editFocus = null;
      const target = appEl.querySelector(`[data-edit-row="${row}"][data-edit-col="${col}"]`);
      target?.focus();
      target?.scrollIntoView({ block: 'center' });
    }
  }

  const a4Layout = document.querySelector('.a4-view-screen');
  const setSettingsCollapsed = (collapsed) => {
    if (!a4Layout) return;
    settingsCollapsed = collapsed;
    a4Layout.classList.toggle('settings-collapsed', collapsed);
    const reopenButton = document.getElementById('reopenSettingsBtn');
    if (reopenButton) reopenButton.hidden = !collapsed;
    requestAnimationFrame(() => {
      renderA4View();
      document
        .getElementById(collapsed ? 'reopenSettingsBtn' : 'collapseSettingsBtn')
        ?.focus();
    });
  };
  const collapseSettingsButton = document.getElementById('collapseSettingsBtn');
  if (collapseSettingsButton) {
    collapseSettingsButton.addEventListener('click', () => setSettingsCollapsed(true));
  }
  const reopenSettingsButton = document.getElementById('reopenSettingsBtn');
  if (reopenSettingsButton) {
    reopenSettingsButton.addEventListener('click', () => setSettingsCollapsed(false));
  }

  const a4ResizeHandle = document.getElementById('a4PanelResize');
  if (a4Layout && a4ResizeHandle) {
    const maxPanelWidth = () => {
      const style = window.getComputedStyle(a4Layout);
      const horizontalPadding = parseFloat(style.paddingLeft) + parseFloat(style.paddingRight);
      return Math.max(240, a4Layout.clientWidth - horizontalPadding - 278);
    };
    const resizeTo = (clientX) => {
      const settingsPanel = a4Layout.querySelector('.a4-settings');
      if (!settingsPanel) return;
      studyPanelWidth = Math.round(
        Math.max(240, Math.min(maxPanelWidth(), clientX - settingsPanel.getBoundingClientRect().left))
      );
      a4Layout.style.setProperty('--study-panel-width', `${studyPanelWidth}px`);
      a4ResizeHandle.setAttribute('aria-valuenow', String(studyPanelWidth));
      a4ResizeHandle.setAttribute('aria-valuemax', String(maxPanelWidth()));
    };
    a4ResizeHandle.setAttribute('aria-valuemax', String(maxPanelWidth()));
    a4ResizeHandle.addEventListener('pointerdown', (event) => {
      a4ResizeHandle.setPointerCapture(event.pointerId);
      resizeTo(event.clientX);
    });
    a4ResizeHandle.addEventListener('pointermove', (event) => {
      if (a4ResizeHandle.hasPointerCapture(event.pointerId)) resizeTo(event.clientX);
    });
    a4ResizeHandle.addEventListener('keydown', (event) => {
      if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
      event.preventDefault();
      const settingsPanel = a4Layout.querySelector('.a4-settings');
      if (!settingsPanel) return;
      const direction = event.key === 'ArrowRight' ? 20 : -20;
      resizeTo(settingsPanel.getBoundingClientRect().left + studyPanelWidth + direction);
    });
  }

  const drop = document.getElementById('drop');
  if (drop) {
    const openFilePicker = () => {
      if (state.busy) return;
      if (!state.directoryConnected) {
        ensureWorkbookDirectoryAccess().then((allowed) => {
          if (!allowed) return;
          state.error = '';
          state.notice = L('저장 폴더 권한을 허용했어요. 파일 업로드를 다시 눌러 주세요.', 'Folder access granted. Click the file upload again.');
          render();
        });
        return;
      }
      fileInput.click();
    };
    drop.addEventListener('click', (event) => {
      if (event.target.closest('button')) return;
      openFilePicker();
    });
    drop.addEventListener('keydown', (event) => {
      if (event.target.closest('button')) return;
      if (event.key !== 'Enter' && event.key !== ' ') return;
      event.preventDefault();
      openFilePicker();
    });
    drop.addEventListener('dragover', (e) => {
      e.preventDefault();
      drop.classList.add('is-over');
    });
    drop.addEventListener('dragleave', () => drop.classList.remove('is-over'));
    drop.addEventListener('drop', (e) => {
      e.preventDefault();
      drop.classList.remove('is-over');
      const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (file) {
        ensureWorkbookDirectoryAccess().then((allowed) => {
          if (allowed) handleFile(file);
        });
      }
    });
  }

  const connectFolderBtn = document.getElementById('connectFolderBtn');
  if (connectFolderBtn) connectFolderBtn.addEventListener('click', connectWorkbookDirectory);
  const formatGuideModal = document.getElementById('formatGuideModal');
  if (formatGuideModal) {
    const openButton = appEl.querySelector('[data-open-format-guide]');
    const closeGuide = () => {
      formatGuideModal.hidden = true;
      openButton?.focus();
    };
    openButton?.addEventListener('click', () => {
      formatGuideModal.hidden = false;
      formatGuideModal.querySelector('.format-guide-close').focus();
    });
    formatGuideModal.querySelector('[data-create-blank-project]')?.addEventListener('click', () => {
      formatGuideModal.hidden = true;
      createBlankProject();
    });
    formatGuideModal.querySelector('[data-start-tutorial-deck]')?.addEventListener('click', () => {
      formatGuideModal.hidden = true;
      startTutorialDeck();
    });
    formatGuideModal.querySelectorAll('[data-close-format-guide]').forEach((el) => {
      el.addEventListener('click', closeGuide);
    });
    formatGuideModal.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') closeGuide();
    });
  }


  const startNewProject = () => {
    if (!state.directoryConnected) {
      ensureWorkbookDirectoryAccess().then((allowed) => {
        if (!allowed) return;
        state.error = '';
        state.notice = L('저장 폴더 권한을 허용했어요. 새 엑셀 카트리지 슬롯을 다시 눌러 주세요.', 'Folder access granted. Click the new cartridge slot again.');
        render();
      });
      return;
    }
    state.error = '';
    state.notice = '';
    fileInput.click();
  };

  appEl.querySelectorAll('[data-open-project]').forEach((button) => {
    button.addEventListener('click', async () => {
      await openProject(button.dataset.openProject);
    });
  });
  appEl.querySelectorAll('[data-project-menu-toggle]').forEach((toggle) => {
    toggle.addEventListener('click', (event) => {
      event.stopPropagation();
      const list = toggle.nextElementSibling;
      const open = list.hidden;
      closeProjectMenus();
      list.hidden = !open;
      toggle.setAttribute('aria-expanded', String(open));
    });
  });
  appEl.querySelectorAll('[data-edit-project]').forEach((button) => {
    button.addEventListener('click', () => {
      closeProjectMenus();
      editProject(button.dataset.editProject);
    });
  });
  appEl.querySelectorAll('[data-share-project]').forEach((button) => {
    button.addEventListener('click', () => {
      closeProjectMenus();
      shareProject(button.dataset.shareProject);
    });
  });
  appEl.querySelectorAll('[data-delete-project]').forEach((button) => {
    button.addEventListener('click', () => {
      closeProjectMenus();
      deleteProject(button.dataset.deleteProject);
    });
  });
  appEl.querySelectorAll('[data-add-project]').forEach((button) => {
    button.addEventListener('click', startNewProject);
    button.addEventListener('dragover', (event) => {
      event.preventDefault();
      button.classList.add('is-dragover');
    });
    button.addEventListener('dragleave', () => button.classList.remove('is-dragover'));
    button.addEventListener('drop', (event) => {
      event.preventDefault();
      button.classList.remove('is-dragover');
      const file = event.dataTransfer?.files?.[0];
      if (file) {
        ensureWorkbookDirectoryAccess().then((allowed) => {
          if (allowed) handleFile(file);
        });
      }
    });
  });

  const homeBtn = document.getElementById('homeBtn');
  if (homeBtn) {
    homeBtn.addEventListener('click', () => {
      state.page = 'workspace';
      state.error = '';
      state.notice = '';
      render();
    });
  }

  const a4PreviewBox = document.getElementById('a4PreviewBox');
  if (a4PreviewBox) {
    a4PreviewBox.addEventListener('click', (event) => {
      const control = event.target.closest('[data-study-action]');
      if (!control) return;
      const id = control.dataset.id;
      const studyAction = control.dataset.studyAction;
      if (studyAction === 'value') {
        if (isA4ValueVisible(id)) {
          if (state.showAllA4Values) state.hiddenA4Values.add(id);
          else state.revealedA4Values.delete(id);
        } else if (state.showAllA4Values) {
          state.hiddenA4Values.delete(id);
        } else {
          state.revealedA4Values.add(id);
        }
        syncStudyValueControls(a4PreviewBox);
      } else if (studyAction === 'hint') {
        const expanded = state.expandedA4Hints.has(id);
        if (expanded) state.expandedA4Hints.delete(id);
        else state.expandedA4Hints.add(id);
        renderA4View();
      }
    });
    a4PreviewBox.addEventListener('input', (event) => {
      const editor = event.target.closest('[data-study-edit]');
      if (!editor) return;
      const editType = editor.dataset.studyEdit;
      const value = editor.textContent || '';
      if (editType === 'title') state.a4Title = value;
      if (editType === 'info') state.a4Info = value;
      a4PreviewBox.querySelectorAll(`[data-study-edit="${editType}"]`).forEach((otherEditor) => {
        if (otherEditor !== editor) otherEditor.textContent = value;
      });
    });
    a4PreviewBox.addEventListener('focusout', (event) => {
      if (!event.target.closest('[data-study-edit]')) return;
      save();
    });
  }

  const a4SwapBtn = document.getElementById('a4SwapBtn');
  if (a4SwapBtn) {
    a4SwapBtn.addEventListener('click', async () => {
      state.studySettings.swapKeyValue = !state.studySettings.swapKeyValue;
      state.studySettings.randomizeKeyValue = false;
      await saveStudySettings();
      render();
    });
  }

  const toggleAllValuesBtn = document.getElementById('toggleAllValuesBtn');
  if (toggleAllValuesBtn) {
    toggleAllValuesBtn.addEventListener('click', () => {
      state.showAllA4Values = !state.showAllA4Values;
      state.revealedA4Values.clear();
      state.hiddenA4Values.clear();
      syncStudyValueControls(a4PreviewBox);
      toggleAllValuesBtn.textContent = state.showAllA4Values ? L('전체 value 가리기', 'Hide all values') : L('전체 value 보기', 'Show all values');
      toggleAllValuesBtn.setAttribute('aria-pressed', String(state.showAllA4Values));
      toggleAllValuesBtn.dataset.studyHint = state.showAllA4Values
        ? L('모든 value를 다시 가립니다.', 'Hide all values again.')
        : L('모든 value를 표시합니다.', 'Show all values.');
    });
  }

  const toggleAllHintsBtn = document.getElementById('toggleAllHintsBtn');
  if (toggleAllHintsBtn) {
    toggleAllHintsBtn.addEventListener('click', () => {
      const pairs = visibleA4HintPairs();
      const allExpanded = pairs.length > 0 && pairs.every((pair) => state.expandedA4Hints.has(pair.id));
      if (allExpanded) {
        pairs.forEach((pair) => state.expandedA4Hints.delete(pair.id));
      } else {
        pairs.forEach((pair) => state.expandedA4Hints.add(pair.id));
      }
      renderA4View();
    });
  }

  const studySearch = document.getElementById('studySearch');
  if (studySearch) {
    studySearch.addEventListener('input', () => {
      state.a4Query = studySearch.value;
      renderA4View();
    });
  }

  const a4SortSelect = document.getElementById('a4SortSelect');
  a4SortSelect?.addEventListener('change', () => {
    state.a4Sort = a4SortSelect.value === 'priority' ? 'priority' : 'order';
    const priorityStatus = document.querySelector('.a4-priority-status');
    if (priorityStatus) {
      priorityStatus.textContent = a4PriorityStatusText();
      priorityStatus.hidden = state.a4Sort !== 'priority' || !priorityStatus.textContent;
    }
    renderA4View();
  });

  const a4ReshuffleBtn = document.getElementById('a4ReshuffleBtn');
  if (a4ReshuffleBtn) {
    a4ReshuffleBtn.addEventListener('click', () => {
      state.a4Order = shuffleA4Order(state.pairs);
      playShuffleEffect(a4ReshuffleBtn);
      renderA4View();
    });
  }

  const a4PrintBtn = document.getElementById('a4PrintBtn');
  if (a4PrintBtn) {
    a4PrintBtn.disabled = !a4Pages().length;
    a4PrintBtn.addEventListener('click', handleA4Print);
  }

}

/* ---------------- 시작 ---------------- */

const resizeObserver = new ResizeObserver(() => {
  const box = document.getElementById('a4PreviewBox');
  if (!box) return;
  const next = Math.min(1, Math.max(0.2, (box.clientWidth || SHEET_W_PX) / SHEET_W_PX));
  if (Math.abs(next - currentScale) > 0.002) renderA4View();
});

async function init() {
  await load();
  if (state.directoryConnected && state.directoryHandle) {
    try {
      const result = await prepareWorkbookDirectory(state.directoryHandle);
      if (result.recoveredCount) {
        state.notice = L(
          `저장 위치에서 기존 프로젝트 ${result.recoveredCount}개를 불러왔어요.`,
          `Loaded ${result.recoveredCount} existing project(s) from the save location.`
        );
      }
    } catch (error) {
      state.error = L(
        `저장 폴더 자료를 불러오지 못했어요: ${error instanceof Error ? error.message : String(error)}`,
        `Could not load projects from the save folder: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }
  mountSupportWidget();
  render();
}

init();
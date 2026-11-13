import { appendStudySession, emptyStudyRecords, removeStudySession } from './lib/studyRecords.js?v=3';
import {
  parseStudyFile,
  normalizeStudyWorkbook,
  readSheetGrid,
  writeSheetGrid,
  createStudyWorkbook,
} from './lib/parseSheet.js';
import { normalizeColumnWidths, resizeAdjacentColumns } from './lib/columnSizing.js';
import { indexStudyPairs } from './lib/studyPairs.js';
import { AnkiImportError, buildAnkiRows, audioOnlyFields, countFieldImages, loadAnkiDeck, previewAnkiCards } from './lib/ankiImport.js?v=3';
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
} from './lib/fsrs.js?v=8';
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
  saveProjectOrder as saveProjectOrderRecord,
  readLibraryFile,
  writeLibraryFile,
  saveStudySettings as saveStudySettingsToStore,
  readStudyRecordFile,
  writeStudyRecordFile,
} from './lib/projectStore.js?v=6';
import { recoverWorkbookProjects } from './lib/projectRecovery.js?v=7';
import {
  DEFAULT_DAILY_NEW_CARDS,
  dailyNewCardPlan,
  firstPassDeadline,
  localStudyDate,
  manualDailyNewCardPlan,
} from './lib/studyPace.js?v=2';

const EDIT_COLUMN_WIDTHS_KEY = 'esper.editColumnWidths.v1';
const STUDY_TIME_LIMIT_RANGE = { min: 1, max: 1440 };
const DEFAULT_STUDY_SETTINGS = {
  roundTimeLimitMinutes: 30,
  showHintOnBack: true,
  swapKeyValue: false,
  randomizeKeyValue: false,
  fontSizes: { front: 72, back: 43, hintFront: 26, hintBack: 26 },
  examDate: '',
  desiredRetention: DEFAULT_DESIRED_RETENTION,
  imageSizes: { study: 55, a4: 'medium', status: 'medium' },
  imageHeights: { front: 55, back: 55, hintFront: 55, hintBack: 55 },
  a4ImageHeight: 14,
  retentionStyle: { color: '#8f3f26', strength: 40, half: 70 },
  shortcuts: {
    flip: 'Space',
    hint: 'ArrowDown',
    again: '1',
    hard: '2',
    good: '3',
    easy: '4',
    undo: 'ArrowLeft',
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
  ['undo', '선택 되돌리기'],
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
  undo: 'Undo rating',
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
  const insertAfter = deltaX > 0;
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
      else state.error = message;
      render();
    });
  }, 0);
}, true);
appEl.addEventListener('click', (event) => {
  if (event.target.closest?.('#connectFolderBtn')) connectWorkbookDirectory();
});
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

// Browser Back is disabled; pages are left only through in-page controls.
history.pushState({ esperGuard: true }, '');
window.addEventListener('popstate', () => history.pushState({ esperGuard: true }, ''));

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
  folderPromptDismissed: false,
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
  editHighlightRow: null,
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
  logPreviewLoading: false,
  logPreviewError: '',
  logPreviewView: 'monthly',
  logSelectedDate: null,
  logCalMonth: localDateForExport().slice(0, 7),
  logYear: Number(localDateForExport().slice(0, 4)),
  logSessionDeleting: false,
  logDeleteError: '',
  editColumnWidths: null,
  libraryNeedsMigration: false,
  learningStatusFilters: ['due', 'scheduled', 'new'],
  learningStatusPlanMessage: '',
  learningStatusQuery: '',
  learningStatusView: 'simple',
  collapsedSections: {},
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
  const savedTimeLimit = settings?.roundTimeLimitMinutes;
  const roundTimeLimitMinutes = savedTimeLimit === '' || savedTimeLimit === null || savedTimeLimit === undefined
    ? Number.NaN
    : Number(savedTimeLimit);
  const shortcuts = { ...DEFAULT_STUDY_SETTINGS.shortcuts };
  for (const [action] of STUDY_ACTIONS) {
    const key = savedShortcuts[action] ?? (action === 'undo' ? savedShortcuts.previous : undefined);
    if (typeof key === 'string' && key.length > 0) shortcuts[action] = key;
  }
  const uniqueKeys = new Set(Object.values(shortcuts).map(shortcutKeyId));
  if (uniqueKeys.size !== STUDY_ACTIONS.length) {
    Object.assign(shortcuts, DEFAULT_STUDY_SETTINGS.shortcuts);
  }
  return {
    roundTimeLimitMinutes: Number.isFinite(roundTimeLimitMinutes)
      ? Math.min(STUDY_TIME_LIMIT_RANGE.max, Math.max(STUDY_TIME_LIMIT_RANGE.min, Math.round(roundTimeLimitMinutes)))
      : DEFAULT_STUDY_SETTINGS.roundTimeLimitMinutes,
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
      hintFront: Number.isFinite(Number(settings?.fontSizes?.hintFront))
        ? Math.min(72, Math.max(12, Number(settings.fontSizes.hintFront)))
        : DEFAULT_STUDY_SETTINGS.fontSizes.hintFront,
      hintBack: Number.isFinite(Number(settings?.fontSizes?.hintBack))
        ? Math.min(72, Math.max(12, Number(settings.fontSizes.hintBack)))
        : DEFAULT_STUDY_SETTINGS.fontSizes.hintBack,
    },
    examDate: /^\d{4}-\d{2}-\d{2}$/.test(settings?.examDate || '') ? settings.examDate : '',
    desiredRetention: normalizeDesiredRetention(settings?.desiredRetention ?? DEFAULT_DESIRED_RETENTION),
    imageSizes: Object.fromEntries(IMAGE_SIZE_VIEWS.map((view) => {
      const saved = settings?.imageSizes?.[view];
      if (view === 'study') return [view, normalizeStudyImageHeight(saved)];
      return [view, IMAGE_SIZE_IDS.includes(saved) ? saved : DEFAULT_STUDY_SETTINGS.imageSizes[view]];
    })),
    imageHeights: Object.fromEntries(['front', 'back', 'hintFront', 'hintBack'].map((field) => {
      const legacyField = field.startsWith('hint') ? field.slice(4).toLowerCase() : field;
      return [
        field,
        normalizeStudyImageHeight(settings?.imageHeights?.[field] ?? settings?.imageHeights?.[legacyField] ?? settings?.imageSizes?.study),
      ];
    })),
    a4ImageHeight: normalizeA4ImageHeight(settings?.a4ImageHeight ?? settings?.imageSizes?.a4),
    retentionStyle: normalizeRetentionStyle(settings?.retentionStyle),
    shortcuts,
  };
}

function normalizeRetentionStyle(style) {
  const defaults = DEFAULT_STUDY_SETTINGS.retentionStyle;
  const clamp = (value, min, max, fallback) => (Number.isFinite(Number(value)) ? Math.round(Math.min(max, Math.max(min, Number(value)))) : fallback);
  return {
    color: /^#[0-9a-f]{6}$/i.test(style?.color || '') ? style.color.toLowerCase() : defaults.color,
    strength: clamp(style?.strength, 5, 100, defaults.strength),
    half: clamp(style?.half, 5, 95, defaults.half),
  };
}

const A4_IMAGE_HEIGHT_RANGE = { min: 6, max: 50 };
const LEGACY_A4_IMAGE_HEIGHTS = { small: 8, medium: 14, large: 24, fill: 45 };

function normalizeA4ImageHeight(value) {
  const number = typeof value === 'string' && value in LEGACY_A4_IMAGE_HEIGHTS
    ? LEGACY_A4_IMAGE_HEIGHTS[value]
    : Number(value);
  if (!Number.isFinite(number)) return DEFAULT_STUDY_SETTINGS.a4ImageHeight;
  return Math.min(A4_IMAGE_HEIGHT_RANGE.max, Math.max(A4_IMAGE_HEIGHT_RANGE.min, Math.round(number)));
}

const IMAGE_SIZE_VIEWS = ['study', 'a4', 'status'];
const STUDY_IMAGE_HEIGHT_RANGE = { min: 8, max: 60 };
const LEGACY_STUDY_IMAGE_HEIGHTS = { small: 12, medium: 24, large: 40, fill: 55 };

function normalizeStudyImageHeight(value) {
  const number = typeof value === 'string' && value in LEGACY_STUDY_IMAGE_HEIGHTS
    ? LEGACY_STUDY_IMAGE_HEIGHTS[value]
    : Number(value);
  if (!Number.isFinite(number)) return DEFAULT_STUDY_SETTINGS.imageSizes.study;
  return Math.min(STUDY_IMAGE_HEIGHT_RANGE.max, Math.max(STUDY_IMAGE_HEIGHT_RANGE.min, Math.round(number)));
}
const IMAGE_SIZE_IDS = ['small', 'medium', 'large', 'fill'];

function fitStudyCardText() {
  const face = document.querySelector('.study-card-face');
  const primaryText = face?.querySelector('.study-key, .study-value');
  if (!face || !primaryText) return;

  const side = primaryText.classList.contains('study-key') ? 'front' : 'back';
  const fields = [
    { element: primaryText, field: side },
    ...[...face.querySelectorAll('.study-hint')].map((element) => ({ element, field: `hint${side === 'front' ? 'Front' : 'Back'}` })),
  ];
  fields.forEach(({ element }) => {
    element.style.maxHeight = 'none';
    element.style.overflowY = 'visible';
  });
  const minimumScale = Math.max(...fields.map(({ field }) => 12 / state.studySettings.fontSizes[field]));
  let low = minimumScale;
  let high = 1;
  for (let attempt = 0; attempt < 12; attempt += 1) {
    const scale = (low + high) / 2;
    fields.forEach(({ element, field }) => {
      element.style.fontSize = `${state.studySettings.fontSizes[field] * scale}px`;
    });
    if (face.scrollHeight <= face.clientHeight + 1) low = scale;
    else high = scale;
  }
  fields.forEach(({ element, field }) => {
    element.style.fontSize = `${Math.max(12, state.studySettings.fontSizes[field] * low)}px`;
  });
}

window.addEventListener('resize', fitStudyCardText);
document.addEventListener('load', (event) => {
  if (event.target instanceof HTMLImageElement && event.target.closest('.study-card')) {
    fitStudyCardText();
  }
}, true);

function applyImageSizes() {
  const sizes = state.studySettings.imageSizes;
  const heights = state.studySettings.imageHeights;
  document.documentElement.style.setProperty('--study-img-h-front', String(heights.front));
  document.documentElement.style.setProperty('--study-img-h-back', String(heights.back));
  document.documentElement.style.setProperty('--study-img-h-hint-front', String(heights.hintFront));
  document.documentElement.style.setProperty('--study-img-h-hint-back', String(heights.hintBack));
  document.documentElement.style.setProperty('--a4-img-h', `${state.studySettings.a4ImageHeight}mm`);
  const retentionColor = state.studySettings.retentionStyle.color;
  document.documentElement.style.setProperty('--retention-rgb', [1, 3, 5].map((i) => parseInt(retentionColor.slice(i, i + 2), 16)).join(', '));
  IMAGE_SIZE_VIEWS.filter((view) => view !== 'study' && view !== 'a4').forEach((view) => {
    document.documentElement.dataset[`img${view[0].toUpperCase()}${view.slice(1)}`] = sizes[view];
  });
}

// 기억률이 낮을수록 카드가 눈에 띄게 흐려지도록 곡선을 가파르게 잡습니다.
// 기억 확률 100%에서 절대 투명도(strength)가 되고, 'half' 지점에서 절반이 되며,
// 그 간격만큼 내려갈 때마다 다시 절반이 되는 반감기 곡선입니다.
function retentionRelative(probability, half) {
  const p = Math.max(0, Math.min(1, probability));
  return 2 ** (-(1 - p) / (1 - half / 100));
}

function applyRetentionOpacity(card, probability) {
  const { strength, half } = state.studySettings.retentionStyle;
  const relative = retentionRelative(probability, half);
  card.style.setProperty('--retention-opacity', Math.max(0.02, (strength / 100) * relative).toFixed(3));
  card.style.setProperty('--retention-content-opacity', (0.15 + 0.85 * relative).toFixed(2));
}

function currentStudySide() {
  return state.study?.flipped ? 'back' : 'front';
}

const STUDY_FONT_RANGE = { front: { min: 18, max: 120 }, back: { min: 16, max: 72 } };

function cardSizeControlsHTML(side, field, fieldLabel, hasImage = true) {
  const settings = state.studySettings;
  const font = field.startsWith('hint') ? { min: 12, max: 72 } : STUDY_FONT_RANGE[side];
  const sideLabel = side === 'front' ? L('앞면', 'Front') : L('뒷면', 'Back');
  return `
    <div class="study-size-group">
      <h3>${fieldLabel}</h3>
      <label class="study-font-setting">
        <span>${L('글자', 'Text')}</span>
        <input type="range" min="${font.min}" max="${font.max}" step="1" data-study-font="${field}" value="${settings.fontSizes[field]}" aria-label="${sideLabel} ${fieldLabel} ${L('글자 크기', 'text size')}">
        <output>${settings.fontSizes[field]}px</output>
      </label>
      <label class="study-font-setting${hasImage ? '' : ' is-disabled'}"${hasImage ? '' : ` title="${L('이 항목에는 이미지가 없어요', 'No image in this field')}"`}>
        <span>${L('이미지', 'Image')}</span>
        <input type="range" min="${STUDY_IMAGE_HEIGHT_RANGE.min}" max="${STUDY_IMAGE_HEIGHT_RANGE.max}" step="1"${hasImage ? '' : ' disabled'} data-study-image="${field}" value="${settings.imageHeights[field]}" aria-label="${sideLabel} ${fieldLabel} ${L('이미지 크기', 'image size')}">
        <output>${settings.imageHeights[field]}</output>
      </label>
    </div>`;
}

function setCardSize(field, kind, rawValue) {
  const settings = state.studySettings;
  if (kind === 'font') {
    const { min, max } = field.startsWith('hint') ? { min: 12, max: 72 } : STUDY_FONT_RANGE[field];
    settings.fontSizes[field] = Math.min(max, Math.max(min, Number(rawValue)));
    return settings.fontSizes[field];
  }
  settings.imageHeights[field] = normalizeStudyImageHeight(rawValue);
  applyImageSizes();
  return settings.imageHeights[field];
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

// The save folder keeps a copy of the project order and colors so they survive changing folders.
async function persistLibrary(ids = state.projects.map((project) => project.id)) {
  const byId = new Map(state.projects.map((project) => [project.id, project]));
  const ordered = ids.map((id) => byId.get(id)).filter(Boolean);
  const previous = await readLibraryFile();
  const savedModes = previous?.dailyPaceModes;
  const dailyPaceModes = savedModes && typeof savedModes === 'object' && !Array.isArray(savedModes)
    ? { ...savedModes }
    : {};
  for (const project of ordered) {
    if (project.plan && typeof project.plan.useRecommendedDailyPace === 'boolean') {
      dailyPaceModes[project.name] = project.plan.useRecommendedDailyPace;
    }
  }
  const savedPreferences = previous?.preferences && typeof previous.preferences === 'object' && !Array.isArray(previous.preferences)
    ? previous.preferences
    : {};
  const savedLearningStatus = savedPreferences.learningStatus && typeof savedPreferences.learningStatus === 'object'
    ? savedPreferences.learningStatus
    : {};
  const savedCollapsedSections = savedLearningStatus.collapsedSections && typeof savedLearningStatus.collapsedSections === 'object'
    ? savedLearningStatus.collapsedSections
    : {};
  const savedMyLog = savedPreferences.myLog && typeof savedPreferences.myLog === 'object'
    ? savedPreferences.myLog
    : {};
  const preferences = {
    ...savedPreferences,
    version: 1,
    language,
    studySettings: {
      ...(savedPreferences.studySettings && typeof savedPreferences.studySettings === 'object'
        ? savedPreferences.studySettings
        : {}),
      ...state.studySettings,
      fontSizes: {
        ...savedPreferences.studySettings?.fontSizes,
        ...state.studySettings.fontSizes,
      },
      imageSizes: {
        ...savedPreferences.studySettings?.imageSizes,
        ...state.studySettings.imageSizes,
      },
      imageHeights: {
        ...savedPreferences.studySettings?.imageHeights,
        ...state.studySettings.imageHeights,
      },
      shortcuts: {
        ...savedPreferences.studySettings?.shortcuts,
        ...state.studySettings.shortcuts,
      },
      retentionStyle: {
        ...savedPreferences.studySettings?.retentionStyle,
        ...state.studySettings.retentionStyle,
      },
    },
    editColumnWidths: state.editColumnWidths || [...DEFAULT_EDIT_COLUMN_WIDTHS],
    learningStatus: {
      ...savedLearningStatus,
      view: state.learningStatusView,
      collapsedSections: {
        ...savedCollapsedSections,
        ...state.collapsedSections,
      },
    },
    myLog: {
      ...savedMyLog,
      view: state.logPreviewView,
    },
  };
  const saved = await writeLibraryFile({
    ...(previous && typeof previous === 'object' && !Array.isArray(previous) ? previous : {}),
    version: 1,
    order: ordered.map((project) => project.name),
    colors: Object.fromEntries(ordered.map((project) => [project.name, project.cartridgeColorIndex ?? 0])),
    retentionStyle: state.studySettings.retentionStyle,
    dailyPaceModes,
    preferences,
  });
  if (!saved && state.directoryConnected) {
    throw new Error(L('저장 폴더 권한을 확인해 주세요.', 'Please check access to the save folder.'));
  }
  return saved;
}

async function persistUserPreferences() {
  try {
    const saved = await persistLibrary();
    if (saved) state.libraryNeedsMigration = false;
    return true;
  } catch (error) {
    state.error = L(
      `설정을 esper-library.json에 저장하지 못했어요: ${error instanceof Error ? error.message : String(error)}`,
      `Could not save preferences to esper-library.json: ${error instanceof Error ? error.message : String(error)}`
    );
    render();
    return false;
  }
}

async function saveProjectOrder(ids) {
  await saveProjectOrderRecord(ids);
  await persistLibrary(ids);
}

function applyLibraryPreferences(library, fallbackLanguage, fallbackStudySettings) {
  const preferences = library?.preferences && typeof library.preferences === 'object' && !Array.isArray(library.preferences)
    ? library.preferences
    : {};
  const libraryStudySettings = preferences.studySettings && typeof preferences.studySettings === 'object'
    ? preferences.studySettings
    : {};
  state.studySettings = normalizeStudySettings({
    ...fallbackStudySettings,
    ...libraryStudySettings,
    fontSizes: { ...fallbackStudySettings.fontSizes, ...libraryStudySettings.fontSizes },
    imageSizes: { ...fallbackStudySettings.imageSizes, ...libraryStudySettings.imageSizes },
    imageHeights: { ...fallbackStudySettings.imageHeights, ...libraryStudySettings.imageHeights },
    shortcuts: { ...fallbackStudySettings.shortcuts, ...libraryStudySettings.shortcuts },
    retentionStyle: libraryStudySettings.retentionStyle ?? library?.retentionStyle ?? fallbackStudySettings.retentionStyle,
  });
  const preferredLanguage = preferences.language === 'ko' || preferences.language === 'en'
    ? preferences.language
    : fallbackLanguage;
  setLanguageValue(preferredLanguage || (Intl.DateTimeFormat().resolvedOptions().timeZone === 'Asia/Seoul' ? 'ko' : 'en'));
  if (preferences.learningStatus?.view === 'simple' || preferences.learningStatus?.view === 'detailed') {
    state.learningStatusView = preferences.learningStatus.view;
  }
  if (preferences.learningStatus?.collapsedSections && typeof preferences.learningStatus.collapsedSections === 'object') {
    state.collapsedSections = Object.fromEntries(
      Object.entries(preferences.learningStatus.collapsedSections)
        .filter(([, collapsed]) => typeof collapsed === 'boolean')
    );
  }
  const myLogView = preferences.myLog?.view;
  if (myLogView === 'monthly' || myLogView === 'yearly') state.logPreviewView = myLogView;
  else if (myLogView === 'simple') state.logPreviewView = 'monthly';
  else if (myLogView === 'detailed') state.logPreviewView = 'yearly';
  const preferredWidths = preferences.editColumnWidths;
  if (validEditColumnWidths(preferredWidths)) {
    state.editColumnWidths = [...preferredWidths];
  } else {
    try {
      const legacyWidths = localStorage.getItem(EDIT_COLUMN_WIDTHS_KEY);
      const parsedWidths = legacyWidths === null ? null : JSON.parse(legacyWidths);
      state.editColumnWidths = validEditColumnWidths(parsedWidths)
        ? parsedWidths
        : [...DEFAULT_EDIT_COLUMN_WIDTHS];
    } catch {
      state.editColumnWidths = [...DEFAULT_EDIT_COLUMN_WIDTHS];
    }
  }
  state.libraryNeedsMigration = !preferences.version ||
    !preferences.studySettings ||
    !preferences.language ||
    !preferences.editColumnWidths ||
    !preferences.learningStatus ||
    !preferences.myLog;
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
    const savedStudySettings = await getStudySettings() || {};
    const library = await readLibraryFile();
    applyLibraryPreferences(library, savedLanguage, savedStudySettings);
    for (const project of state.projects) {
      const savedMode = library?.dailyPaceModes?.[project.name];
      if (typeof savedMode === 'boolean') {
        project.plan = { ...(project.plan || {}), useRecommendedDailyPace: savedMode };
      }
    }
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
  // 튜토리얼 페이지가 같은 언어로 영상 자막을 설정할 수 있도록 공유합니다.
  try { localStorage.setItem('esper-landing-lang', language); } catch { /* 저장소 접근 불가 시 무시 */ }
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
  const preferencesSaved = await persistUserPreferences();
  if (preferencesSaved) state.error = '';
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
    state.studyRecords = null;
    state.logDeleteError = '';
    state.logPreviewError = '';
    state.fsrsData = fsrsData;
    state.pairs = project.pairs;
    state.sourceFile = file;
    state.revealedA4Values = new Set();
    state.hiddenA4Values = new Set();
    state.expandedA4Hints = new Set();
    state.showAllA4Values = true;
    state.a4Title = null;
    state.a4Info = null;
    state.a4Order = items.map((item) => item.id);
    state.a4Query = '';
    state.sourceName = file.name;
    state.title = project.title;
    await startStudy();
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
  const audio = audioOnlyFields(deck, noteTypeId);
  const usable = Array.from({ length: fieldCount }, (_, index) => index).filter((index) => !audio[index]);
  const imageFields = usable.slice(2).filter((index) => imageStats[index].found);
  return {
    key: usable.length > 0 ? [usable[0]] : [],
    value: usable.length > 1 ? [usable[1]] : [],
    hint: imageFields.length ? imageFields : usable.length > 2 ? [usable[2]] : [],
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
      previewIndex: 0,
      previewFlipped: false,
      previewHintOpen: false,
      previewShowHint: false,
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

const ankiFontSize = (size) => `${(size / 7.6).toFixed(3)}cqw`;

function ankiImportHTML() {
  const anki = state.anki;
  const noteType = anki.deck.noteTypes.find((type) => type.id === anki.noteTypeId);
  const imageStats = countFieldImages(anki.deck, anki.noteTypeId);
  const audioFields = audioOnlyFields(anki.deck, anki.noteTypeId);
  const preview = previewAnkiCards(anki.deck, anki.noteTypeId, anki.mapping, 'newline', anki.previewIndex);
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
  const settings = state.studySettings;
  const swapped = settings.swapKeyValue;
  const frontParts = preview.card ? (swapped ? preview.card.value : preview.card.key) : [];
  const backParts = preview.card ? (swapped ? preview.card.key : preview.card.value) : [];
  const hintContent = preview.card?.hint.length
    ? renderParts(preview.card.hint)
    : `<span class="anki-preview-nohint">${L('등록된 hint가 없어요.', 'No hint available.')}</span>`;
  const hintButton = (open) => `<button type="button" class="anki-hint-toggle" data-anki-hint-toggle>${open ? L('hint 닫기', 'Hide hint') : L('hint 보기', 'Show hint')}</button>`;
  const fontStyle = (side) => `style="font-size:${ankiFontSize(settings.fontSizes[side])}"`;
  const previewFace = preview.card
    ? anki.previewFlipped
      ? `<div class="anki-preview-card is-back" data-anki-flip tabindex="0" role="button" aria-label="${L('카드 뒤집기', 'Flip card')}">
          <div class="anki-preview-content" ${fontStyle('back')}>${renderParts(backParts)}</div>
          ${settings.showHintOnBack || anki.previewHintOpen
            ? `<div class="anki-preview-hint" ${fontStyle('back')}>${hintContent}</div>`
            : hintButton(false)}
        </div>`
      : `<div class="anki-preview-card" data-anki-flip tabindex="0" role="button" aria-label="${L('카드 뒤집기', 'Flip card')}">
          <div class="anki-preview-content" ${fontStyle('front')}>${renderParts(frontParts)}</div>
          ${anki.previewHintOpen ? `<div class="anki-preview-hint">${hintContent}</div>` : ''}
          ${hintButton(anki.previewHintOpen)}
        </div>`
    : `<p class="anki-preview-empty">${L('미리 볼 수 있는 카드가 없어요. Key와 Value 필드 배치를 확인해 주세요.', 'No cards to preview. Check the Key and Value field mapping.')}</p>`;
  const ankiFaceHasImage = () => {
    const hasImage = (parts) => (parts || []).some((part) => part.type === 'image');
    const hint = preview.card?.hint;
    return anki.previewFlipped
      ? hasImage(backParts) || ((settings.showHintOnBack || anki.previewHintOpen) && hasImage(hint))
      : hasImage(frontParts) || (anki.previewHintOpen && hasImage(hint));
  };
  const previewSettings = `
    <aside class="anki-preview-settings" aria-label="${L('학습 카드 설정', 'Card settings')}">
      <h3>${anki.previewFlipped ? L('카드 크기 · 뒷면', 'Card size · Back') : L('카드 크기 · 앞면', 'Card size · Front')}</h3>
      ${cardSizeControlsHTML(anki.previewFlipped ? 'back' : 'front', 'anki', ankiFaceHasImage())}
      <label class="anki-mini-check"><input type="checkbox" data-anki-setting="showHintOnBack"${settings.showHintOnBack ? ' checked' : ''}> ${L('뒷면에 hint 기본 표시', 'Show hint on back')}</label>
      <label class="anki-mini-check"><input type="checkbox" data-anki-setting="swapKeyValue"${settings.swapKeyValue ? ' checked' : ''}> ${L('Key·Value 반전', 'Swap key and value')}</label>
      <p class="anki-mini-note">${L('지금 보고 있는 면의 모든 카드에 적용되고, 학습 화면에도 그대로 반영돼요.', 'Applies to this side of every card and carries over to study mode.')}</p>
    </aside>`;
  const badge = (index) => {
    const { found, missing } = imageStats[index];
    return `${found ? ` <span class="anki-badge">🖼 ${found}</span>` : ''}${missing ? ` <span class="anki-badge is-warn">⚠ ${missing}</span>` : ''}`;
  };
  return `
    <div class="anki-overlay">
      <div class="anki-dialog" role="dialog" aria-modal="true" aria-labelledby="ankiTitle">
        <h2 id="ankiTitle">${L('Anki 덱 가져오기', 'Import Anki deck')}</h2>
        <p class="anki-sub">${esc(anki.fileName)} · ${anki.deck.hasAudio ? `<span class="anki-audio-note">${L('이 덱에는 오디오가 포함되어 있어요. 아직 오디오 파일은 지원하지 않아 가져오지 않아요. 대신 학습 중 브라우저 내장 음성(TTS)이 key 값을 읽어 줘요.', 'This deck contains audio. Audio files are not supported yet, so they will not be imported. Instead, the browser’s built-in text-to-speech (TTS) reads the key aloud while you study.')}</span>` : L('오디오는 가져오지 않아요.', 'Audio is not imported.')}</p>
        <div class="anki-row">
          <label>${L('노트 타입', 'Note type')}
            <select data-anki-type${disabled}>
              ${anki.deck.noteTypes.map((type) => `<option value="${esc(type.id)}"${type.id === anki.noteTypeId ? ' selected' : ''}>${esc(type.name)} (${type.noteCount})</option>`).join('')}
            </select>
          </label>
        </div>
        <div class="anki-areas">
          ${areas.map(([area, title, description]) => `
            <fieldset class="anki-area"${disabled}>
              <legend>${title}</legend>
              <small>${description}</small>
              ${[
                ...anki.mapping[area],
                ...noteType.fields.map((_, index) => index).filter((index) => !audioFields[index] && !anki.mapping[area].includes(index)),
              ].map((index) => {
                const field = noteType.fields[index];
                const position = anki.mapping[area].indexOf(index);
                const checked = position >= 0;
                const move = (direction, symbol, label, edge) => `<button type="button" class="anki-move" data-anki-move="${area}:${index}:${direction}" aria-label="${label}"${disabled}${edge ? ' disabled' : ''}>${symbol}</button>`;
                return `
                <div class="anki-field-row">
                  <label><input type="checkbox" data-anki-area="${area}" value="${index}"${checked ? ' checked' : ''}> <span>${esc(field)}${badge(index)}</span></label>
                  ${checked && anki.mapping[area].length > 1 ? `<span class="anki-move-group">${move(-1, '↑', L('위로 이동', 'Move up'), position === 0)}${move(1, '↓', L('아래로 이동', 'Move down'), position === anki.mapping[area].length - 1)}</span>` : ''}
                </div>`;
              }).join('')}
            </fieldset>`).join('')}
        </div>
        <p class="anki-count">${L(`가져올 항목 ${preview.total}개`, `${preview.total} items to import`)}</p>
        <div class="anki-divider" role="separator"><span>${L('미리보기', 'Preview')}</span></div>
        <section class="anki-card-preview" aria-label="${L('학습 카드 미리보기', 'Flashcard preview')}">
          ${previewSettings}
          <div class="anki-preview-stage">
            ${previewFace}
            <div class="anki-preview-navigation">
              <button type="button" class="btn" data-anki-nav="-1"${disabled}${anki.previewIndex <= 0 ? ' disabled' : ''} aria-label="${L('이전 카드', 'Previous card')}">←</button>
              <span aria-live="polite" aria-atomic="true">${preview.total ? `${anki.previewIndex + 1} / ${preview.total}` : '0 / 0'}</span>
              <button type="button" class="btn" data-anki-nav="1"${disabled}${anki.previewIndex >= preview.total - 1 ? ' disabled' : ''} aria-label="${L('다음 카드', 'Next card')}">→</button>
            </div>
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
  const dialog = overlay.querySelector('.anki-dialog');
  if (anki.scrollTop) {
    dialog.scrollTop = anki.scrollTop;
    anki.scrollTop = 0;
  }
  const refresh = (focus) => {
    anki.focus = focus;
    anki.error = '';
    anki.scrollTop = dialog.scrollTop;
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
  overlay.querySelectorAll('[data-anki-area]').forEach((box) => {
    box.addEventListener('change', () => {
      const area = box.dataset.ankiArea;
      const index = Number(box.value);
      const next = anki.mapping[area].filter((item) => item !== index);
      if (box.checked) next.push(index);
      anki.mapping[area] = next;
      releaseAnkiPreviewUrls(anki);
      refresh(`[data-anki-area="${area}"][value="${index}"]`);
    });
  });
  overlay.querySelectorAll('[data-anki-move]').forEach((button) => {
    button.addEventListener('click', () => {
      const [area, indexText, directionText] = button.dataset.ankiMove.split(':');
      const index = Number(indexText);
      const list = [...anki.mapping[area]];
      const from = list.indexOf(index);
      const to = from + Number(directionText);
      if (from < 0 || to < 0 || to >= list.length) return;
      [list[from], list[to]] = [list[to], list[from]];
      anki.mapping[area] = list;
      releaseAnkiPreviewUrls(anki);
      refresh(`[data-anki-move="${button.dataset.ankiMove}"]`);
    });
  });
  const flipCard = () => {
    anki.previewFlipped = !anki.previewFlipped;
    refresh('[data-anki-flip]');
  };
  const flipEl = overlay.querySelector('[data-anki-flip]');
  flipEl?.addEventListener('click', flipCard);
  flipEl?.addEventListener('keydown', (event) => {
    if (event.target === flipEl && (event.key === ' ' || event.key === 'Enter')) {
      event.preventDefault();
      flipCard();
    }
  });
  overlay.querySelector('[data-anki-hint-toggle]')?.addEventListener('click', (event) => {
    event.stopPropagation();
    anki.previewHintOpen = !anki.previewHintOpen;
    refresh('[data-anki-hint-toggle]');
  });
  [['font', 'data-anki-font'], ['image', 'data-anki-image']].forEach(([kind, attr]) => {
    overlay.querySelectorAll(`[${attr}]`).forEach((input) => {
      const side = anki.previewFlipped ? 'back' : 'front';
      input.addEventListener('input', () => {
        const value = setCardSize(side, kind, input.value);
        input.nextElementSibling.textContent = kind === 'font' ? `${value}px` : String(value);
        if (kind === 'font') {
          overlay.querySelectorAll('.anki-preview-content, .anki-preview-hint').forEach((el) => {
            el.style.fontSize = ankiFontSize(value);
          });
        }
      });
      input.addEventListener('change', () => { saveStudySettings(); });
    });
  });
  overlay.querySelectorAll('[data-anki-setting]').forEach((input) => {
    input.addEventListener('change', async () => {
      const setting = input.dataset.ankiSetting;
      state.studySettings[setting] = input.checked;
      if (input.checked && setting === 'swapKeyValue') state.studySettings.randomizeKeyValue = false;
      await saveStudySettings();
      refresh();
    });
  });
  overlay.querySelectorAll('[data-anki-nav]').forEach((button) => {
    button.addEventListener('click', () => {
      anki.previewIndex += Number(button.dataset.ankiNav);
      anki.previewFlipped = false;
      anki.previewHintOpen = false;
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
  target?.focus({ preventScroll: true });
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
      separator: 'newline',
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
    state.studyRecords = null;
    state.logDeleteError = '';
    state.logPreviewError = '';
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
    state.a4Order = state.pairs.map((pair) => pair.id);
    state.a4Query = '';
    await startStudy();
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
  const library = await readLibraryFile();
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
    const color = library?.colors?.[project.name];
    if (Number.isInteger(color) && color >= 0) project.cartridgeColorIndex = color;
    const savedMode = library?.dailyPaceModes?.[project.name];
    if (typeof savedMode === 'boolean') {
      project.plan = { ...(project.plan || {}), useRecommendedDailyPace: savedMode };
    }
    await saveProject(project);
  }
  if (recoveredProjects.length) {
    const projects = await listProjects();
    const savedNames = Array.isArray(library?.order) ? library.order : [];
    const idByName = new Map(projects.map((project) => [project.name, project.id]));
    const order = savedNames.length
      ? savedNames.map((name) => idByName.get(name)).filter(Boolean)
      : await getProjectOrder();
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

async function resetProjectsForDirectory() {
  const stale = await listProjects();
  for (const project of stale) await deleteProjectRecord(project.id);
  state.projects = [];
  state.projectOrder = [];
  state.project = null;
  await saveProjectOrderRecord([]);
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
      startIn: 'documents',
    });
    await saveWorkbookDirectory(directory);
    state.directoryHandle = directory;
    state.directorySaved = true;
    state.directoryConnected = true;
    state.directoryName = directory.name;
    state.directoryAccessIssue = '';
    state.logPreviewError = '';
    state.error = '';
    state.notice = L(`${directory.name} 폴더를 저장 위치로 연결했어요.`, `Connected the ${directory.name} folder as your save location.`);
    const library = await readLibraryFile();
    if (library) applyLibraryPreferences(library, language, state.studySettings);
    else state.libraryNeedsMigration = true;
    let prepared;
    try {
      await resetProjectsForDirectory();
      prepared = await prepareWorkbookDirectory(directory);
    } catch (error) {
      state.error = L(
        `폴더는 연결했지만 프로젝트를 불러오지 못했어요: ${error instanceof Error ? error.message : String(error)}`,
        `The folder was connected, but projects could not be loaded: ${error instanceof Error ? error.message : String(error)}`
      );
      render();
      return;
    }
    if (state.libraryNeedsMigration) await persistLibrary();
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
    if (state.libraryNeedsMigration) await persistLibrary();
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

function validEditColumnWidths(widths) {
  return Array.isArray(widths) &&
    widths.length === 3 &&
    widths.every((width) => Number.isFinite(width) && width > 0) &&
    Math.abs(widths.reduce((sum, width) => sum + width, 0) - 1) <= 0.001;
}

function loadEditColumnWidths() {
  if (state.editColumnWidths) return [...state.editColumnWidths];
  const saved = localStorage.getItem(EDIT_COLUMN_WIDTHS_KEY);
  if (saved === null) return [...DEFAULT_EDIT_COLUMN_WIDTHS];
  const widths = JSON.parse(saved);
  if (!validEditColumnWidths(widths)) {
    throw new Error(L('저장된 열 너비 설정이 올바르지 않습니다.', 'Saved column width settings are invalid.'));
  }
  return widths;
}

function saveEditColumnWidths(widths) {
  const normalized = normalizeColumnWidths(widths);
  state.editColumnWidths = normalized;
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

async function editProject(id, focusPairId = null) {
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
    state.studyRecords = null;
    state.logDeleteError = '';
    state.logPreviewError = '';
    state.fsrsData = synchronized.data;
    state.pairs = synchronized.pairs.map(normalizeStudyPair);
    state.sourceFile = file;
    state.editRows = rows;
    state.editDirty = false;
    state.editMessage = '';
    state.editMessageIsError = false;
    const focusPair = state.pairs.find((pair) => pair.id === focusPairId);
    let focusRow = focusPair
      ? rows.findIndex((row, index) => index > 0 && row[0] === focusPair.key && row[1] === focusPair.value && (row[2] ?? '') === focusPair.hint)
      : -1;
    if (focusPair && focusRow < 0) focusRow = Math.min(rows.length - 1, state.pairs.indexOf(focusPair) + 1);
    state.editFocus = focusRow > 0 ? { row: focusRow, col: 0 } : null;
    state.editHighlightRow = focusRow > 0 ? focusRow : null;
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
  state.editHighlightRow = null;
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
    state.studyRecords = null;
    state.logDeleteError = '';
    state.logPreviewError = '';
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
    <tr class="${rowIndex === 0 ? 'is-header-row' : ''}${rowIndex === state.editHighlightRow ? ' is-target-row' : ''}">
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
  const activePage = state.page === 'learning-status' ? 'status' : state.page;
  return `
    <header class="site-head">
      <button type="button" id="homeBtn" class="brand${projectsActive ? ' is-active' : ''}" aria-label="${L('프로젝트로 이동', 'Go to projects')}" title="${L('프로젝트로 이동합니다.', 'Go to projects.')}"${projectsActive ? ' aria-current="page"' : ''}>
        <p class="brand-name">ESPER</p>
        <span class="brand-acronym">Excel-based Spaced-repetition Program for Efficient Retention</span>
      </button>
      ${['workspace', 'home'].includes(state.page) ? '' : `<nav class="global-nav" aria-label="${L('주요 페이지','Main pages')}">
        <button type="button" data-global-nav="study"${activePage === 'study' ? ' class="is-active" aria-current="page"' : ''}${canNavigateProject ? '' : ' disabled'}>${L('학습하기','Study')}</button>
        <button type="button" data-global-nav="status"${activePage === 'status' ? ' class="is-active" aria-current="page"' : ''}${canNavigateProject ? '' : ' disabled'}>${L('학습 상태','Learning status')}</button>
        <button type="button" data-global-nav="a4"${activePage === 'a4' ? ' class="is-active" aria-current="page"' : ''}${canNavigateProject ? '' : ' disabled'}>A4 view</button>
      </nav>`}
      <div class="site-head-right">
        <a class="how-to-use-link" href="${TUTORIAL_PAGE_URL}" target="_blank" rel="noopener noreferrer">${L('튜토리얼', 'Tutorial')}</a>
        <a class="how-to-use-link" href="${SHARED_DECKS_URL}" target="_blank" rel="noopener noreferrer">${L('공유 덱', 'Shared Decks')}</a>
        ${folderConnectionHTML(true, ['workspace', 'home'].includes(state.page) ? state.error : '')}
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
          title: 'MCAT concept',
          sheet: 'MCAT',
          rows: [
              ['Henderson–Hasselbalch equation', 'pH = pKa + log₁₀([A⁻]/[HA])', 'At pH = pKa, [A⁻] = [HA].'],
              ['Michaelis–Menten kinetics', 'v = Vmax[S] / (Km + [S])', 'Km is [S] at ½ Vmax.'],
              ['Action potential', 'Na⁺ influx depolarizes; K⁺ efflux repolarizes.', { type: 'action-potential-diagram' }],
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
          title: 'MCAT 개념 암기',
          sheet: 'MCAT',
          rows: [
              ['Henderson–Hasselbalch equation', 'pH = pKa + log₁₀([A⁻]/[HA])', 'pH = pKa일 때 [A⁻]와 [HA]가 같아요.'],
              ['Michaelis–Menten kinetics', 'v = Vmax[S] / (Km + [S])', 'Km은 반응 속도가 ½Vmax일 때의 [S]예요.'],
              ['Action potential', 'Na⁺ 유입은 탈분극, K⁺ 유출은 재분극을 일으켜요.', { type: 'action-potential-diagram' }],
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
  const exampleValueHTML = (value, ci) => {
    if (value?.type === 'action-potential-diagram') {
      return `
        <span>${L('활동 전위 변화', 'Action potential')}</span>
        <svg class="format-example-diagram" viewBox="0 0 280 116" role="img" aria-label="${L('시간에 따른 활동 전위: 탈분극 후 재분극', 'Action potential over time: depolarization followed by repolarization')}">
          <path d="M34 12v78h232" fill="none" stroke="#8b929a" stroke-width="1.5"/>
          <path d="M35 67h230M35 35h230" fill="none" stroke="#e1e4e7" stroke-width="1" stroke-dasharray="3 4"/>
          <path d="M38 67h52c10 0 12-2 17-11l19-39c5-10 10-10 14 2l13 39c5 15 13 19 24 19h88" fill="none" stroke="#8f3f26" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
          <text x="5" y="18" class="diagram-axis">${L('전위', 'mV')}</text>
          <text x="13" y="71" class="diagram-axis">−70</text>
          <text x="20" y="39" class="diagram-axis">0</text>
          <text x="70" y="105" class="diagram-label">${L('휴지', 'Rest')}</text>
          <text x="105" y="105" class="diagram-label">${L('탈분극', 'Depolarization')}</text>
          <text x="188" y="105" class="diagram-label">${L('재분극', 'Repolarization')}</text>
        </svg>`;
    }
    return value ? esc(value) : (ci === 2 ? `<span class="sheet-optional-mark">${optionalMark}</span>` : '');
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
              ${example.rows.map((row, ri) => `<tr><th class="sheet-rownum">${ri + 2}</th>${row.map((value, ci) => cellHTML(ci, exampleValueHTML(value, ci))).join('')}</tr>`).join('')}
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

function folderConnectionHTML(compact = false, errorText = '') {
  const hasLocation = Boolean(state.directoryName && !state.directoryAccessIssue);
  const buttonLabel = hasLocation
    ? `${L('저장 위치', 'Save location')}: ${state.directoryName}`
    : `${L('저장 위치', 'Save location')}: ${L('미정', 'Not set')}`;
  const warningText = L('이전에 저장한 내용을 불러오려면 같은 저장 위치를 선택해야 해요.', 'To load previously saved content, choose the same save location as before.');
  const buttonTitle = hasLocation ? `${buttonLabel}. ${warningText}` : L('저장 위치 선택', 'Choose save location');
  const selectionButton = `<button type="button" id="connectFolderBtn" class="btn folder-location-button" title="${esc(buttonTitle)}">${esc(buttonLabel)}</button>`;
  const warning = state.directoryName && !compact
    ? `<span class="folder-location-warning">${warningText}</span>`
    : '';

  return `
    <div class="folder-connection${compact ? ' folder-connection-compact' : ''}${hasLocation ? ' has-location' : ' no-location'}">
      <div class="folder-connection-action">
        ${selectionButton}
        ${warning}
      </div>
      ${errorText ? `<p class="folder-connection-error" role="alert">${esc(errorText)}</p>` : ''}
    </div>`;
}

function folderPromptHTML() {
  if (state.directorySaved || state.folderPromptDismissed) return '';
  return `
    <dialog class="folder-prompt" id="folderPrompt" aria-labelledby="folderPromptTitle">
      <div class="folder-prompt-card">
        <h2 id="folderPromptTitle">${L('프로젝트를 저장할 폴더를 선택해 주세요', 'Choose a folder to save your projects')}</h2>
        <p>${L(
          '다음 창에서 고른 폴더가 앞으로 모든 프로젝트와 학습 기록이 저장되는 위치가 돼요. 문서나 바탕화면에 ESPER 전용 폴더를 하나 만들어 선택하는 것을 추천해요.',
          'The folder you pick in the next window will be where all your projects and study records are saved from now on. We recommend creating a dedicated ESPER folder in Documents or on your Desktop.'
        )}</p>
        <div class="folder-prompt-actions">
          <button type="button" class="btn" data-folder-prompt-later>${L('나중에', 'Later')}</button>
          <button type="button" class="btn btn-primary" data-folder-prompt-choose>${L('폴더 선택', 'Choose folder')}</button>
        </div>
      </div>
    </dialog>`;
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
    ${folderPromptHTML()}
    <main class="shell workspace-screen">
      ${formatGuideHTML()}
      ${state.notice ? `<p class="notice" role="status">${esc(state.notice)}</p>` : ''}
      <div class="workspace-heading">
        <div class="workspace-heading-copy">
          <p class="workspace-eyebrow">${L('나의 학습 공간','YOUR STUDY SPACE')}</p>
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
      let storedDailyProgress = null;
      if (stored !== null) {
        const data = JSON.parse(stored);
        storedDailyProgress = data?.dailyProgress || null;
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
      const projectRecords = project.pairs.map((pair) => activeById.get(pair.id)).filter(Boolean);
      const dueReviews = reviewedCards.filter((record) => Date.parse(record.card.dueAt) <= now).length;
      const dueCount = dueReviews + dailyNewAllowance(project, projectRecords, newCount, storedDailyProgress).availableNewCards;
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

function formatFsrsDate(value) {
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) return L('날짜 없음', 'No date');
  return new Intl.DateTimeFormat(language === 'ko' ? 'ko-KR' : 'en', {
    dateStyle: 'medium',
    timeStyle: 'short',
  }).format(date);
}

function formatStudyTime(milliseconds) {
  const seconds = Math.max(0, Math.floor(milliseconds / 1000));
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds % 60).padStart(2, '0')}`;
}

const RETENTION_WARNING_THRESHOLD = 95;

// Exam date and target retention belong to each project. Projects created before
// this change fall back to the old global values.
function projectPlan(project = state.project) {
  const plan = project?.plan;
  if (plan) {
    const examDate = /^\d{4}-\d{2}-\d{2}$/.test(plan.examDate || '') ? plan.examDate : '';
    return {
      examDate,
      desiredRetention: normalizeDesiredRetention(plan.desiredRetention ?? DEFAULT_DESIRED_RETENTION),
      dailyNewCards: Math.min(1000, Math.max(1, Math.floor(Number(plan.dailyNewCards) || DEFAULT_DAILY_NEW_CARDS))),
      useRecommendedDailyPace: plan.useRecommendedDailyPace !== false,
      firstPassDeadline: /^\d{4}-\d{2}-\d{2}$/.test(plan.firstPassDeadline || '')
        ? plan.firstPassDeadline
        : firstPassDeadline(examDate),
    };
  }
  const examDate = state.studySettings.examDate;
  return {
    examDate,
    desiredRetention: state.studySettings.desiredRetention,
    dailyNewCards: DEFAULT_DAILY_NEW_CARDS,
    useRecommendedDailyPace: true,
    firstPassDeadline: firstPassDeadline(examDate),
  };
}

function selectedDailyPace(plan, remainingCards, now = new Date()) {
  if (!plan.useRecommendedDailyPace) {
    return manualDailyNewCardPlan({ remainingCards, dailyCards: plan.dailyNewCards, now });
  }
  return dailyNewCardPlan({
    examDate: plan.examDate,
    firstPassDate: plan.firstPassDeadline,
    remainingCards,
    configuredDailyCards: DEFAULT_DAILY_NEW_CARDS,
    now,
  });
}

function dailyNewAllowance(project, activeRecords, newCardCount, dailyProgress) {
  const today = localStudyDate();
  const learnedToday = new Set(activeRecords
    .filter((record) => record.reviewLogs?.[0] && localStudyDate(new Date(record.reviewLogs[0].reviewTime)) === today)
    .map((record) => record.id));
  const pace = selectedDailyPace(projectPlan(project), newCardCount);
  const extraNewCards = dailyProgress?.date === today && Number.isInteger(dailyProgress.extraNewCards)
    ? Math.max(0, dailyProgress.extraNewCards)
    : 0;
  const canLearnNew = !pace.isPastDeadline;
  const allowance = canLearnNew ? pace.recommended + extraNewCards : 0;
  return {
    today,
    learnedToday,
    extraNewCards,
    allowance,
    pace,
    canLearnNew,
    availableNewCards: canLearnNew ? Math.min(newCardCount, Math.max(0, allowance - learnedToday.size)) : 0,
  };
}

function currentDailyNewProgress() {
  const pairIds = new Set(state.pairs.map((pair) => pair.id));
  const activeRecords = (state.fsrsData?.cards || []).filter((record) =>
    record.archivedAt === null && pairIds.has(record.id)
  );
  const newCards = activeRecords.filter((record) => record.card.state === FsrsState.New);
  const daily = dailyNewAllowance(state.project, activeRecords, newCards.length, state.fsrsData?.dailyProgress);
  return {
    today: daily.today,
    learnedToday: daily.learnedToday.size,
    remainingNewCards: newCards.length,
    extraNewCards: daily.extraNewCards,
    allowance: daily.allowance,
    recommended: daily.pace.recommended,
    completionDate: daily.pace.completionDate,
    isPastDeadline: daily.pace.isPastDeadline,
    newCards,
    learnedIds: daily.learnedToday,
    availableNewCards: daily.availableNewCards,
    canOfferExtra: daily.canLearnNew && newCards.length > 0 && daily.learnedToday.size >= daily.allowance,
  };
}

function retentionWarningText() {
  return L('100%에 가까워질수록 복습양이 기하급수적으로 증가합니다. 시작은 90%로 두길 권장합니다', 'Review load grows exponentially as you approach 100%. Starting at 90% is recommended');
}

function retentionWarningHTML(percent, attr) {
  return `<p class="retention-warning" ${attr} role="status" hidden>${retentionWarningText()}</p>`;
}

function examDdayLabel(examDate = projectPlan().examDate) {
  const date = parseLocalDate(examDate);
  if (!date) return L('시험일 미설정', 'No exam date');
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const days = Math.round((date.getTime() - today.getTime()) / 86400000);
  return days === 0 ? 'D-Day' : days > 0 ? `D-${days}` : `D+${-days}`;
}

function examDdayHTML(label) {
  return String(label).replace(/(\d+)/, '<em class="plan-accent">$1</em>');
}

async function refreshExamSummary() {
  const elements = [...appEl.querySelectorAll('[data-exam-retention]')];
  if (!elements.length) return;
  const { examDate, desiredRetention } = projectPlan();
  const set = (text) => elements.forEach((element) => { if (element.isConnected) element.textContent = text; });
  const examEnd = /^\d{4}-\d{2}-\d{2}$/.test(examDate || '') ? new Date(`${examDate}T23:59:59`) : null;
  if (!examEnd) {
    set('—');
    return;
  }
  const records = (state.fsrsData?.cards || []).filter((record) => record.archivedAt === null);
  try {
    const summary = await summarizeExamReadiness(records, examEnd, desiredRetention);
    if (projectPlan().examDate !== examDate) return;
    set(summary.daysLeft < 0 || summary.averageRetention === null ? '—' : `${Math.round(summary.averageRetention * 100)}%`);
  } catch {
    set('—');
  }
}

function studyFaceHasImage(target = 'primary') {
  const session = state.study;
  const pair = session?.started ? studyCurrentPair() : null;
  if (!pair) return true;
  const shown = studyShownPair(pair);
  const has = (text) => /\[img:/.test(String(text ?? ''));
  if (target === 'hint') return has(pair.hint);
  return session.flipped
    ? has(shown.value)
    : has(shown.key);
}

function studySettingsHTML() {
  const settings = state.studySettings;
  const side = currentStudySide();
  const hintField = side === 'front' ? 'hintFront' : 'hintBack';
  const primaryLabel = side === 'front' ? 'key' : 'value';
  return `
    <div class="study-settings-panel" id="studySettingsPanel" aria-label="${L('설정', 'Settings')}">
      <section class="study-setting-card">
        <h2>${side === 'front' ? L('카드 크기 · 앞면', 'Card size · Front') : L('카드 크기 · 뒷면', 'Card size · Back')}</h2>
        ${cardSizeControlsHTML(side, side, L(primaryLabel, primaryLabel), studyFaceHasImage())}
        ${cardSizeControlsHTML(side, hintField, L('hint', 'hint'), studyFaceHasImage('hint'))}
        <p class="study-font-setting-note">${L('글이 길면 카드 안에 맞게 자동으로 작아져요.', 'Long text shrinks to fit the card.')}</p>
      </section>
      <section class="study-setting-card">
        <h2>${L('회차 시간 제한', 'Round time limit')}</h2>
        <label class="study-hint-setting study-time-limit">
          <span>${L('한 회차 최대 시간', 'Maximum time per round')}</span>
          <span class="study-time-limit-control">
            <input class="study-time-limit-input" type="number" min="${STUDY_TIME_LIMIT_RANGE.min}" max="${STUDY_TIME_LIMIT_RANGE.max}" step="1" data-study-setting="roundTimeLimitMinutes" value="${settings.roundTimeLimitMinutes}" aria-label="${L('회차 시간 제한(분)', 'Round time limit (minutes)')}">
            <span>${L('분', 'min')}</span>
          </span>
        </label>
        <small class="study-time-limit-note">${L('제한 시간에 도달하면 회차 기록을 저장하고 자동 종료해요.', 'The round is automatically ended and saved when the time limit is reached.')}</small>
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
      <div class="study-edit-card-row">
        <button type="button" class="btn study-edit-card-button" data-study-action="edit-card"${state.study?.started && !state.studySaving ? '' : ' disabled'} title="${L('학습을 종료하고 기록을 저장한 뒤, 지금 보던 카드의 행을 편집 페이지에서 열어요.', 'Ends the round, saves its record, and opens this card\'s row on the edit page.')}">✎ ${L('카드 바로 수정', 'Edit this card')}</button>
      </div>
    </div>`;
}

function studyHTML() {
  const session = state.study;
  const queue = buildStudyQueue();
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
          `복습 예정 ${queue.dueReviewIds.length}개 · 오늘의 신규 카드 ${queue.newCardIds.length}개를 준비했어요.`,
          `${queue.dueReviewIds.length} reviews due · ${queue.newCardIds.length} new cards ready for today.`
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
        <p>${session ? '' : queue.progress.isPastDeadline && queue.progress.remainingNewCards
          ? L('1회독 기간이 끝나 이제 새 카드 없이 복습만 진행해요.', 'The first-pass period has ended. Only reviews are available now.')
          : overview.nextDueAt
          ? L(`다음 복습 예정: ${formatFsrsDate(overview.nextDueAt)}`, `Next review: ${formatFsrsDate(overview.nextDueAt)}`)
          : L('모든 카드가 복습되었어요. 새 항목을 추가하면 FSRS가 복습 일정을 시작해요.', 'All cards have been reviewed. Add items to start an FSRS schedule.')}</p>
        ${!session && queue.progress.canOfferExtra
          ? `<button type="button" class="btn btn-primary" data-study-action="more-new"${state.studySaving ? ' disabled' : ''}>${L('신규 카드 10장 더 학습하기','Study 10 more new cards')}</button>`
          : `<button type="button" class="btn" data-study-action="restart">${L('다시 학습하기','Study again')}</button>`}
      </div>`;
  const shortcuts = Object.entries(session?.settings.shortcuts || state.studySettings.shortcuts)
    .map(([action, key]) => `${esc(shortcutKeyLabel(key))} ${esc(studyActionLabel(STUDY_ACTIONS.find(([id]) => id === action) || [action, '']))}`)
    .join(' · ');
  return `
    ${headerHTML()}
    <main class="study-screen study-workspace${studySidebarCollapsed ? ' study-sidebar-collapsed' : ''}${state.studyFullscreen && session?.started ? ' is-fullscreen' : ''}">
      <button type="button" id="reopenStudySidebarBtn" class="study-sidebar-reopen" aria-label="${L('설정 펼치기','Expand settings')}" title="${L('설정 펼치기','Expand settings')}"${studySidebarCollapsed ? '' : ' hidden'}>
        <svg aria-hidden="true" viewBox="0 0 24 24" fill="none">
          <path d="M19.14 12.94c.04-.3.06-.61.06-.94s-.02-.64-.07-.94l2.03-1.58a.5.5 0 0 0 .12-.64l-1.92-3.32a.5.5 0 0 0-.61-.22l-2.39.96a7.1 7.1 0 0 0-1.63-.94l-.36-2.54a.49.49 0 0 0-.49-.42h-3.84a.49.49 0 0 0-.49.42l-.36 2.54c-.59.23-1.14.55-1.63.94l-2.39-.96a.49.49 0 0 0-.61.22L2.63 8.84a.5.5 0 0 0 .12.64l2.03 1.58c-.05.3-.08.62-.08.94s.03.64.08.94l-2.03 1.58a.5.5 0 0 0-.12.64l1.92 3.32c.12.21.37.3.61.22l2.39-.96c.49.39 1.04.71 1.63.94l.36 2.54c.04.24.24.42.49.42h3.84c.25 0 .45-.18.49-.42l.36-2.54c.59-.23 1.14-.55 1.63-.94l2.39.96c.23.09.49-.01.61-.22l1.92-3.32a.5.5 0 0 0-.12-.64l-2.01-1.56ZM12 15.5a3.5 3.5 0 1 1 0-7 3.5 3.5 0 0 1 0 7Z" fill="currentColor"/>
        </svg>
      </button>
      <aside class="study-sidebar" aria-label="${L('설정','Settings')}">
        <div class="study-sidebar-topbar">
        </div>
        <button type="button" id="collapseStudySidebarBtn" class="study-sidebar-collapse" aria-label="${L('설정 접기','Collapse settings')}" title="${L('설정 접기','Collapse settings')}">
          <svg aria-hidden="true" viewBox="0 0 24 24" fill="none"><path d="M19 12H5m6-6-6 6 6 6" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
        </button>
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
          ${session.history.length && !state.pendingReview ? `<button type="button" class="btn study-undo-button" data-study-action="undo"${state.studySaving ? ' disabled' : ''} title="${esc(shortcutKeyLabel(state.studySettings.shortcuts.undo))}">↩ ${L('선택 되돌리기', 'Undo rating')}</button>` : ''}
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

function studySessionSeconds(session) {
  return session.durationSeconds || 0;
}

function studyLogStatsHTML(sessions, periodLabel) {
  const items = sessions.reduce((sum, session) => sum + session.itemCount, 0);
  const seconds = sessions.reduce((sum, session) => sum + studySessionSeconds(session), 0);
  return `<div class="study-log-stats">
    <div><span>${L(`${periodLabel} 학습 카드 수`, `${periodLabel} cards studied`)}</span><strong>${items}${L('개', '')}</strong></div>
    <div><span>${L(`${periodLabel} 학습 시간`, `${periodLabel} study time`)}</span><strong>${formatStudyTime(seconds * 1000)}</strong></div>
  </div>`;
}

function studyLogPeriodNavHTML() {
  if (state.logPreviewView === 'monthly') {
    const [year, month] = state.logCalMonth.split('-').map(Number);
    const label = L(`${year}년 ${month}월`, `${['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][month - 1]} ${year}`);
    return `<div class="study-log-period-nav">
      <button type="button" data-log-step="month:-1" aria-label="${L('이전 달', 'Previous month')}">‹</button>
      <strong>${label}</strong>
      <button type="button" data-log-step="month:1" aria-label="${L('다음 달', 'Next month')}">›</button>
    </div>`;
  }
  return `<div class="study-log-period-nav">
    <button type="button" data-log-step="year:-1" aria-label="${L('이전 해', 'Previous year')}">‹</button>
    <strong>${state.logYear}</strong>
    <button type="button" data-log-step="year:1" aria-label="${L('다음 해', 'Next year')}">›</button>
  </div>`;
}

function studyLogDailySessionsHTML(sessions) {
  if (!sessions.length) return `<p class="study-log-empty">${L('이 날짜의 학습 기록이 없어요.', 'No study sessions for this day.')}</p>`;
  return `<div class="study-log-session-list">${sessions.map(({ session, index }, round) => `
    <article class="study-log-session">
      <strong>${L(`${round + 1}회차`, `Round ${round + 1}`)}</strong>
      <span>${L('학습 카드', 'Cards studied')} <b>${session.itemCount}${L('개', '')}</b></span>
      <span>${L('학습 시간', 'Study time')} <b>${formatStudyTime(studySessionSeconds(session) * 1000)}</b></span>
      <button type="button" class="study-log-delete" data-log-delete="${index}" aria-label="${state.logSessionDeleting ? L('삭제 중', 'Deleting session') : L('회차 삭제', 'Delete session')}" title="${state.logSessionDeleting ? L('삭제 중', 'Deleting session') : L('회차 삭제', 'Delete session')}"${state.logSessionDeleting ? ' disabled' : ''}>×</button>
    </article>`).join('')}</div>`;
}

function studyLogCalendarHTML(sessions, selectedDate) {
  const [year, month] = state.logCalMonth.split('-').map(Number);
  const prefix = `${state.logCalMonth}-`;
  const dailyCounts = new Map();
  for (const session of sessions) {
    if (!session.date.startsWith(prefix)) continue;
    const entry = dailyCounts.get(session.date) || { items: 0, rounds: 0 };
    entry.items += session.itemCount;
    entry.rounds += 1;
    dailyCounts.set(session.date, entry);
  }
  const maxActivity = Math.max(1, ...[...dailyCounts.values()].map(({ items, rounds }) => Math.max(items, rounds)));
  const level = (entry) => entry
    ? Math.min(4, Math.max(1, Math.ceil((Math.max(entry.items, entry.rounds) / maxActivity) * 4)))
    : 0;
  const cells = [];
  for (let blank = 0; blank < new Date(year, month - 1, 1).getDay(); blank += 1) {
    cells.push('<span class="study-log-calendar-cell is-empty"></span>');
  }
  for (let day = 1; day <= new Date(year, month, 0).getDate(); day += 1) {
    const date = `${prefix}${String(day).padStart(2, '0')}`;
    const entry = dailyCounts.get(date);
    const items = entry?.items || 0;
    const dateLabel = L(`${month}월 ${day}일`, `${date}`);
    const activityLabel = items ? L(`, ${items}개 학습`, `, ${items} cards studied`) : '';
    cells.push(`<button type="button" class="study-log-calendar-cell level-${level(entry)}${date === selectedDate ? ' is-selected' : ''}${date === localDateForExport() ? ' is-today' : ''}" data-log-date="${date}" aria-label="${dateLabel}${activityLabel}" aria-pressed="${date === selectedDate}"${entry ? ` title="${dateLabel} · ${items}${L('개 학습', ' cards studied')}"` : ` title="${dateLabel}"`}>${day}</button>`);
  }
  return `<section class="study-log-calendar" aria-label="${L('월간 학습 달력', 'Monthly study calendar')}">
    <div class="study-log-calendar-weekdays">${L(['일', '월', '화', '수', '목', '금', '토'], ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']).map((day) => `<span>${day}</span>`).join('')}</div>
    <div class="study-log-calendar-grid">${cells.join('')}</div>
  </section>`;
}

function studyLogYearGrassHTML(sessions) {
  const dailyCounts = new Map();
  for (const session of sessions) {
    const entry = dailyCounts.get(session.date) || { items: 0, rounds: 0 };
    entry.items += session.itemCount;
    entry.rounds += 1;
    dailyCounts.set(session.date, entry);
  }
  const start = new Date(0);
  start.setFullYear(state.logYear, 0, 1);
  start.setHours(0, 0, 0, 0);
  start.setDate(start.getDate() - start.getDay());
  const end = new Date(0);
  end.setFullYear(state.logYear, 11, 31);
  end.setHours(0, 0, 0, 0);
  const weeks = [];
  const maxActivity = Math.max(1, ...[...dailyCounts]
    .filter(([date]) => date.startsWith(`${state.logYear}-`))
    .map(([, { items, rounds }]) => Math.max(items, rounds)));
  const level = (entry) => entry
    ? Math.min(4, Math.max(1, Math.ceil((Math.max(entry.items, entry.rounds) / maxActivity) * 4)))
    : 0;
  const monthLabels = L(['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'], ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']);
  for (let cursor = new Date(start); cursor <= end; cursor.setDate(cursor.getDate() + 7)) {
    const weekStart = new Date(cursor);
    let monthLabel = '';
    const days = [];
    for (let weekday = 0; weekday < 7; weekday += 1) {
      const date = new Date(weekStart);
      date.setDate(date.getDate() + weekday);
      if (date.getFullYear() !== state.logYear) {
        days.push('<span class="study-log-grass-cell is-empty"></span>');
        continue;
      }
      if (date.getDate() === 1) monthLabel = monthLabels[date.getMonth()];
      const key = `${state.logYear}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
      const entry = dailyCounts.get(key);
      const items = entry?.items || 0;
      const dateLabel = `${key} · ${items}${L('개 학습', ' cards studied')}`;
      days.push(`<span class="study-log-grass-cell level-${level(entry)}" title="${dateLabel}" aria-label="${dateLabel}"></span>`);
    }
    weeks.push(`<div class="study-log-grass-week"><span class="study-log-grass-month">${monthLabel}</span>${days.join('')}</div>`);
  }
  return `<div class="study-log-grass-scroll"><div class="study-log-grass" role="group" aria-label="${L(`${state.logYear}년 학습 잔디`, `${state.logYear} study activity`)}">${weeks.join('')}</div><div class="study-log-grass-legend">${L('적게', 'Less')} <span class="study-log-grass-cell level-1"></span><span class="study-log-grass-cell level-2"></span><span class="study-log-grass-cell level-3"></span><span class="study-log-grass-cell level-4"></span> ${L('많이', 'More')}</div></div>`;
}

function studyLogPreviewBodyHTML() {
  if (state.logPreviewError) return `<p class="study-log-preview-empty is-error" role="alert">${esc(state.logPreviewError)}</p>`;
  const records = state.studyRecords;
  if (!records) return `<p class="study-log-preview-empty" role="status">${L('불러오는 중…', 'Loading…')}</p>`;
  const sessions = records.sessions;
  const error = state.logDeleteError ? `<p class="study-log-action-error" role="alert">${esc(state.logDeleteError)}</p>` : '';

  if (state.logPreviewView === 'monthly') {
    const selected = state.logSelectedDate || localDateForExport();
    const monthSessions = sessions.filter((session) => session.date.startsWith(state.logCalMonth));
    const selectedSessions = sessions
      .map((session, index) => ({ session, index }))
      .filter(({ session }) => session.date === selected);
    const dailySessions = selectedSessions.map(({ session }) => session);
    return `<div class="study-log-monthly">
      <div class="study-log-monthly-calendar">
        ${studyLogPeriodNavHTML()}
        ${studyLogCalendarHTML(sessions, selected)}
      </div>
      <section class="study-log-day-detail" aria-label="${L('선택한 날짜의 학습 기록', 'Study sessions for selected day')}">
        <h3>${selected}</h3>
        ${error}
        ${studyLogDailySessionsHTML(selectedSessions)}
      </section>
      <aside class="study-log-day-total" aria-label="${L('일일 누적 기록', 'Daily totals')}">
        <h3>${L('일일 누적', 'Daily total')}</h3>
        ${studyLogStatsHTML(dailySessions, L('일일', 'Daily'))}
        <p>${L(`${dailySessions.length}회차`, `${dailySessions.length} study rounds`)}</p>
      </aside>
      <aside class="study-log-month-total" aria-label="${L('이번 달 누적 기록', 'Monthly totals')}">
        <h3>${L('이번 달 누적', 'Monthly total')}</h3>
        ${studyLogStatsHTML(monthSessions, L('월간', 'Monthly'))}
        <p>${L(`${monthSessions.length}회차`, `${monthSessions.length} study rounds`)}</p>
      </aside>
    </div>`;
  }

  const yearSessions = sessions.filter((session) => Number(session.date.slice(0, 4)) === state.logYear);
  const studyDays = new Set(yearSessions.map((session) => session.date)).size;
  return `<div class="study-log-yearly">
    <div class="study-log-yearly-activity">
      ${studyLogPeriodNavHTML()}
      ${studyLogYearGrassHTML(yearSessions)}
    </div>
    <aside class="study-log-year-total" aria-label="${L('연간 누적 기록', 'Yearly totals')}">
      <h3>${L('연간 누적', 'Yearly total')}</h3>
      ${studyLogStatsHTML(yearSessions, L('연간', 'Yearly'))}
      <p>${L(`${studyDays}일 학습 · ${yearSessions.length}회차`, `${studyDays} study days · ${yearSessions.length} rounds`)}</p>
    </aside>
    ${error}
  </div>`;
}

function studyLogPreviewHTML() {
  const failed = Boolean(state.logPreviewError) || !state.studyRecords;
  const view = state.logPreviewView;
  const toggle = (id, label) => `<button type="button" class="learning-status-view-button${view === id ? ' is-active' : ''}" data-log-view="${id}" aria-pressed="${view === id}"${failed ? ' disabled' : ''}>${label}</button>`;
  return `<section class="study-log-preview ls-section${failed ? ' is-disabled' : ''}${state.collapsedSections.log ? ' is-collapsed' : ''}" data-ls-section="log" data-log-preview aria-label="${L('나의 기록', 'My Log')}">
        <div class="study-log-preview-title">
          ${sectionToggleHTML('log', L('나의 기록', 'MY LOG'))}
          <div class="learning-status-view-switch" role="group" aria-label="${L('보기 방식', 'View mode')}">${toggle('monthly', L('월별', 'Monthly'))}${toggle('yearly', L('연별', 'Yearly'))}</div>
        </div>
        <div class="study-log-preview-body" data-log-preview-body>${studyLogPreviewBodyHTML()}</div>
      </section>`;
}

function bindStudyLogPreview() {
  bindSectionToggles();
  appEl.querySelectorAll('[data-log-view]').forEach((button) => button.addEventListener('click', () => {
    if (button.disabled) return;
    state.logPreviewView = button.dataset.logView;
    const card = appEl.querySelector('[data-log-preview]');
    if (!card) return;
    card.outerHTML = studyLogPreviewHTML();
    bindStudyLogPreview();
    persistUserPreferences();
  }));
  const refreshBody = () => {
    const body = appEl.querySelector('[data-log-preview-body]');
    if (body) body.innerHTML = studyLogPreviewBodyHTML();
    bindStudyLogActions();
  };
  async function deleteStudySession(index) {
    if (state.logSessionDeleting || !state.studyRecords || !state.project) return;
    if (!window.confirm(L(
      '이 My Log 회차 기록만 삭제할까요? FSRS 복습 기록은 삭제되지 않습니다.',
      'Delete only this My Log session? FSRS review history will not be deleted.'
    ))) return;
    const projectName = state.project.name;
    const currentRecords = state.studyRecords;
    state.logSessionDeleting = true;
    state.logDeleteError = '';
    render();
    try {
      const nextRecords = removeStudySession(currentRecords, index);
      await writeStudyRecordFile(projectName, nextRecords);
      if (state.project?.name === projectName) state.studyRecords = nextRecords;
    } catch (error) {
      console.error('Could not delete the study session.', error);
      if (state.project?.name === projectName) {
        state.logDeleteError = L(
          `학습 회차 기록을 삭제하지 못했어요: ${error instanceof Error ? error.message : String(error)}`,
          `Could not delete the study session: ${error instanceof Error ? error.message : String(error)}`
        );
      }
    } finally {
      state.logSessionDeleting = false;
      render();
    }
  }
  function bindStudyLogActions() {
    appEl.querySelectorAll('[data-log-date]').forEach((button) => button.addEventListener('click', () => {
      state.logSelectedDate = button.dataset.logDate;
      refreshBody();
    }));
    appEl.querySelectorAll('[data-log-step]').forEach((button) => button.addEventListener('click', () => {
      const [unit, amount] = button.dataset.logStep.split(':');
      const step = Number(amount);
      if (unit === 'month') {
        const [year, month] = state.logCalMonth.split('-').map(Number);
        const next = new Date(0);
        next.setFullYear(year, month - 1 + step, 1);
        state.logCalMonth = `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, '0')}`;
        const selectedDay = Number((state.logSelectedDate || localDateForExport()).slice(8, 10));
        const [nextYear, nextMonth] = state.logCalMonth.split('-').map(Number);
        const day = Math.min(selectedDay, new Date(nextYear, nextMonth, 0).getDate());
        state.logSelectedDate = `${state.logCalMonth}-${String(day).padStart(2, '0')}`;
      } else if (unit === 'year') {
        state.logYear = Math.max(1, Math.min(9999, state.logYear + step));
      }
      refreshBody();
    }));
    appEl.querySelectorAll('[data-log-delete]').forEach((button) => button.addEventListener('click', () => {
      deleteStudySession(Number(button.dataset.logDelete));
    }));
  }
  bindStudyLogActions();
}

async function loadStudyLogPreview() {
  if (state.studyRecords || state.logPreviewLoading || state.logPreviewError || !state.project) return;
  state.logPreviewLoading = true;
  try {
    const file = await readStudyRecordFile(state.project.name);
    state.studyRecords = file || emptyStudyRecords();
  } catch (error) {
    console.error('Could not load study records.', error);
    state.logPreviewError = L('학습 기록을 불러오지 못했어요: ', 'Could not load study records: ') + (error instanceof Error ? error.message : String(error));
  } finally {
    state.logPreviewLoading = false;
    const card = state.page === 'learning-status' ? appEl.querySelector('[data-log-preview]') : null;
    if (card) {
      card.outerHTML = studyLogPreviewHTML();
      bindStudyLogPreview();
    }
  }
}

function retentionStyleDialogHTML() {
  const { color, strength, half } = state.studySettings.retentionStyle;
  const percents = [100, 85, 70, 55, 40, 20, 0];
  return `
    <dialog class="retention-dialog" aria-labelledby="retentionDialogTitle">
      <form method="dialog" class="retention-dialog-body">
        <div class="retention-dialog-head">
          <h2 id="retentionDialogTitle">${L('카드 투명도 설정', 'Card opacity')}</h2>
          <button type="submit" class="fsrs-guide-close" aria-label="${L('닫기', 'Close')}">×</button>
        </div>
        <p class="retention-dialog-lead">${L(
          '현재 기억 확률이 높을수록 카드가 진하게, 낮을수록 흐리게 보여요. 아래 샘플은 기억 확률별로 카드가 어떻게 보이는지 보여 줘요.',
          'The higher the current recall probability, the more opaque the card. The samples below show how cards look at each probability.'
        )}</p>
        <div class="retention-preview" data-retention-preview>${percents.map((percent) => `
          <div class="retention-swatch"><span class="retention-swatch-card" data-retention-swatch="${percent}"><b>${percent}%</b></span></div>`).join('')}
        </div>
        <div class="retention-field retention-color-field">
          <span>${L('색상', 'Color')}</span>
          <div class="retention-picker">
            <button type="button" class="retention-color-swatch" data-retention-swatch-btn aria-haspopup="true" aria-expanded="false" aria-label="${L('색상 선택', 'Pick a color')}" style="background:${color}"></button>
            <input type="text" class="retention-hex" data-retention-hex value="${color}" maxlength="7" spellcheck="false" autocomplete="off" placeholder="#8f3f26" aria-label="HEX">
            <div class="retention-picker-pop" data-retention-pop hidden>
              <div class="retention-sv" data-retention-sv><i data-retention-sv-handle></i></div>
              <input type="range" class="retention-hue" min="0" max="360" step="1" data-retention-hue aria-label="${L('색조', 'Hue')}">
            </div>
          </div>
        </div>
        <label class="retention-field"><span>${L('절대 투명도', 'Absolute opacity')}</span><input type="range" min="5" max="100" step="1" data-retention-strength value="${strength}" aria-label="${L('절대 투명도', 'Absolute opacity')}"><output data-retention-strength-out>${strength}%</output></label>
        <p class="retention-hint">${L('높일수록 모든 카드가 전체적으로 진해져요.', 'Higher values make every card darker overall.')}</p>
        <div class="retention-field retention-half-field">
          <span>${L('상대 투명도', 'Relative opacity')}</span>
          <div class="retention-half">
            <div class="retention-half-bar" data-retention-half-bar><b data-retention-half-arrow>▲</b></div>
            <div class="retention-half-scale"><span>100%</span><span>50%</span><span>0%</span></div>
          </div>
        </div>
        <p class="retention-hint" data-retention-half-note></p>
        <div class="retention-dialog-actions">
          <button type="button" class="btn" data-retention-reset>${L('기본값으로', 'Reset')}</button>
        </div>
      </form>
    </dialog>`;
}

function sectionToggleHTML(id, label, tag = 'span') {
  const open = !state.collapsedSections[id];
  return `<${tag} class="study-ready-eyebrow ls-section-title"><button type="button" class="ls-section-toggle" data-ls-toggle="${id}" aria-expanded="${open}"><i aria-hidden="true">▾</i>${label}</button></${tag}>`;
}

function bindSectionToggles() {
  appEl.querySelectorAll('[data-ls-toggle]').forEach((button) => {
    button.onclick = async () => {
      const id = button.dataset.lsToggle;
      state.collapsedSections[id] = !state.collapsedSections[id];
      button.setAttribute('aria-expanded', String(!state.collapsedSections[id]));
      button.closest('[data-ls-section]')?.classList.toggle('is-collapsed', Boolean(state.collapsedSections[id]));
      await persistUserPreferences();
    };
  });
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
  const recordsById = new Map(records.map((item) => [item.record.id, item]));
  const sourceOrderedRecords = state.pairs.flatMap((pair) => {
    const item = recordsById.get(pair.id);
    return item ? [item] : [];
  });
  const counts = {
    all: records.length,
    due: records.filter((item) => item.category === 'due').length,
    new: records.filter((item) => item.category === 'new').length,
    scheduled: records.filter((item) => item.category === 'scheduled').length,
  };
  const filters = [
    ['due', L('복습 예정', 'To review'), counts.due],
    ['scheduled', L('학습 완료', 'Learned'), counts.scheduled],
    ['new', L('학습 전', 'Not started'), counts.new],
  ];
  const sortedRecords = records.sort((a, b) => {
    const order = { due: 0, new: 1, scheduled: 2 };
    return order[a.category] - order[b.category] ||
      Date.parse(a.record.card.dueAt) - Date.parse(b.record.card.dueAt);
  });
  const statusLabels = {
    due: L('복습할 때예요', 'Review due'),
    new: L('학습 전', 'Not started'),
    scheduled: L('학습 완료', 'Learned'),
  };
  const plan = projectPlan();
  const pace = selectedDailyPace(plan, counts.new);
  const completionDate = pace.completionDate
    ? new Intl.DateTimeFormat(language === 'ko' ? 'ko-KR' : 'en', { dateStyle: 'medium' })
      .format(parseLocalDate(pace.completionDate))
    : '—';
  return `
    ${headerHTML()}
    <main class="learning-status-screen">
      <section class="ls-section${state.collapsedSections.plan ? ' is-collapsed' : ''}" data-ls-section="plan">
      <header class="learning-status-heading">
        ${sectionToggleHTML('plan', L('학습설정', 'LEARNING SETTINGS'), 'p')}
        ${invalidPairCount ? `<p class="error" role="alert">${L(
          `학습 항목 ${invalidPairCount}개에서 카드 ID를 찾지 못해 상태 목록에서 제외했어요.`,
          `${invalidPairCount} study item(s) have no card ID and were excluded from the status list.`
        )}</p>` : ''}
      </header>
      <div class="ls-section-body">
      <div class="learning-status-top">
      <section class="exam-plan-panel exam-plan-rows">
        <div class="exam-plan-item" data-plan-item>
          <button type="button" class="exam-plan-head" data-plan-edit aria-expanded="false">
            <span>${L('시험까지', 'Exam in')}</span><strong data-dday-summary>${examDdayHTML(examDdayLabel())}</strong>
          </button>
          <div class="exam-plan-editor" hidden>
            <span hidden data-dday-value>${examDdayLabel()}</span>
            <label><small>${L('시험 날짜', 'Exam date')}</small>${datePickerHTML('data-plan-setting="examDate"', plan.examDate)}</label>
            <div class="exam-plan-actions">
              <button type="button" class="btn" data-plan-cancel>${L('취소', 'Cancel')}</button>
              <button type="button" class="btn btn-primary" data-plan-save>${L('저장', 'Save')}</button>
            </div>
          </div>
        </div>
        <div class="exam-plan-item" data-plan-item>
          <button type="button" class="exam-plan-head" data-plan-edit aria-expanded="false">
            <span>${L('시험일 암기율', 'Exam-day retention')}<br>${L('(예상 / 목표)', '(est. / goal)')}</span>
            <strong><b data-exam-retention class="plan-accent">…</b> / <b data-target-retention-summary>${Math.round(plan.desiredRetention * 100)}%</b></strong>
          </button>
          <div class="exam-plan-editor" hidden>
            <label class="exam-plan-field">
              <small>${L('목표 암기율', 'Target retention')} <strong data-retention-label>${Math.round(plan.desiredRetention * 100)}%</strong></small>
              <input type="range" min="${Math.round(MIN_DESIRED_RETENTION * 100)}" max="${Math.round(MAX_DESIRED_RETENTION * 100)}" step="1" data-plan-setting="desiredRetention" value="${Math.round(plan.desiredRetention * 100)}">
            </label>
            ${retentionWarningHTML(Math.round(plan.desiredRetention * 100), 'data-retention-note')}
            <div class="exam-plan-actions">
              <button type="button" class="btn" data-plan-cancel>${L('취소', 'Cancel')}</button>
              <button type="button" class="btn btn-primary" data-plan-save>${L('저장', 'Save')}</button>
            </div>
          </div>
        </div>
        <div class="exam-plan-item" data-plan-item>
          <button type="button" class="exam-plan-head" data-plan-edit aria-expanded="false">
            <span>${L('하루 새 카드', 'New cards per day')}</span>
            <strong><b data-daily-recommended-summary class="plan-accent">${pace.recommended}${L('개', ' cards')}</b></strong>
          </button>
          <div class="exam-plan-editor" hidden>
            <span hidden data-daily-recommended>${pace.recommended}</span>
            <label class="exam-plan-recommended-toggle">
              <input type="checkbox" data-plan-setting="useRecommendedDailyPace"${plan.useRecommendedDailyPace ? ' checked' : ''}>
              <span>${L('권장 진도 자동 적용', 'Use recommended pace')}</span>
            </label>
            <input type="number" min="1" max="1000" step="1" data-plan-setting="dailyNewCards" data-manual-value="${plan.dailyNewCards}" value="${plan.useRecommendedDailyPace ? pace.recommended : plan.dailyNewCards}"${plan.useRecommendedDailyPace ? ' disabled hidden' : ''} aria-label="${L('사용자 지정 일일 신규 카드량', 'Custom daily new-card amount')}">
            <small data-plan-guidance>${plan.useRecommendedDailyPace
              ? L('남은 기간의 80% 안에 1회독하도록 자동 계산해요.', 'Calculated to finish the first pass within 80% of the remaining time.')
              : L('D-day와 무관하게 직접 설정한 양으로 학습해요.', 'Studies the amount you set, independent of D-day.')}</small>
            <div class="exam-plan-actions">
              <button type="button" class="btn" data-plan-cancel>${L('취소', 'Cancel')}</button>
              <button type="button" class="btn btn-primary" data-plan-save>${L('저장', 'Save')}</button>
            </div>
          </div>
        </div>
        <div class="exam-plan-item">
          <div class="exam-plan-head exam-plan-head-static">
            <span>${L('1회독 완료 예정', 'First pass done by')}</span><strong data-first-pass-completion-summary>${completionDate}</strong>
          </div>
          <span hidden data-first-pass-completion>${completionDate}</span>
        </div>
        <p class="exam-plan-line" data-plan-status role="status"></p>
      </section>
      <div class="learning-status-actions">
        <button type="button" class="btn learning-status-action" data-fsrs-guide-open aria-haspopup="dialog">${L('FSRS-7이란?', 'What is FSRS-7?')}</button>
      </div>
      <dialog class="fsrs-guide fsrs-guide-dialog" aria-label="${L('FSRS-7이란?', 'What is FSRS-7?')}">
        <button type="button" class="fsrs-guide-close" data-fsrs-guide-close aria-label="${L('닫기', 'Close')}">×</button>
        <h3>${L('FSRS-7이란?', 'What is FSRS-7?')}</h3>
        <p class="fsrs-guide-lead">${L('FSRS-7은 복습 기록으로 카드마다 기억 상태를 추적하고, 망각 곡선으로 "언제 복습하면 좋을지"를 계산하는 간격 반복 알고리즘이에요. 고정된 일정표 대신, 내가 실제로 기억하는 정도에 맞춰 복습 시점이 카드마다 달라져요.', 'FSRS-7 is a spaced-repetition algorithm that tracks each card’s memory state from your review history and uses a forgetting curve to decide when to review next. Instead of a fixed schedule, timing adapts to how well you actually remember each card.')}</p>
        <h4>${L('카드마다 추적하는 값', 'What is tracked per card')}</h4>
        <dl>
          <dt>${L('기억 확률 (Retrievability)', 'Recall probability (Retrievability)')}</dt>
          <dd>${L('지금 이 카드를 떠올릴 수 있을 확률 추정치예요. 복습 직후 100%에 가깝고, 시간이 지날수록 망각 곡선을 따라 떨어져요. 학습 상태 페이지의 카드 농도가 바로 이 값이에요.', 'Estimated chance you can recall the card right now. It is near 100% right after a review and decays along the forgetting curve. The card opacity on the Learning Status page is this value.')}</dd>
          <dt>${L('안정성 (Stability)', 'Stability')}</dt>
          <dd>${L('기억 확률이 100%에서 90%로 떨어지기까지 걸리는 기간(일)이에요. 클수록 천천히 잊어요. 복습에 성공할수록 커지고, 잊어버리면 크게 줄어요. 복습 간격은 이 값에 비례해 늘어나요.', 'Days it takes for recall probability to fall from 100% to 90%. Larger means slower forgetting. It grows after successful reviews and drops sharply after a lapse. Review intervals grow in proportion to it.')}</dd>
          <dt>${L('난이도 (Difficulty)', 'Difficulty')}</dt>
          <dd>${L('1~10. 높을수록 외우기 어려운 카드예요. 난이도가 높으면 복습에 성공해도 안정성이 느리게 늘어요. Again/Hard를 자주 누르면 올라가고, Easy를 누르면 내려가요.', '1–10. Higher means harder to remember, so stability grows more slowly even when you succeed. It rises when you press Again/Hard often and falls with Easy.')}</dd>
          <dt>${L('이중 흔적 (빠른/느린 기억)', 'Dual trace (fast / slow memory)')}</dt>
          <dd>${L('FSRS-7의 특징이에요. 갓 외운 단기 기억(빠르게 사라짐)과 반복으로 굳어진 장기 기억(천천히 사라짐)을 따로 추적해 두 곡선을 합쳐 기억 확률을 계산해요. 그래서 같은 날 여러 번 복습한 효과와 며칠 간격 복습의 효과를 더 정확히 구분해요.', 'A hallmark of FSRS-7. It tracks a fast-fading short-term trace and a slow-fading long-term trace separately and combines the two curves into the recall probability. This lets it tell same-day repetition apart from spaced repetition more accurately.')}</dd>
        </dl>
        <h4>${L('복습 평가와 다음 복습', 'Ratings and the next review')}</h4>
        <dl>
          <dt>${L('Again · Hard · Good · Easy', 'Again · Hard · Good · Easy')}</dt>
          <dd>${L('복습 후 고르는 4단계 평가예요. 평가에 따라 안정성과 난이도가 갱신되고 다음 복습일이 정해져요. 대부분의 카드는 Good이면 충분해요.', 'The four ratings you choose after a review. They update stability and difficulty and set the next due date. Good is right for most cards.')}</dd>
          <dt>${L('목표 기억률 (Desired retention)', 'Target retention (Desired retention)')}</dt>
          <dd>${L('복습 시점에 유지하고 싶은 기억 확률이에요. 높이면 복습 횟수가 늘고 더 확실히 기억하며, 낮추면 복습은 줄지만 더 많이 잊어요. 보통 85~95%를 권장해요.', 'The recall probability you want to keep at review time. Raising it means more reviews and firmer memory; lowering it means fewer reviews but more forgetting. 85–95% is typical.')}</dd>
          <dt>${L('시험 날짜', 'Exam date')}</dt>
          <dd>${L('시험일을 넘기는 복습 간격은 만들지 않도록 최대 간격을 시험일까지로 제한해요.', 'The maximum interval is capped so no review is scheduled past your exam date.')}</dd>
        </dl>
        <p class="fsrs-guide-note">${L('모든 수치는 복습 기록에서 계산한 추정치예요. 개인의 실제 기억과는 차이가 있을 수 있고, 기록이 쌓일수록 정확해져요.', 'All values are estimates computed from your review history. They may differ from your actual memory and get more accurate as history accumulates.')}</p>
        <p class="fsrs-guide-source">${L('출처', 'Source')}: <a href="https://github.com/open-spaced-repetition/ts-fsrs" target="_blank" rel="noopener noreferrer">ts-fsrs</a> · <a href="https://github.com/open-spaced-repetition/free-spaced-repetition-scheduler" target="_blank" rel="noopener noreferrer">Free Spaced Repetition Scheduler</a></p>
      </dialog>
      </div>
      </div>
      </section>
      ${studyLogPreviewHTML()}
      <section class="ls-section${state.collapsedSections.cards ? ' is-collapsed' : ''}" data-ls-section="cards">
      ${sectionToggleHTML('cards', L('카드 이해도', 'CARD UNDERSTANDING'), 'h2').replace('class="study-ready-eyebrow ls-section-title"','class="study-ready-eyebrow learning-status-section-title ls-section-title"')}
      <div class="ls-section-body">
      <div class="learning-status-view-row">
      <div class="learning-status-view-switch" role="group" aria-label="${L('보기 방식', 'View mode')}">
        <button type="button" class="learning-status-view-button${state.learningStatusView === 'simple' ? ' is-active' : ''}" data-learning-status-view="simple" aria-controls="learningStatusOverview" aria-pressed="${state.learningStatusView === 'simple'}">${L('간단히', 'Simple')}</button>
        <button type="button" class="learning-status-view-button${state.learningStatusView === 'detailed' ? ' is-active' : ''}" data-learning-status-view="detailed" aria-controls="learningStatusDetailed" aria-pressed="${state.learningStatusView === 'detailed'}">${L('자세히', 'Detailed')}</button>
      </div>
      <header class="learning-status-section-head">
        <p>${L('카드별 현재 기억 확률이에요. 진할수록 잘 기억하고 있고, 회색은 아직 학습하지 않은 카드예요. 확률은 매시간 갱신되는 추정치예요.', 'Current recall probability per card. Darker means better remembered; gray cards are not studied yet. Estimates update hourly.')}</p>
      </header>
      </div>
      <section id="learningStatusOverview" class="learning-status-overview"${state.learningStatusView === 'simple' ? '' : ' hidden'} style="--retention-max-opacity:${(state.studySettings.retentionStyle.strength / 100).toFixed(2)}" aria-label="${L('카드별 이해도 요약', 'Card understanding overview')}">
        <div class="learning-status-overview-head">
          <div class="learning-status-legend" aria-label="${L('이해도 낮음에서 높음 순서', 'Understanding from low to high')}">
            <span>${L('낮음', 'Low')}</span><i aria-hidden="true"></i><span>${L('높음', 'High')}</span>
          </div>
        </div>
        <div class="learning-status-overview-groups">
          ${filters.map(([id, label, count]) => {
            const group = sourceOrderedRecords.filter(({ category }) => category === id);
            return `<section class="learning-status-overview-group is-${id}" aria-label="${label}">
            <h3>${label} <span>${count}</span></h3>
            <div class="learning-status-overview-grid" role="list" aria-label="${label}">
              ${group.length ? group.map(({ pair, record }) => {
                const tileLabel = `${pair.key} · ${id === 'new' ? label : L('계산 중', 'Calculating')}`;
                return `<span class="learning-status-tile" role="listitem" data-learning-status-tile="${esc(record.id)}" data-learning-status-category="${id}" data-tile-key="${esc(pair.key)}" aria-label="${esc(tileLabel)}"></span>`;
              }).join('') : `<p class="learning-status-empty">${L('해당하는 카드가 없어요.', 'No cards here.')}</p>`}
            </div>
          </section>`;
          }).join('')}
        </div>
      </section>
      <section id="learningStatusDetailed" class="learning-status-detailed"${state.learningStatusView === 'detailed' ? '' : ' hidden'}>
      <section class="learning-status-toolbar" aria-label="${L('카드 검색 및 필터', 'Search and filter cards')}">
        <label class="learning-status-search">
          <span>${L('항목 검색', 'Search items')}</span>
          <input type="search" data-learning-status-search value="${esc(state.learningStatusQuery)}" placeholder="${L('key, value, hint 검색', 'Search key, value, or hint')}" autocomplete="off">
        </label>
        <div class="learning-status-filters" role="group" aria-label="${L('학습 상태 필터', 'Learning status filter')}">
          ${filters.map(([id, label, count]) => `<button type="button" class="learning-status-filter${state.learningStatusFilters.includes(id) ? ' is-active' : ''}" data-learning-status-filter="${id}" aria-pressed="${state.learningStatusFilters.includes(id)}">${label}<span>${count}</span></button>`).join('')}
        </div>
        <button type="button" class="btn learning-status-action retention-style-open" data-retention-style-open aria-haspopup="dialog">${L('투명도 설정', 'Opacity')}</button>
      </section>
      ${retentionStyleDialogHTML()}
      <section class="learning-status-list" aria-label="${L('카드별 상태 목록', 'Card status list')}">
        ${sortedRecords.length
          ? sortedRecords.map(({ pair, record, category }) => `
            <button type="button" class="learning-status-card" data-learning-status-card="${category}" data-learning-status-open="${esc(record.id)}" aria-haspopup="dialog" data-search="${esc(`${pair.key} ${pair.value} ${pair.hint}`.toLocaleLowerCase())}" title="${L('눌러서 카드 정보 보기', 'Select to view card details')}">
              <span class="learning-status-card-key">${richText(swappedPair(pair).key)}</span>
              <span class="learning-status-card-recall" data-fsrs-retrievability="${esc(record.id)}">${category === 'new' ? L('학습 전', 'Not studied') : '…'}</span>
            </button>
            <template data-learning-status-detail="${esc(record.id)}">
              <div class="learning-status-dialog-card">
                <button type="button" class="learning-status-dialog-close" data-learning-status-close aria-label="${L('닫기', 'Close')}">×</button>
                <span class="learning-status-badge is-${category}">${statusLabels[category]}</span>
                <div class="learning-status-dialog-content">
                  <div class="learning-status-back-wrap">
                    <div class="learning-status-back">
                      <div class="learning-status-back-value" style="font-size:${ankiFontSize(state.studySettings.fontSizes.back)}">${richText(swappedPair(pair).value)}</div>
                      <div class="learning-status-back-hint" style="font-size:${ankiFontSize(state.studySettings.fontSizes.back)}">${pair.hint ? richText(pair.hint) : `<span class="learning-status-nohint">${L('등록된 hint가 없어요.', 'No hint available.')}</span>`}</div>
                    </div>
                  </div>
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
      </section>
      </div>
      </section>
    </main>`;
}

function a4ViewHTML() {
  const itemCounts = a4ItemCounts();
  const a4HasImage = state.pairs.some((pair) => /\[img:/.test(`${pair.key}\n${pair.value}\n${pair.hint ?? ''}`));
  return `
    ${headerHTML()}
    <main class="shell a4-view-screen${settingsCollapsed ? ' settings-collapsed' : ''}" style="--study-panel-width:${studyPanelWidth}px">
      <button type="button" id="reopenSettingsBtn" class="a4-settings-reopen" aria-label="${L('설정 펼치기','Expand settings')}" title="${L('설정 펼치기','Expand settings')}"${settingsCollapsed ? '' : ' hidden'}>
        <svg aria-hidden="true" viewBox="0 0 24 24" fill="none">
          <path d="M19.14 12.94c.04-.3.06-.61.06-.94s-.02-.64-.07-.94l2.03-1.58a.5.5 0 0 0 .12-.64l-1.92-3.32a.5.5 0 0 0-.61-.22l-2.39.96a7.1 7.1 0 0 0-1.63-.94l-.36-2.54a.49.49 0 0 0-.49-.42h-3.84a.49.49 0 0 0-.49.42l-.36 2.54c-.59.23-1.14.55-1.63.94l-2.39-.96a.49.49 0 0 0-.61.22L2.63 8.84a.5.5 0 0 0 .12.64l2.03 1.58c-.05.3-.08.62-.08.94s.03.64.08.94l-2.03 1.58a.5.5 0 0 0-.12.64l1.92 3.32c.12.21.37.3.61.22l2.39-.96c.49.39 1.04.71 1.63.94l.36 2.54c.04.24.24.42.49.42h3.84c.25 0 .45-.18.49-.42l.36-2.54c.59-.23 1.14-.55 1.63-.94l2.39.96c.23.09.49-.01.61-.22l1.92-3.32a.5.5 0 0 0-.12-.64l-2.01-1.56ZM12 15.5a3.5 3.5 0 1 1 0-7 3.5 3.5 0 0 1 0 7Z" fill="currentColor"/>
        </svg>
      </button>
      <section class="panel a4-settings" aria-label="${L('설정','Settings')}">
        <button type="button" id="collapseSettingsBtn" class="a4-settings-collapse" aria-label="${L('설정 접기','Collapse settings')}" title="${L('설정 접기','Collapse settings')}">
          <svg aria-hidden="true" viewBox="0 0 24 24" fill="none"><path d="m15 5-7 7 7 7" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
        </button>
        <div class="a4-settings-content">
          ${state.notice ? `<p class="notice a4-notice" role="status">${esc(state.notice)}</p>` : ''}
          <p id="a4StorageError" class="error" role="alert"${state.error ? '' : ' hidden'}>${esc(state.error)}</p>
          <section class="a4-settings-group" aria-labelledby="a4SearchTitle">
            <h2 id="a4SearchTitle">${L('검색', 'Search')}</h2>
            <input id="studySearch" class="search" type="search" aria-label="${L('key 또는 value 검색','Search by key or value')}" placeholder="${L('key 또는 value 검색','Search by key or value')}" value="${esc(state.a4Query)}" />
          </section>
          <section class="a4-settings-group" aria-labelledby="a4ViewSettingsTitle">
            <h2 id="a4ViewSettingsTitle">${L('표시 설정', 'Display')}</h2>
            <button type="button" id="toggleAllValuesBtn" class="btn a4-settings-control a4-values-toggle" aria-pressed="${state.showAllA4Values}" data-study-hint="${L('전체 value 보기','Show all values')}">${state.showAllA4Values ? L('전체 value 가리기','Hide all values') : L('전체 value 보기','Show all values')}</button>
            <button type="button" id="toggleAllHintsBtn" class="btn a4-settings-control" aria-pressed="false" data-study-hint="${L('전체 hint 보기','Show all hints')}">${L('전체 hint 보기','Show all hints')}</button>
            <button type="button" id="a4SwapBtn" class="btn a4-settings-control" aria-pressed="${state.studySettings.swapKeyValue}" data-study-hint="${L('key와 value의 앞뒤를 바꿔서 표시합니다.','Swap which side is shown as key and value.')}">${L('key-value 반전','Swap key-value')}</button>
            <label class="study-font-setting a4-image-gauge${a4HasImage ? '' : ' is-disabled'}"${a4HasImage ? '' : ` title="${L('이미지가 없어요', 'No images in this deck')}"`}>
              <span>${L('이미지', 'Image')}</span>
              <input type="range"${a4HasImage ? '' : ' disabled'} min="${A4_IMAGE_HEIGHT_RANGE.min}" max="${A4_IMAGE_HEIGHT_RANGE.max}" step="1" id="a4ImageHeight" value="${state.studySettings.a4ImageHeight}" aria-label="${L('이미지 크기', 'Image size')}">
              <output>${state.studySettings.a4ImageHeight}</output>
            </label>
          </section>
          <section class="a4-settings-group" aria-labelledby="a4OrderSettingsTitle">
            <h2 id="a4OrderSettingsTitle">${L('학습 순서', 'Study order')}</h2>
            <label class="a4-sort-check"><input type="checkbox" id="a4SortPriority"${state.a4Sort === 'priority' ? ' checked' : ''}> ${L('학습 우선순위로 정렬', 'Sort by learning priority')}</label>
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

function a4PrintSheetHTML(rows, pageNumber, pageCount, isLastPage = false) {
  return `<div class="sheet a4-sheet a4-print-sheet${isLastPage ? ' a4-last-page' : ''}">
      <div class="sheet-head">
        <div>
          <div class="sheet-title">${esc(a4SheetTitle())}</div>
          <div class="sheet-meta">${esc(a4SheetInfo())}</div>
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
  const overviewTiles = new Map([...appEl.querySelectorAll('[data-learning-status-tile]')]
    .map((element) => [element.dataset.learningStatusTile, element]));
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
        applyRetentionOpacity(card, probability);
        card.dataset.recallLevel = percent >= 85 ? 'high' : percent >= 60 ? 'mid' : 'low';
      }
      const tile = overviewTiles.get(record.id);
      if (tile) {
        applyRetentionOpacity(tile, probability);
        tile.dataset.recallPercent = String(percent);
        const label = L(
          `${tile.dataset.tileKey} · 현재 기억 확률 ${percent}%`,
          `${tile.dataset.tileKey} · current recall probability ${percent}%`
        );
        tile.setAttribute('aria-label', label);
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
      return `${sectionLabel ? `<div class="a4-section-divider" style="width:${SHEET_W_PX * scale}px"><span>${esc(sectionLabel)}</span></div>` : ''}<div class="frame" style="width:${SHEET_W_PX * scale}px;height:${SHEET_H_PX * scale}px">
        <div class="sheet-scaler" style="transform:scale(${scale});transform-origin:top left">
          ${a4SheetHTML(rows, i + 1, totalPages, isLastPage)}
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
  if (!state.page === 'study' && studyTimerInterval) {
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
            : state.page === 'learning-status'
                ? learningStatusHTML()
                : state.page === 'project-setup'
                  ? projectSetupHTML()
            : a4ViewHTML()}${state.languageChoicePending ? languageChoiceHTML() : ''}${state.anki ? ankiImportHTML() : ''}`;
  appEl.innerHTML = pageHTML;
  renderProcessingIndicator();
  applyImageSizes();
  if (state.page === 'study') fitStudyCardText();
  appEl.classList.toggle('home-page', state.page === 'home');
  appEl.classList.toggle('workspace-page', state.page === 'workspace');
  appEl.classList.toggle('a4-view-page', state.page === 'a4');
  appEl.classList.toggle('study-page', state.page === 'study');
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
          <a class="support-sponsors-link" href="https://github.com/sponsors/sunghokim128" target="_blank" rel="noopener noreferrer">${L('GitHub Sponsors에서 후원하기 ↗', 'Sponsor on GitHub ↗')}</a>
          <p id="supportDescription">${L('후원하신 금액은 전액 개발자의 식비와<br>Copilot 구독비로 탕진됩니다.', 'Every donation goes directly toward keeping the developer fed and his Copilot running.')}</p>
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
    .map(({ rows }, i) => {
      const isLastPage = i === list.length - 1;
      return a4PrintSheetHTML(rows, i + 1, totalPages, isLastPage);
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

function buildStudyQueue() {
  const pairIds = new Set(state.pairs.map((pair) => pair.id));
  const now = Date.now();
  const activeRecords = (state.fsrsData?.cards || []).filter((record) =>
    record.archivedAt === null && pairIds.has(record.id)
  );
  const progress = currentDailyNewProgress();
  const dueReviewIds = activeRecords
    .filter((record) => record.card.state !== FsrsState.New && Date.parse(record.card.dueAt) <= now)
    .map((record) => record.id);
  const newCardIds = shuffle(progress.newCards
    .slice(0, progress.availableNewCards)
    .map((record) => record.id));
  return { dueReviewIds, newCardIds, progress };
}

async function startStudy() {
  const plan = projectPlan();
  state.studyError = '';
  if (state.project && plan.examDate && !state.project.plan?.firstPassDeadline) {
    try {
      await saveProjectPlan({ ...plan, firstPassDeadline: plan.firstPassDeadline });
    } catch (error) {
      state.studyError = L(
        `일일 진도 일정을 저장하지 못했어요: ${error instanceof Error ? error.message : String(error)}`,
        `Could not save the daily study schedule: ${error instanceof Error ? error.message : String(error)}`
      );
      console.error('Could not persist the first-pass deadline.', error);
    }
  }
  const queue = buildStudyQueue();
  const cards = shuffle(queue.dueReviewIds).concat(queue.newCardIds);
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

async function addMoreNewCards() {
  const progress = currentDailyNewProgress();
  if (!progress.canOfferExtra || !state.fsrsData || !state.project || state.studySaving) return;
  state.studySaving = true;
  render();
  const today = localStudyDate();
  const nextData = structuredClone(state.fsrsData);
  const extraNewCards = nextData.dailyProgress?.date === today
    ? nextData.dailyProgress.extraNewCards
    : 0;
  nextData.dailyProgress = { date: today, extraNewCards: extraNewCards + 10 };
  try {
    await saveFsrsData(state.project, nextData);
    state.fsrsData = nextData;
    state.studySaving = false;
    await startStudy();
  } catch (error) {
    state.studySaving = false;
    state.studyError = L(
      `추가 학습량을 저장하지 못했어요: ${error instanceof Error ? error.message : String(error)}`,
      `Could not save the extra study allowance: ${error instanceof Error ? error.message : String(error)}`
    );
    console.error('Could not save the additional daily new-card allowance.', error);
    render();
  }
}

function beginStudySession() {
  const session = state.study;
  if (!session || session.started || state.studySaving) return;
  session.started = true;
  session.startedAt = Date.now();
  enterStudyCard(session);
  render();
}

function advanceStudy(rating, before) {
  const session = state.study;
  if (!session || state.studySaving) return;
  const pair = studyCurrentPair();
  if (!pair) return;
  session.history.push({ id: pair.id, rating, before });
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
  const before = state.fsrsData;
  state.studySaving = true;
  state.studyError = '';
  render();
  try {
    const result = await reviewFsrsCard(state.project, state.fsrsData, pair.id, rating, new Date(), projectPlan().desiredRetention, projectPlan().examDate);
    state.pendingReview = { data: result.data, rating, before };
    await saveFsrsData(state.project, result.data);
    state.fsrsData = result.data;
    state.pendingReview = null;
    state.studySaving = false;
    advanceStudy(rating, before);
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
    const { rating, before } = state.pendingReview;
    state.fsrsData = state.pendingReview.data;
    state.pendingReview = null;
    state.studySaving = false;
    advanceStudy(rating, before);
  } catch (error) {
    state.studySaving = false;
    state.studyError = L(
      `복습 결과를 아직 저장하지 못했어요. 폴더 권한과 여유 공간을 확인한 뒤 다시 시도해 주세요. ${error instanceof Error ? error.message : String(error)}`,
      `The review is still unsaved. Check folder access and free space, then retry. ${error instanceof Error ? error.message : String(error)}`
    );
    render();
  }
}

async function undoStudyRating() {
  const session = state.study;
  if (!session?.started || !session.history.length || state.studySaving || state.pendingReview) return;
  const last = session.history.at(-1);
  state.studySaving = true;
  state.studyError = '';
  render();
  try {
    await saveFsrsData(state.project, last.before);
    state.fsrsData = last.before;
    session.history.pop();
    session.index = session.cards.indexOf(last.id);
    session.enterEffect = 'back';
    state.studySaving = false;
    enterStudyCard(session);
    render();
  } catch (error) {
    state.studySaving = false;
    state.studyError = L(
      `선택을 되돌리지 못했어요. 다시 시도해 주세요. ${error instanceof Error ? error.message : String(error)}`,
      `Could not undo the rating. Please try again. ${error instanceof Error ? error.message : String(error)}`
    );
    render();
  }
}

async function editCurrentStudyCard() {
  const session = state.study;
  const pair = session?.started ? studyCurrentPair() : null;
  if (!pair || !state.project || state.studySaving) return;
  const projectId = state.project.id;
  await finishStudy(false, 'workspace');
  if (state.page === 'study') return;
  await editProject(projectId, pair.id);
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
    const savedToLibrary = await persistLibrary();
    if (savedToLibrary) state.libraryNeedsMigration = false;
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

async function finishStudy(askForConfirmation = false, returnPage = 'study', { timeLimitReached = false } = {}) {
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
      const duration = timeLimitReached
        ? Math.min(elapsed, session.settings.roundTimeLimitMinutes * 60_000)
        : elapsed;
      const itemCount = [...session.seenIds].filter((id) => state.pairs.some((pair) => pair.id === id)).length;
      const recordFile = await readStudyRecordFile(state.project.name);
      session.pendingSave = {
        records: appendStudySession(recordFile, {
          date,
          itemCount,
          durationSeconds: duration / 1000,
        }),
      };
    }
    await writeStudyRecordFile(state.project.name, session.pendingSave.records);
    state.studyRecords = null;
    state.logDeleteError = '';
    if (!(await save())) throw new Error(state.error || L('프로젝트 정보를 저장하지 못했어요.', 'Could not save project information.'));
    if (window.speechSynthesis) window.speechSynthesis.cancel();
    state.study = null;
    state.studySaving = false;
    state.studyError = '';
    if (returnPage === 'study') {
      await startStudy();
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
  await persistLibrary();
  // Keep a copy in the project folder so it survives changing the study directory.
  if (state.fsrsData) {
    state.fsrsData.plan = { ...plan };
    try {
      await saveFsrsData(project, state.fsrsData);
    } catch (error) {
      console.error('Could not save the plan to the FSRS file.', error);
    }
  }
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
        dailyNewCards: state.project?.plan?.dailyNewCards ?? DEFAULT_DAILY_NEW_CARDS,
        useRecommendedDailyPace: projectPlan().useRecommendedDailyPace,
        firstPassDeadline: dateInput.value === projectPlan().examDate && projectPlan().firstPassDeadline
          ? projectPlan().firstPassDeadline
          : firstPassDeadline(dateInput.value),
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
  const status = appEl.querySelector('[data-plan-status]');
  bindStudyLogPreview();
  loadStudyLogPreview();
  const slider = appEl.querySelector('[data-plan-setting="desiredRetention"]');
  const dateInput = appEl.querySelector('[data-plan-setting="examDate"]');
  const dailyInput = appEl.querySelector('[data-plan-setting="dailyNewCards"]');
  const recommendedToggle = appEl.querySelector('[data-plan-setting="useRecommendedDailyPace"]');
  const dailyRecommended = appEl.querySelector('[data-daily-recommended]');
  const completionOutput = appEl.querySelector('[data-first-pass-completion]');
  slider?.addEventListener('input', () => {
    const retention = Number(slider.value) / 100;
    appEl.querySelector('[data-retention-label]').textContent = `${Math.round(retention * 100)}%`;
    appEl.querySelector('[data-target-retention-summary]').textContent = `${Math.round(retention * 100)}%`;
    appEl.querySelector('[data-retention-note]').hidden = Math.round(retention * 100) <= RETENTION_WARNING_THRESHOLD;
  });
  const updatePacePreview = () => {
    const plan = projectPlan();
    const dday = examDdayLabel(dateInput.value);
    appEl.querySelector('[data-dday-value]').textContent = dday;
    appEl.querySelector('[data-dday-summary]').innerHTML = examDdayHTML(dday);
    const remainingCards = state.fsrsData?.cards.filter((record) =>
      record.archivedAt === null &&
      state.pairs.some((pair) => pair.id === record.id) &&
      record.card.state === FsrsState.New
    ).length || 0;
    const preview = selectedDailyPace({
      ...plan,
      examDate: dateInput.value,
      firstPassDeadline: dateInput.value === plan.examDate
        ? plan.firstPassDeadline
        : firstPassDeadline(dateInput.value),
      dailyNewCards: recommendedToggle.checked
        ? plan.dailyNewCards
        : Number(dailyInput.value),
      useRecommendedDailyPace: recommendedToggle.checked,
    }, remainingCards);
    if (recommendedToggle.checked) dailyInput.value = String(preview.recommended);
    const formattedCompletion = preview.completionDate
      ? new Intl.DateTimeFormat(language === 'ko' ? 'ko-KR' : 'en', { dateStyle: 'medium' })
        .format(parseLocalDate(preview.completionDate))
      : '—';
    dailyRecommended.textContent = String(preview.recommended);
    completionOutput.textContent = formattedCompletion;
    appEl.querySelector('[data-daily-recommended-summary]').textContent =
      `${preview.recommended}${L('개', ' cards')}`;
    appEl.querySelector('[data-first-pass-completion-summary]').textContent = formattedCompletion;
    appEl.querySelector('[data-plan-guidance]').textContent = recommendedToggle.checked
      ? L('남은 기간의 80% 안에 1회독하도록 일일 권장 진도를 자동 계산해요. 복습 카드는 일일 진도와 별도로 학습할 수 있어요.', 'Recommended pace is calculated to finish the first pass within 80% of the remaining exam period. Reviews are separate from the daily new-card pace.')
      : L('권장 진도를 끄면 D-day와 무관하게 직접 설정한 양으로 학습해요. 복습 카드는 일일 진도와 별도로 학습할 수 있어요.', 'With recommendations off, your chosen daily amount is independent of D-day. Reviews are separate from the daily new-card pace.');
    return preview;
  };
  recommendedToggle?.addEventListener('change', () => {
    if (recommendedToggle.checked) {
      dailyInput.dataset.manualValue = String(Number(dailyInput.value) || projectPlan().dailyNewCards);
      dailyInput.disabled = true;
      dailyInput.hidden = true;
      dailyInput.value = String(updatePacePreview().recommended);
    } else {
      dailyInput.disabled = false;
      dailyInput.hidden = false;
      dailyInput.value = dailyInput.dataset.manualValue || String(projectPlan().dailyNewCards);
    }
    updatePacePreview();
  });
  dailyInput?.addEventListener('input', () => {
    if (!recommendedToggle.checked) dailyInput.dataset.manualValue = dailyInput.value;
    updatePacePreview();
  });
  dateInput?.addEventListener('change', updatePacePreview);
  appEl.querySelectorAll('[data-plan-edit]').forEach((head) => {
    head.addEventListener('click', () => {
      const item = head.closest('[data-plan-item]');
      const open = !item.classList.contains('is-editing');
      item.classList.toggle('is-editing', open);
      head.setAttribute('aria-expanded', String(open));
      item.querySelector('.exam-plan-editor').hidden = !open;
    });
  });
  appEl.querySelectorAll('[data-plan-cancel]').forEach((button) => button.addEventListener('click', () => {
    state.learningStatusPlanMessage = '';
    render();
  }));
  appEl.querySelectorAll('[data-plan-save]').forEach((saveButton) => saveButton.addEventListener('click', async () => {
    appEl.querySelectorAll('[data-plan-save]').forEach((button) => { button.disabled = true; });
    const plan = {
      examDate: dateInput.value,
      desiredRetention: normalizeDesiredRetention(Number(slider.value) / 100),
      dailyNewCards: recommendedToggle.checked
        ? Math.min(1000, Math.max(1, Math.floor(Number(dailyInput.dataset.manualValue) || projectPlan().dailyNewCards)))
        : Math.min(1000, Math.max(1, Math.floor(Number(dailyInput.value) || DEFAULT_DAILY_NEW_CARDS))),
      useRecommendedDailyPace: recommendedToggle.checked,
      firstPassDeadline: dateInput.value === projectPlan().examDate
        ? projectPlan().firstPassDeadline
        : firstPassDeadline(dateInput.value),
    };
    try {
      const result = await rescheduleFsrsData(state.project, state.fsrsData, plan.desiredRetention, plan.examDate);
      await saveProjectPlan(plan);
      result.data.plan = { ...plan };
      await saveFsrsData(state.project, result.data);
      state.fsrsData = result.data;
      state.learningStatusPlanMessage = L(`저장했어요. 복습 중인 카드 ${result.changed}개의 다음 복습일을 새 기준으로 변경했어요.`, `Saved. Updated the next review date of ${result.changed} card(s) in review to the new settings.`);
    } catch (error) {
      state.learningStatusPlanMessage = L('설정을 적용하지 못했어요: ', 'Could not apply the settings: ') + (error instanceof Error ? error.message : String(error));
      console.error('Could not reschedule FSRS cards.', error);
    }
    render();
  }));
  if (status && state.learningStatusPlanMessage) status.textContent = state.learningStatusPlanMessage;

  const search = appEl.querySelector('[data-learning-status-search]');
  const empty = appEl.querySelector('[data-learning-status-empty]');
  const cardList = appEl.querySelector('.learning-status-list');
  const detailDialog = appEl.querySelector('.learning-status-dialog');
  appEl.querySelectorAll('[data-learning-status-view]').forEach((button) => {
    button.addEventListener('click', () => {
      const view = button.dataset.learningStatusView;
      if (!['simple', 'detailed'].includes(view)) return;
      state.learningStatusView = view;
      appEl.querySelector('#learningStatusOverview').hidden = view !== 'simple';
      appEl.querySelector('#learningStatusDetailed').hidden = view !== 'detailed';
      appEl.querySelectorAll('[data-learning-status-view]').forEach((item) => {
        const selected = item.dataset.learningStatusView === view;
        item.classList.toggle('is-active', selected);
        item.setAttribute('aria-pressed', String(selected));
      });
      persistUserPreferences();
    });
  });
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
      applyRetentionOpacity(card, probability);
      const tile = [...appEl.querySelectorAll('[data-learning-status-tile]')]
        .find((item) => item.dataset.learningStatusTile === record.id);
      if (tile) {
        applyRetentionOpacity(tile, probability);
        tile.dataset.recallPercent = String(percent);
        const label = L(
          `${tile.dataset.tileKey} · 현재 기억 확률 ${percent}%`,
          `${tile.dataset.tileKey} · current recall probability ${percent}%`
        );
        tile.setAttribute('aria-label', label);
      }
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
      const matchesFilter = state.learningStatusFilters.includes(card.dataset.learningStatusCard);
      const matchesQuery = !query || card.dataset.search.includes(query);
      card.hidden = !matchesFilter || !matchesQuery;
      if (!card.hidden) visibleCount += 1;
    });
    if (empty) empty.hidden = visibleCount > 0;
  };
  const fsrsGuide = appEl.querySelector('.fsrs-guide-dialog');
  appEl.querySelector('[data-fsrs-guide-open]')?.addEventListener('click', () => fsrsGuide?.showModal());
  attachRetentionStyleDialog();
  fsrsGuide?.addEventListener('click', (event) => {
    if (event.target === fsrsGuide || event.target.closest('[data-fsrs-guide-close]')) fsrsGuide.close();
  });
  search?.addEventListener('input', () => {
    state.learningStatusQuery = search.value;
    applyFilters();
  });
  appEl.querySelectorAll('[data-learning-status-filter]').forEach((button) => {
    button.addEventListener('click', () => {
      const id = button.dataset.learningStatusFilter;
      state.learningStatusFilters = state.learningStatusFilters.includes(id)
        ? state.learningStatusFilters.filter((item) => item !== id)
        : [...state.learningStatusFilters, id];
      render();
    });
  });
  applyFilters();
}

function hexToHsv(hex) {
  const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b);
  const d = max - Math.min(r, g, b);
  let h = 0;
  if (d) {
    if (max === r) h = ((g - b) / d) % 6;
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
  }
  return { h: (h * 60 + 360) % 360, s: max ? d / max : 0, v: max };
}

function hsvToHex({ h, s, v }) {
  const f = (n) => {
    const k = (n + h / 60) % 6;
    return v - v * s * Math.max(0, Math.min(k, 4 - k, 1));
  };
  return `#${[f(5), f(3), f(1)].map((x) => Math.round(x * 255).toString(16).padStart(2, '0')).join('')}`;
}

function attachRetentionStyleDialog() {
  const dialog = appEl.querySelector('.retention-dialog');
  if (!dialog) return;
  const swatches = [...dialog.querySelectorAll('[data-retention-swatch]')];
  const svBox = dialog.querySelector('[data-retention-sv]');
  const svHandle = dialog.querySelector('[data-retention-sv-handle]');
  const hueInput = dialog.querySelector('[data-retention-hue]');
  const hexInput = dialog.querySelector('[data-retention-hex]');
  let hsv = hexToHsv(state.studySettings.retentionStyle.color);
  let currentColor = state.studySettings.retentionStyle.color;
  const strengthInput = dialog.querySelector('[data-retention-strength]');
  const halfBar = dialog.querySelector('[data-retention-half-bar]');
  const halfArrow = dialog.querySelector('[data-retention-half-arrow]');
  let half = state.studySettings.retentionStyle.half;
  const apply = () => {
    applyImageSizes();
    appEl.querySelector('.learning-status-overview')?.style.setProperty(
      '--retention-max-opacity',
      (Number(strengthInput.value) / 100).toFixed(2)
    );
    swatches.forEach((swatch) => applyRetentionOpacity(swatch, Number(swatch.dataset.retentionSwatch) / 100));
    dialog.querySelector('[data-retention-strength-out]').textContent = `${strengthInput.value}%`;
    const rgb = getComputedStyle(document.documentElement).getPropertyValue('--retention-rgb').trim();
    const stops = Array.from({ length: 11 }, (_, i) => `rgba(${rgb}, ${(retentionRelative(1 - i / 10, half) * 0.9).toFixed(3)}) ${i * 10}%`);
    halfBar.style.background = `linear-gradient(to right, ${stops.join(', ')})`;
    halfArrow.style.left = `${100 - half}%`;
    dialog.querySelector('[data-retention-half-note]').textContent = L(
      `화살표가 있는 ${half}% 지점에서 투명도가 절반이 되고, 그만큼 내려갈 때마다 다시 절반이 돼요. 왼쪽으로 옮기면 흐려지는 속도가 느려져요.`,
      `Opacity is halved at the ${half}% mark, and halves again for every step down by the same distance. Drag the arrow to the right to fade faster.`
    );
    appEl.querySelectorAll('.learning-status-card:not([data-learning-status-card="new"]), .learning-status-tile:not([data-learning-status-category="new"])').forEach((card) => {
      const percent = Number(card.dataset.recallPercent ?? card.querySelector('[data-fsrs-retrievability]')?.textContent?.replace('%', ''));
      if (Number.isFinite(percent)) applyRetentionOpacity(card, percent / 100);
    });
  };
  const paintPicker = () => {
    svBox.style.backgroundColor = `hsl(${hsv.h}, 100%, 50%)`;
    svHandle.style.left = `${hsv.s * 100}%`;
    svHandle.style.top = `${(1 - hsv.v) * 100}%`;
    swatchBtn.style.background = currentColor;
    hueInput.value = String(Math.round(hsv.h));
  };
  const read = () => {
    state.studySettings.retentionStyle = normalizeRetentionStyle({ color: currentColor, strength: strengthInput.value, half });
    apply();
  };
  const pickColor = (next, { syncHex = true } = {}) => {
    hsv = next;
    currentColor = hsvToHex(hsv);
    if (syncHex) hexInput.value = currentColor;
    hexInput.classList.remove('is-invalid');
    paintPicker();
    read();
  };
  const swatchBtn = dialog.querySelector('[data-retention-swatch-btn]');
  const pop = dialog.querySelector('[data-retention-pop]');
  swatchBtn.addEventListener('click', () => {
    pop.hidden = !pop.hidden;
    swatchBtn.setAttribute('aria-expanded', String(!pop.hidden));
    if (!pop.hidden) paintPicker();
  });
  dialog.addEventListener('click', (event) => {
    if (!pop.hidden && !event.target.closest('.retention-picker')) {
      pop.hidden = true;
      swatchBtn.setAttribute('aria-expanded', 'false');
    }
  });
  const pickFromSv = (event) => {
    const rect = svBox.getBoundingClientRect();
    const x = Math.min(1, Math.max(0, (event.clientX - rect.left) / rect.width));
    const y = Math.min(1, Math.max(0, (event.clientY - rect.top) / rect.height));
    pickColor({ h: hsv.h, s: x, v: 1 - y });
  };
  svBox.addEventListener('pointerdown', (event) => {
    svBox.setPointerCapture(event.pointerId);
    pickFromSv(event);
    const move = (moveEvent) => pickFromSv(moveEvent);
    const up = () => {
      svBox.removeEventListener('pointermove', move);
      svBox.removeEventListener('pointerup', up);
      save();
    };
    svBox.addEventListener('pointermove', move);
    svBox.addEventListener('pointerup', up);
  });
  hueInput.addEventListener('input', () => pickColor({ ...hsv, h: Number(hueInput.value) }));
  hueInput.addEventListener('change', () => save());
  const save = async () => {
    await saveStudySettings();
    await persistLibrary();
  };
  strengthInput.addEventListener('input', read);
  strengthInput.addEventListener('change', save);
  const pickHalf = (event) => {
    const rect = halfBar.getBoundingClientRect();
    const x = Math.min(0.95, Math.max(0.05, (event.clientX - rect.left) / rect.width));
    half = Math.round(100 - x * 100);
    read();
  };
  halfBar.addEventListener('pointerdown', (event) => {
    halfBar.setPointerCapture(event.pointerId);
    pickHalf(event);
    const up = () => {
      halfBar.removeEventListener('pointermove', pickHalf);
      halfBar.removeEventListener('pointerup', up);
      save();
    };
    halfBar.addEventListener('pointermove', pickHalf);
    halfBar.addEventListener('pointerup', up);
  });
  hexInput.addEventListener('input', () => {
    const value = hexInput.value.trim().replace(/^#?/, '#');
    const valid = /^#[0-9a-f]{6}$/i.test(value);
    hexInput.classList.toggle('is-invalid', !valid);
    if (valid) pickColor(hexToHsv(value.toLowerCase()), { syncHex: false });
  });
  hexInput.addEventListener('change', () => {
    hexInput.value = currentColor;
    hexInput.classList.remove('is-invalid');
    save();
  });
  dialog.querySelector('[data-retention-reset]').addEventListener('click', () => {
    const defaults = DEFAULT_STUDY_SETTINGS.retentionStyle;
    strengthInput.value = defaults.strength;
    half = defaults.half;
    pickColor(hexToHsv(defaults.color));
    save();
  });
  dialog.addEventListener('click', (event) => { if (event.target === dialog) dialog.close(); });
  appEl.querySelector('[data-retention-style-open]')?.addEventListener('click', () => {
    paintPicker();
    apply();
    dialog.showModal();
  });
}

async function routeToProjectDestination(destination) {
  if (destination === 'a4') {
    state.page = 'a4';
    render();
  } else if (destination === 'status') {
    state.logPreviewError = '';
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
  // Leaving an in-progress study round ends it (after confirmation).
  if (destination !== 'study' && state.study?.started) {
    const interim = destination === 'a4' ? 'a4' : 'learning-status';
    await finishStudy(true, interim);
    if (state.study) return;
    if (destination === 'a4' || destination === 'status') return;
  }
  if (destination === 'study') {
    if (!state.study) await startStudy();
    else {
      state.page = 'study';
      render();
    }
  } else if (destination === 'a4') {
    state.page = 'a4';
    render();
  } else if (destination === 'status') {
    state.page = 'learning-status';
    render();
  }
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
  } else if (action === 'undo') {
    await undoStudyRating();
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
    const session = state.study;
    if (!timer || !session?.started) return;
    const elapsed = Date.now() - session.startedAt;
    timer.textContent = formatStudyTime(elapsed);
    if (
      elapsed >= session.settings.roundTimeLimitMinutes * 60_000 &&
      !state.pendingReview &&
      !state.studySaving
    ) {
      void finishStudy(false, 'study', { timeLimitReached: true });
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
      else if (action === 'restart') await startStudy();
      else if (action === 'more-new') await addMoreNewCards();
      else if (action === 'fullscreen') await toggleStudyFullscreen();
      else if (action === 'retry-review') await retryPendingReviewSave();
      else if (action === 'undo') await undoStudyRating();
      else if (action === 'edit-card') await editCurrentStudyCard();
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
      const value = input.type === 'checkbox' ? input.checked : input.value;
      state.studySettings = normalizeStudySettings({ ...state.studySettings, [setting]: value });
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

  appEl.querySelectorAll('[data-study-font], [data-study-image]').forEach((input) => {
    const kind = input.hasAttribute('data-study-font') ? 'font' : 'image';
    const field = input.dataset.studyFont || input.dataset.studyImage;
    const update = () => {
      const value = setCardSize(field, kind, input.value);
      input.nextElementSibling.textContent = kind === 'font' ? `${value}px` : String(value);
      fitStudyCardText();
    };
    input.addEventListener('input', update);
    input.addEventListener('change', async () => {
      update();
      if (!(await saveStudySettings())) render();
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
          persistLibrary().catch((error) => {
            setEditStatus(`${L('열 너비를 라이브러리에 저장하지 못했어요.', 'Could not save column widths to the library.')}${error instanceof Error ? ` ${error.message}` : ''}`, true);
          });
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
      state.editHighlightRow = null;
      state.editRows.push(['', '', '']);
      state.editDirty = true;
      state.editFocus = { row: state.editRows.length - 1, col: 0 };
      render();
    });
    appEl.querySelectorAll('[data-edit-delete-row]').forEach((button) => {
      button.addEventListener('click', () => {
        state.editHighlightRow = null;
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

  const folderPrompt = document.getElementById('folderPrompt');
  if (folderPrompt) {
    folderPrompt.addEventListener('cancel', () => { state.folderPromptDismissed = true; });
    folderPrompt.querySelector('[data-folder-prompt-later]').addEventListener('click', () => {
      state.folderPromptDismissed = true;
      folderPrompt.close();
    });
    folderPrompt.querySelector('[data-folder-prompt-choose]').addEventListener('click', () => {
      state.folderPromptDismissed = true;
      folderPrompt.close();
      connectWorkbookDirectory();
    });
    folderPrompt.showModal();
  }
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

  const a4SortCheck = document.getElementById('a4SortPriority');
  a4SortCheck?.addEventListener('change', () => {
    state.a4Sort = a4SortCheck.checked ? 'priority' : 'order';
    const priorityStatus = document.querySelector('.a4-priority-status');
    if (priorityStatus) {
      priorityStatus.textContent = a4PriorityStatusText();
      priorityStatus.hidden = state.a4Sort !== 'priority' || !priorityStatus.textContent;
    }
    renderA4View();
  });

  const a4ImageGauge = document.getElementById('a4ImageHeight');
  a4ImageGauge?.addEventListener('input', () => {
    state.studySettings.a4ImageHeight = normalizeA4ImageHeight(a4ImageGauge.value);
    a4ImageGauge.nextElementSibling.textContent = String(state.studySettings.a4ImageHeight);
    applyImageSizes();
  });
  a4ImageGauge?.addEventListener('change', () => {
    saveStudySettings();
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
  if (state.directoryConnected && state.libraryNeedsMigration) {
    try {
      const saved = await persistLibrary();
      if (saved) state.libraryNeedsMigration = false;
    } catch (error) {
      state.error = L(
        `기존 설정을 esper-library.json으로 옮기지 못했어요: ${error instanceof Error ? error.message : String(error)}`,
        `Could not migrate existing preferences to esper-library.json: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }
  mountSupportWidget();
  render();
}

init();
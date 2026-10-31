// Interactive first-run tutorial. It drives the real ESPER screens (sample project,
// setup, study, FSRS review, learning status) and only adds an overlay on top of them.

const STUDY_CARDS = 5;
const GLOW_PADDING = 6;
const BLOCKED_EVENTS = [
  'click', 'dblclick', 'auxclick', 'mousedown', 'mouseup', 'pointerdown', 'pointerup',
  'touchstart', 'touchend', 'contextmenu', 'keydown', 'keypress', 'dragstart', 'dragenter', 'dragover', 'drop', 'submit',
];

function escapeHTML(value) {
  return String(value).replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
}

const isVisible = (element) => {
  if (!element) return false;
  const rect = element.getBoundingClientRect();
  return rect.width > 0 && rect.height > 0;
};

const query = (selector, scope = document) => {
  const element = scope.querySelector(selector);
  return isVisible(element) ? element : null;
};

export function createTutorial(api) {
  const { state } = api;
  const T = (ko, en) => (api.getLanguage() === 'ko' ? ko : en);

  let tour = null;
  let root = null;
  let shield = null;
  let ring = null;
  let cardEl = null;
  let frameId = 0;

  const currentCardId = () => state.study?.cards?.[state.study.index] || null;
  const scheduledCard = () => query('[data-learning-status-card="scheduled"]');
  const validDate = () => (document.querySelector('[data-setup-date]')?.value || '') > todayISO();
  const todayISO = () => new Date().toISOString().slice(0, 10);

  const rateSteps = Array.from({ length: STUDY_CARDS }, (_, index) => [
    ...(index === 0 ? [{
      id: 'image-hint',
      target: () => query('.study-hint-toggle'),
      isDone: () => Boolean(state.study?.hintOpen),
      auto: true,
      title: () => T('이미지 힌트', 'Image hint'),
      body: () => T('hint 보기를 누르면 이 단어와 어울리는 이미지가 나와요.', 'Click Show hint to see an image clue for this word.'),
    }] : []),
    {
      id: `reveal-${index}`,
      target: () => query('.study-card'),
      isDone: () => Boolean(state.study?.flipped) && currentCardId() !== tour.lastRatedId,
      auto: true,
      title: () => T(`카드 ${index + 1}/${STUDY_CARDS}`, `Card ${index + 1} of ${STUDY_CARDS}`),
      body: () => T('답을 떠올린 뒤 카드를 눌러 확인하세요.', 'Think of the answer, then click the card.'),
    },
    {
      id: `rate-${index}`,
      target: () => query('.study-card-actions'),
      enter: () => {
        tour.rateCardId = currentCardId();
        tour.rateBase = api.fsrsRecordForCard(tour.rateCardId)?.reviewLogs.length ?? 0;
      },
      isDone: () => {
        const record = tour.rateCardId ? api.fsrsRecordForCard(tour.rateCardId) : null;
        if (!record || record.reviewLogs.length <= tour.rateBase) return false;
        tour.lastRatedId = tour.rateCardId;
        return true;
      },
      auto: true,
      title: () => T('얼마나 기억했나요?', 'How well did you remember?'),
      body: () => (index === 0
        ? T('솔직하게 고르세요. 버튼의 시간은 FSRS가 정한 다음 복습 시점이에요.', 'Be honest. The time on each button is when FSRS will show it again.')
        : T('솔직하게 평가하세요.', 'Rate honestly.')),
    },
  ]).flat();

  const steps = [
    {
      id: 'language',
      target: () => (state.languageChoicePending
        ? query('[data-choose-language]')?.parentElement || null
        : query('#languageSelect')),
      isDone: () => Boolean(tour.flags.language) && !state.languageChoicePending,
      auto: true,
      keepLabel: () => (state.languageChoicePending ? '' : T('이 언어로 계속', 'Keep this language')),
      onKeep: () => { tour.flags.language = true; },
      title: () => T('언어', 'Language'),
      body: () => T('언어를 고르세요. 튜토리얼도 그 언어로 진행돼요.', 'Pick a language. The tutorial follows it.'),
    },
    {
      id: 'sample',
      target: () => {
        if (state.busy) return query('[data-start-sample]');
        const sample = state.projects.find((project) => project.sourceName === api.getSampleFileName());
        return query(sample ? `[data-open-project="${CSS.escape(sample.id)}"]` : '[data-start-sample]');
      },
      isDone: () => state.page === 'project-setup' && Boolean(state.project),
      auto: true,
      title: () => T('예시 프로젝트', 'Example project'),
      body: () => T('카트리지를 눌러 예시 프로젝트를 여세요. 저장 폴더를 먼저 물어볼 수 있어요.', 'Click the cartridge to open the example. ESPER may ask for a save folder first.'),
    },
    {
      id: 'date',
      target: () => query('[data-date-picker-trigger]')?.closest('.date-picker') || null,
      spotlightTargets: () => [
        query('[data-date-picker-trigger]')?.closest('.date-picker'),
        query('.date-picker-popover:not([hidden])'),
      ].filter(Boolean),
      isDone: () => Boolean(tour.flags.date) && validDate(),
      auto: true,
      keepLabel: () => (validDate() ? T('이 날짜로 계속', 'Keep this date') : ''),
      onKeep: () => { tour.flags.date = true; },
      title: () => T('시험 날짜', 'Exam date'),
      body: () => T('시험 날짜를 고르세요. 한 달 이상 뒤가 좋아요.', 'Pick your exam date. A month or more away works best.'),
    },
    {
      id: 'retention',
      target: () => query('[data-setup-retention]')?.closest('label') || null,
      isDone: () => Boolean(tour.flags.retention),
      auto: true,
      keepLabel: () => T('그대로 두기', 'Keep as is'),
      onKeep: () => { tour.flags.retention = true; },
      title: () => T('목표 기억률', 'Target retention'),
      body: () => T('시험 때 얼마나 기억하고 싶은지 정하세요. 슬라이더를 움직여 보세요.', 'Choose how much you want to remember. Move the slider.'),
    },
    {
      id: 'confirm',
      target: () => query('[data-setup-confirm]'),
      isDone: () => state.page === 'study',
      auto: true,
      title: () => T('시작', 'Start'),
      body: () => T('설정하고 시작하기를 누르세요.', 'Click Save and start.'),
    },
    {
      id: 'start',
      skip: () => Boolean(state.study?.started),
      target: () => query('.study-start-button'),
      isDone: () => Boolean(state.study?.started),
      auto: true,
      title: () => T('학습', 'Study'),
      body: () => T('시작을 누르세요. 카드 5장을 학습해 볼 거예요.', 'Click Start. You will study 5 cards.'),
    },
    ...rateSteps,
    {
      id: 'status',
      target: () => query('[data-global-nav="status"]'),
      isDone: () => state.page === 'learning-status',
      auto: true,
      title: () => T('학습 상태', 'Learning status'),
      body: () => T('방금 한 학습이 어떻게 반영됐는지 볼까요?', 'See how FSRS used your answers.'),
    },
    {
      id: 'recall',
      target: () => scheduledCard(),
      clickAdvance: true,
      auto: true,
      title: () => T('기억 확률', 'Recall probability'),
      body: () => T('지금 이 카드를 기억할 확률이에요. 시간이 지날수록 떨어져요. 눌러서 계속하세요.', 'Your chance of recalling this card now. It falls over time. Click to continue.'),
    },
    {
      id: 'next-review',
      target: () => query('.learning-status-stats', scheduledCard() || document),
      clickAdvance: true,
      auto: true,
      title: () => T('다음 복습', 'Next review'),
      body: () => T('평가를 바탕으로 FSRS가 정한 날짜예요. 잘 기억할수록 더 멀어져요.', 'FSRS set this date from your rating. Better recall means a later date.'),
    },
    {
      id: 'plan',
      target: () => query('.exam-plan-panel'),
      clickAdvance: true,
      auto: true,
      title: () => T('복습 주기 조절', 'Interval control'),
      body: () => T('시험 날짜와 목표 기억률도 복습 간격을 정해요.', 'Your exam date and target retention shape the intervals too.'),
    },
    {
      id: 'finish',
      target: () => query('.fsrs-guide'),
      open: true,
      clickAdvance: true,
      auto: true,
      title: () => T('핵심', 'The idea'),
      body: () => T('복습 날짜는 직접 정하지 않아요. 얼마나 기억했는지만 알려 주면 FSRS가 정해요. 눌러서 마치세요.', 'You never pick review dates. You rate your memory; FSRS decides. Click to finish.'),
    },
  ];

  const currentStep = () => steps[tour.index];

  function isLocked(step) {
    return Boolean(step.isDone && tour.done[step.id]);
  }

  function go() {
    let index = tour.index + 1;
    while (steps[index]?.skip?.()) index += 1;
    if (index >= steps.length) {
      exit();
      return;
    }
    tour.index = index;
    tour.scrolledFor = null;
    steps[index].enter?.();
  }

  function allowedElements() {
    const step = currentStep();
    if (isLocked(step)) return [];
    if (step.spotlightTargets) return step.spotlightTargets();
    const target = step.target?.();
    return target ? [target] : [];
  }

  function block(event) {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
  }

  function guard(event) {
    if (!tour) return;
    const node = event.target;
    if (node instanceof Node && root.contains(node)) {
      if (event.type === 'keydown') event.stopPropagation();
      return;
    }
    if (event.type === 'keydown' && event.key === 'Escape') {
      block(event);
      exit();
      return;
    }
    const inside = node instanceof Node && allowedElements().some((element) => element.contains(node));
    if (!inside) {
      block(event);
      return;
    }
    if (event.type === 'keydown') {
      const editable = node instanceof Element && node.matches('input, textarea, select');
      if (!editable && !['Tab', 'Enter', ' '].includes(event.key)) {
        block(event);
        return;
      }
      // Keep ESPER's own study shortcuts from firing while the tour is running.
      if (!editable) event.stopPropagation();
    }
    if (event.type === 'click' && currentStep().clickAdvance) {
      tour.flags[`click-${currentStep().id}`] = true;
      block(event);
    }
  }

  function detectProgress(step) {
    const done = step.clickAdvance ? tour.flags[`click-${step.id}`] : step.isDone?.();
    if (!done || tour.done[step.id]) return;
    tour.done[step.id] = true;
    if (step.auto) go();
  }

  function cardHTML(step) {
    const keepLabel = step.keepLabel?.() || '';
    return `
      <h2 id="esperTourTitle">${escapeHTML(step.title())}</h2>
      <p class="esper-tour-body">${escapeHTML(step.body())}</p>
      ${state.error ? `<p class="esper-tour-error">${escapeHTML(state.error)}</p>` : ''}
      <div class="esper-tour-actions">
        <button type="button" class="esper-tour-exit" data-tour="exit">${T('종료', 'Exit')}</button>
        <span class="esper-tour-spacer"></span>
        ${keepLabel ? `<button type="button" class="esper-tour-previous" data-tour="keep">${escapeHTML(keepLabel)}</button>` : ''}
      </div>`;
  }

  function placeCard(rect) {
    const margin = 16;
    const width = cardEl.offsetWidth;
    const height = cardEl.offsetHeight;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    const clamp = (value, min, max) => Math.max(min, Math.min(max, value));
    let left;
    let top;
    if (!rect) {
      left = (vw - width) / 2;
      top = (vh - height) / 2;
    } else {
      const options = [
        { left: clamp(rect.left, margin, vw - width - margin), top: rect.bottom + GLOW_PADDING + 12 },
        { left: clamp(rect.left, margin, vw - width - margin), top: rect.top - height - GLOW_PADDING - 12 },
        { left: rect.right + GLOW_PADDING + 12, top: clamp(rect.top, margin, vh - height - margin) },
        { left: rect.left - width - GLOW_PADDING - 12, top: clamp(rect.top, margin, vh - height - margin) },
      ];
      const fit = options.find((option) => option.left >= margin && option.top >= margin &&
        option.left + width <= vw - margin && option.top + height <= vh - margin);
      const fallback = rect.top + rect.height / 2 > vh / 2
        ? { left: vw - width - margin, top: margin }
        : { left: vw - width - margin, top: vh - height - margin };
      ({ left, top } = fit || fallback);
    }
    cardEl.style.left = `${Math.round(left)}px`;
    cardEl.style.top = `${Math.round(top)}px`;
  }

  function tick() {
    const step = currentStep();
    if (step.open) {
      const details = step.target?.();
      if (details && !details.open) details.open = true;
    }
    detectProgress(step);
    if (!tour) return;
    const active = currentStep();
    const target = active.target?.() || null;
    const targets = active.spotlightTargets?.() || (target ? [target] : []);
    const html = cardHTML(active);
    if (html !== tour.lastHTML) {
      cardEl.innerHTML = html;
      tour.lastHTML = html;
    }
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    let rect = null;
    if (target) {
      if (tour.scrolledFor !== target) {
        tour.scrolledFor = target;
        const box = target.getBoundingClientRect();
        if (box.top < 0 || box.bottom > vh) target.scrollIntoView({ block: 'center', behavior: 'smooth' });
      }
      const boxes = targets.map((item) => item.getBoundingClientRect());
      const left = Math.min(...boxes.map((box) => box.left));
      const top = Math.min(...boxes.map((box) => box.top));
      const right = Math.max(...boxes.map((box) => box.right));
      const bottom = Math.max(...boxes.map((box) => box.bottom));
      rect = { left, top, right, bottom, width: right - left, height: bottom - top };
    }
    if (rect) {
      const left = Math.max(0, rect.left - GLOW_PADDING);
      const top = Math.max(0, rect.top - GLOW_PADDING);
      const right = Math.min(vw, rect.right + GLOW_PADDING);
      const bottom = Math.min(vh, rect.bottom + GLOW_PADDING);
      shield.style.clipPath = `path(evenodd, "M0 0H${vw}V${vh}H0ZM${left} ${top}H${right}V${bottom}H${left}Z")`;
      Object.assign(ring.style, {
        display: 'block',
        left: `${left}px`,
        top: `${top}px`,
        width: `${right - left}px`,
        height: `${bottom - top}px`,
      });
    } else {
      shield.style.clipPath = 'none';
      ring.style.display = 'none';
    }
    placeCard(rect);
  }

  function frame() {
    if (!tour) return;
    try {
      tick();
    } catch (error) {
      console.error('Tutorial update failed.', error);
    }
    if (tour) frameId = requestAnimationFrame(frame);
  }

  function onTourClick(event) {
    const button = event.target.closest('[data-tour]');
    if (!button || button.disabled) return;
    const action = button.dataset.tour;
    if (action === 'exit') exit();
    else if (action === 'keep') currentStep().onKeep?.();
  }

  function onFormEvent(event) {
    if (!tour) return;
    const node = event.target;
    if (node instanceof Element && node.matches('[data-language-select]')) tour.flags.language = true;
    if (node instanceof Element && node.matches('[data-setup-date]')) tour.flags.date = true;
    if (node instanceof Element && node.matches('[data-setup-retention]') && event.type === 'change') tour.flags.retention = true;
  }

  function exit() {
    if (!tour) return;
    cancelAnimationFrame(frameId);
    BLOCKED_EVENTS.forEach((type) => document.removeEventListener(type, guard, true));
    document.removeEventListener('change', onFormEvent, true);
    document.removeEventListener('input', onFormEvent, true);
    root.remove();
    root = null;
    tour = null;
    api.render();
  }

  function start() {
    if (tour || state.busy) return;
    state.error = '';
    state.notice = '';
    tour = { index: 0, done: {}, flags: {}, lastRatedId: null };
    state.page = 'workspace';
    api.restoreSampleCartridge();
    api.render();

    root = document.createElement('div');
    root.className = 'esper-tour';
    root.innerHTML = '<div class="esper-tour-shield"></div><div class="esper-tour-ring"></div><div class="esper-tour-card" role="dialog" aria-labelledby="esperTourTitle"></div>';
    document.body.appendChild(root);
    shield = root.querySelector('.esper-tour-shield');
    ring = root.querySelector('.esper-tour-ring');
    cardEl = root.querySelector('.esper-tour-card');
    cardEl.addEventListener('click', onTourClick);
    BLOCKED_EVENTS.forEach((type) => document.addEventListener(type, guard, { capture: true, passive: false }));
    document.addEventListener('change', onFormEvent, true);
    document.addEventListener('input', onFormEvent, true);
    frameId = requestAnimationFrame(frame);
  }

  return { start, exit, isActive: () => Boolean(tour) };
}

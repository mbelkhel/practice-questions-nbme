const SECONDS_PER_QUESTION = 90;
const MB = 1024 * 1024;
// Vercel Function request/response payload hard limit is 4.5MB.
const VERCEL_HARD_REQUEST_LIMIT_BYTES = 4.5 * MB;
const SERVER_UPLOAD_LIMIT_BYTES = Math.floor(4.2 * MB);
const FRONTEND_SAFE_UPLOAD_BYTES = Math.floor(4.0 * MB);
const ALLOWED_FILE_EXTENSIONS = ['.pdf', '.docx', '.doc', '.txt', '.md'];
const DOCX_COMPRESSION_PASSES = [
  { label: 'light', maxDimension: 1500, quality: 0.72, minSourceBytes: 30 * 1024 },
  { label: 'medium', maxDimension: 1100, quality: 0.58, minSourceBytes: 14 * 1024 },
  { label: 'aggressive', maxDimension: 800, quality: 0.45, minSourceBytes: 1 },
  { label: 'ultra', maxDimension: 560, quality: 0.35, minSourceBytes: 1 },
];
const FILE_PLACEHOLDER_TEXT = 'Drag and drop a file here, or click Choose File.';
const EXPLANATION_FALLBACK_PHRASE = 'could not be generated at this time';
const EXPLANATION_SOURCE_MISSING_PHRASE = 'not available in the source';
const EXPLANATION_BATCH_SIZE = 8;
const EXPLANATION_BATCH_MAX_ATTEMPTS = 8;
const EXPLANATION_BATCH_DELAY_MS = 8000;
const EXPLANATION_BATCH_RATE_LIMIT_DELAY_MS = 25000;

const dom = {
  uploadSection: document.getElementById('upload-section'),
  uploadForm: document.getElementById('upload-form'),
  fileDropzone: document.getElementById('file-dropzone'),
  pickFileBtn: document.getElementById('pick-file-btn'),
  selectedFileName: document.getElementById('selected-file-name'),
  documentInput: document.getElementById('document-input'),
  timedMode: document.getElementById('timed-mode'),
  tutorMode: document.getElementById('tutor-mode'),
  uploadBtn: document.getElementById('upload-btn'),
  uploadProgressWrap: document.getElementById('upload-progress-wrap'),
  uploadProgressLabel: document.getElementById('upload-progress-label'),
  uploadStatus: document.getElementById('upload-status'),
  quizSection: document.getElementById('quiz-section'),
  quizTitle: document.getElementById('quiz-title'),
  topPrevBtn: document.getElementById('top-prev-btn'),
  topNextBtn: document.getElementById('top-next-btn'),
  examItem: document.getElementById('exam-item'),
  timerLabel: document.getElementById('timer-label'),
  timer: document.getElementById('timer'),
  stopwatch: document.getElementById('stopwatch'),
  chipTimed: document.getElementById('chip-timed'),
  chipTutor: document.getElementById('chip-tutor'),
  progressLabel: document.getElementById('progress-label'),
  palette: document.getElementById('palette'),
  qCounter: document.getElementById('q-counter'),
  qNumber: document.getElementById('q-number'),
  flagBtn: document.getElementById('flag-btn'),
  stem: document.getElementById('stem'),
  questionImages: document.getElementById('question-images'),
  options: document.getElementById('options'),
  submitAnswerBtn: document.getElementById('submit-answer-btn'),
  explanation: document.getElementById('explanation'),
  prevBtn: document.getElementById('prev-btn'),
  nextBtn: document.getElementById('next-btn'),
  finishBtn: document.getElementById('finish-btn'),
  newTestBtn: document.getElementById('new-test-btn'),
  summary: document.getElementById('summary'),
};

const state = {
  quiz: null,
  currentIndex: 0,
  answers: {},
  selections: {},
  submittedTutor: {},
  flagged: new Set(),
  struck: {},
  tutorMode: true,
  timedMode: false,
  secondsLeft: null,
  elapsedSeconds: 0,
  timerId: null,
  completed: false,
  completionReason: '',
  processing: null,
  localImageMap: {},
  explanationRequestInFlight: {},
  explanationRequestTried: {},
  explanationRequestError: {},
  explanationBackfillToken: 0,
  explanationBackfillRunning: false,
  explanationBackfillAttempts: {},
};

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function setStatus(message, type = '') {
  dom.uploadStatus.textContent = message;
  dom.uploadStatus.classList.remove('error', 'success');
  if (type) {
    dom.uploadStatus.classList.add(type);
  }
}

function setSelectedFileName(file) {
  if (!dom.selectedFileName) {
    return;
  }

  if (!file) {
    dom.selectedFileName.textContent = FILE_PLACEHOLDER_TEXT;
    return;
  }

  dom.selectedFileName.textContent = `${file.name} (${formatFileSize(file.size)})`;
}

function setDropzoneActive(active) {
  if (!dom.fileDropzone) {
    return;
  }
  dom.fileDropzone.classList.toggle('drag-active', Boolean(active));
}

function setUploadProgress(percent, label = '') {
  const clamped = Math.max(0, Math.min(100, Math.round(percent)));
  if (dom.uploadProgressLabel) {
    dom.uploadProgressLabel.textContent = label || (clamped >= 100 ? 'Finalizing quiz...' : 'Working...');
  }
  dom.uploadProgressWrap?.classList.remove('hidden');
}

function hideUploadProgress() {
  dom.uploadProgressWrap?.classList.add('hidden');
}

function hasFileTransfer(event) {
  const types = Array.from(event?.dataTransfer?.types || []);
  return types.includes('Files');
}

function setInputFile(file) {
  if (!file || !dom.documentInput) {
    return;
  }

  try {
    const transfer = new DataTransfer();
    transfer.items.add(file);
    dom.documentInput.files = transfer.files;
  } catch (error) {
    return;
  }

  setSelectedFileName(file);
}

function formatClock(totalSeconds) {
  const seconds = Math.max(0, Math.floor(totalSeconds));
  const hrs = Math.floor(seconds / 3600);
  const mins = Math.floor((seconds % 3600) / 60);
  const secs = seconds % 60;

  if (hrs > 0) {
    return `${hrs.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}:${secs
      .toString()
      .padStart(2, '0')}`;
  }

  return `${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
}

function formatFileSize(bytes) {
  const size = Number(bytes) || 0;
  if (size < 1024) {
    return `${size} B`;
  }
  if (size < MB) {
    return `${(size / 1024).toFixed(1)} KB`;
  }
  return `${(size / MB).toFixed(2)} MB`;
}

function extensionOf(fileName) {
  const name = String(fileName || '').trim().toLowerCase();
  const dot = name.lastIndexOf('.');
  return dot >= 0 ? name.slice(dot) : '';
}

function canCompressDocx(file) {
  return extensionOf(file?.name) === '.docx';
}

function clearLocalImageMap() {
  const urls = new Set(Object.values(state.localImageMap || {}));
  for (const url of urls) {
    if (typeof url === 'string' && /^blob:/i.test(url)) {
      URL.revokeObjectURL(url);
    }
  }
  state.localImageMap = {};
}

function currentQuestion() {
  if (!state.quiz || !state.quiz.questions.length) {
    return null;
  }
  return state.quiz.questions[state.currentIndex];
}

function saveStemMarkup() {
  const question = currentQuestion();
  if (!question) {
    return;
  }
  question.stemHtml = dom.stem.innerHTML;
}

function removeHighlightNode(markNode) {
  const parent = markNode.parentNode;
  while (markNode.firstChild) {
    parent.insertBefore(markNode.firstChild, markNode);
  }
  parent.removeChild(markNode);
  parent.normalize();
}

function applySelectionHighlight() {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) {
    return;
  }

  const range = selection.getRangeAt(0);
  if (range.collapsed || !selection.toString().trim()) {
    return;
  }

  if (!dom.stem.contains(range.commonAncestorContainer)) {
    return;
  }

  const span = document.createElement('span');
  span.className = 'user-highlight';

  try {
    range.surroundContents(span);
  } catch (error) {
    try {
      const fragment = range.extractContents();
      span.appendChild(fragment);
      range.insertNode(span);
    } catch (fallbackError) {
      return;
    }
  }

  selection.removeAllRanges();
  saveStemMarkup();
}

function stopClock() {
  if (state.timerId) {
    clearInterval(state.timerId);
    state.timerId = null;
  }
}

function shouldClockRun() {
  if (!state.quiz || state.completed) {
    return false;
  }

  if (!state.tutorMode) {
    return true;
  }

  const question = currentQuestion();
  if (!question) {
    return false;
  }

  return !Boolean(state.submittedTutor[question.id]);
}

function tickClock() {
  if (!shouldClockRun()) {
    return;
  }

  state.elapsedSeconds += 1;

  if (state.timedMode && typeof state.secondsLeft === 'number') {
    state.secondsLeft -= 1;

    if (state.secondsLeft <= 0) {
      state.secondsLeft = 0;
      updateClockDisplay();
      completeQuiz('time');
      return;
    }
  }

  updateClockDisplay();
}

function startClock() {
  stopClock();
  state.elapsedSeconds = 0;
  state.secondsLeft = state.timedMode ? state.quiz.questions.length * SECONDS_PER_QUESTION : null;
  updateClockDisplay();
  syncClockState();
}

function syncClockState() {
  if (shouldClockRun()) {
    if (!state.timerId) {
      state.timerId = setInterval(tickClock, 1000);
    }
  } else {
    stopClock();
  }
}

function updateClockDisplay() {
  dom.stopwatch.textContent = formatClock(state.elapsedSeconds);

  if (state.timedMode) {
    dom.timerLabel.textContent = state.tutorMode && !shouldClockRun() ? 'Time Remaining (Paused)' : 'Time Remaining';
    dom.timer.textContent = formatClock(state.secondsLeft || 0);
    dom.timer.classList.toggle('warn', (state.secondsLeft || 0) <= 300);
  } else {
    dom.timerLabel.textContent = state.tutorMode && !shouldClockRun() ? 'Untimed (Paused)' : 'Untimed Mode';
    dom.timer.textContent = '--:--';
    dom.timer.classList.remove('warn');
  }
}

function questionTypeOf(question) {
  const type = String(question?.type || '').trim().toLowerCase();
  if (type === 'multi_select' || type === 'true_false') {
    return type;
  }
  return 'single_select';
}

function isMultiSelectQuestion(question) {
  return questionTypeOf(question) === 'multi_select';
}

function normalizeSelectedLabels(value) {
  if (Array.isArray(value)) {
    const uniq = [];
    for (const item of value) {
      const label = String(item || '').trim().toUpperCase();
      if (/^[A-F]$/.test(label) && !uniq.includes(label)) {
        uniq.push(label);
      }
    }
    return uniq;
  }

  const single = String(value || '').trim().toUpperCase();
  if (/^[A-F]$/.test(single)) {
    return [single];
  }

  return [];
}

function correctLabelsFor(question) {
  const byArray = normalizeSelectedLabels(question?.correctOptions);
  if (byArray.length) {
    return byArray;
  }
  return normalizeSelectedLabels(question?.correctOption);
}

function labelsEqual(left, right) {
  if (left.length !== right.length) {
    return false;
  }

  const a = [...left].sort();
  const b = [...right].sort();
  for (let i = 0; i < a.length; i += 1) {
    if (a[i] !== b[i]) {
      return false;
    }
  }
  return true;
}

function isQuestionCorrect(question) {
  const selected = normalizeSelectedLabels(state.answers[question.id]);
  if (!selected.length) {
    return false;
  }

  const correct = correctLabelsFor(question);
  if (!correct.length) {
    return false;
  }

  return labelsEqual(selected, correct);
}

function formatLabels(labels) {
  return labels.length ? labels.join(', ') : '-';
}

function isAnswered(question) {
  return normalizeSelectedLabels(state.answers[question.id]).length > 0;
}

function isWrong(question) {
  if (!state.completed) {
    return false;
  }

  const selected = normalizeSelectedLabels(state.answers[question.id]);
  const correct = correctLabelsFor(question);
  if (!selected.length || !correct.length) {
    return false;
  }

  return !labelsEqual(selected, correct);
}

function explanationFor(question, label) {
  return (
    question.explanations?.[label] ||
    question.sourceExplanation ||
    'No explanation available for this option.'
  );
}

function isFallbackExplanationText(text) {
  const normalized = String(text || '').trim().toLowerCase();
  if (!normalized) {
    return true;
  }
  return normalized.includes(EXPLANATION_FALLBACK_PHRASE) || normalized.includes(EXPLANATION_SOURCE_MISSING_PHRASE);
}

function questionNeedsLiveExplanation(question) {
  if (!question || !Array.isArray(question.options) || question.options.length < 2) {
    return false;
  }

  return question.options.some((option) => isFallbackExplanationText(question.explanations?.[option.label]));
}

function buildQuestionPayload(question) {
  return {
    id: question.id,
    number: question.number,
    type: question.type,
    stem: question.stem,
    options: question.options,
    correctOption: question.correctOption,
    correctOptions: question.correctOptions,
    explanations: question.explanations,
    sourceExplanation: question.sourceExplanation,
    explanationSource: question.explanationSource,
  };
}

function mergeUpdatedQuestion(updated) {
  if (!state.quiz || !updated || !updated.id) {
    return false;
  }

  const current = state.quiz.questions.find((question) => question.id === updated.id);
  if (!current) {
    return false;
  }

  current.explanations = updated.explanations || current.explanations || {};
  current.explanationSource = updated.explanationSource || current.explanationSource;
  if (updated.correctOption) {
    current.correctOption = updated.correctOption;
  }
  if (Array.isArray(updated.correctOptions) && updated.correctOptions.length > 0) {
    current.correctOptions = updated.correctOptions;
  }

  return true;
}

function getBackfillBatch() {
  if (!state.quiz?.questions?.length) {
    return [];
  }

  const batch = [];
  for (const question of state.quiz.questions) {
    if (!questionNeedsLiveExplanation(question)) {
      continue;
    }

    if (state.explanationRequestInFlight[question.id]) {
      continue;
    }

    const attempts = state.explanationBackfillAttempts[question.id] || 0;
    if (attempts >= EXPLANATION_BATCH_MAX_ATTEMPTS) {
      continue;
    }

    batch.push(question);
    if (batch.length >= EXPLANATION_BATCH_SIZE) {
      break;
    }
  }

  return batch;
}

async function requestLiveExplanationBatch(batch) {
  const response = await fetch('/api/quiz/explain-batch', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      questions: batch.map((question) => buildQuestionPayload(question)),
    }),
  });

  const body = parseResponseText(await response.text());
  if (!response.ok) {
    throw new Error(body.error || `Batch explanation generation failed (HTTP ${response.status}).`);
  }

  return body;
}

async function backfillMissingExplanations() {
  if (!state.quiz || state.explanationBackfillRunning) {
    return;
  }

  const seedBatch = getBackfillBatch();
  if (!seedBatch.length) {
    return;
  }

  state.explanationBackfillRunning = true;
  const token = ++state.explanationBackfillToken;

  try {
    while (token === state.explanationBackfillToken) {
      let nextDelayMs = EXPLANATION_BATCH_DELAY_MS;
      const batch = getBackfillBatch();
      if (!batch.length) {
        break;
      }

      for (const question of batch) {
        state.explanationBackfillAttempts[question.id] = (state.explanationBackfillAttempts[question.id] || 0) + 1;
        state.explanationRequestInFlight[question.id] = true;
        state.explanationRequestError[question.id] = '';
      }
      renderAll();

      try {
        const body = await requestLiveExplanationBatch(batch);
        const updatedQuestions = Array.isArray(body.questions) ? body.questions : [];
        const geminiInfo = body.processing?.gemini || {};
        const hitRateLimit = geminiInfo.updatedQuestions === 0 && Number(geminiInfo.rateLimitFallbackCount || 0) > 0;

        for (const updated of updatedQuestions) {
          if (!mergeUpdatedQuestion(updated) || !updated.id) {
            continue;
          }

          const current = state.quiz.questions.find((item) => item.id === updated.id);
          if (current && !questionNeedsLiveExplanation(current)) {
            state.explanationRequestTried[updated.id] = true;
          }
        }

        for (const question of batch) {
          const current = state.quiz.questions.find((item) => item.id === question.id);
          if (!current) {
            continue;
          }

          if (!questionNeedsLiveExplanation(current)) {
            state.explanationRequestError[question.id] = '';
            state.explanationRequestTried[question.id] = true;
            continue;
          }

          if ((state.explanationBackfillAttempts[question.id] || 0) >= EXPLANATION_BATCH_MAX_ATTEMPTS) {
            state.explanationRequestTried[question.id] = true;
          }
        }

        if (hitRateLimit) {
          nextDelayMs = EXPLANATION_BATCH_RATE_LIMIT_DELAY_MS;
        }
      } catch (error) {
        const message = error.message || 'Unable to generate explanation right now.';
        const isRateLimit = /429|rate limit|quota|resource_exhausted/i.test(message);
        if (isRateLimit) {
          nextDelayMs = EXPLANATION_BATCH_RATE_LIMIT_DELAY_MS;
        }
        for (const question of batch) {
          state.explanationRequestError[question.id] = message;
          if ((state.explanationBackfillAttempts[question.id] || 0) >= EXPLANATION_BATCH_MAX_ATTEMPTS) {
            state.explanationRequestTried[question.id] = true;
          }
        }
      } finally {
        for (const question of batch) {
          state.explanationRequestInFlight[question.id] = false;
        }
        renderAll();
      }

      await new Promise((resolve) => setTimeout(resolve, nextDelayMs));
    }
  } finally {
    if (token === state.explanationBackfillToken) {
      state.explanationBackfillRunning = false;
    }
  }
}

async function requestLiveExplanation(question) {
  if (!question || !question.id) {
    return;
  }

  if (state.explanationBackfillRunning) {
    return;
  }

  if (state.explanationRequestInFlight[question.id] || state.explanationRequestTried[question.id]) {
    return;
  }

  state.explanationRequestInFlight[question.id] = true;
  state.explanationRequestTried[question.id] = true;
  state.explanationRequestError[question.id] = '';
  renderAll();

  try {
    const response = await fetch('/api/quiz/explain', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        question: buildQuestionPayload(question),
      }),
    });

    const body = parseResponseText(await response.text());
    if (!response.ok) {
      throw new Error(body.error || `Explanation generation failed (HTTP ${response.status}).`);
    }

    const updated = body.question;
    if (updated && updated.id === question.id) {
      mergeUpdatedQuestion(updated);
    }
  } catch (error) {
    state.explanationRequestError[question.id] = error.message || 'Unable to generate explanation right now.';
  } finally {
    state.explanationRequestInFlight[question.id] = false;
    renderAll();
  }
}

function safeImageSrc(src) {
  const trimmed = String(src || '').trim();
  if (!trimmed) {
    return '';
  }

  const byExact = state.localImageMap?.[trimmed];
  if (byExact) {
    return byExact;
  }

  const lowered = trimmed.toLowerCase();
  const byLower = state.localImageMap?.[lowered];
  if (byLower) {
    return byLower;
  }

  if (/^data:image\//i.test(trimmed)) {
    return trimmed;
  }

  if (/^https?:\/\//i.test(trimmed)) {
    return trimmed;
  }

  if (/^(\/|\.\/|\.\.\/)/.test(trimmed)) {
    return trimmed;
  }

  return '';
}

function renderQuestionImages(question) {
  dom.questionImages.innerHTML = '';

  const images = Array.isArray(question.images) ? question.images : [];
  const safe = images.map((src) => safeImageSrc(src)).filter(Boolean);

  if (!safe.length) {
    dom.questionImages.classList.add('hidden');
    return;
  }

  safe.forEach((src, index) => {
    const img = document.createElement('img');
    img.src = src;
    img.alt = `Question image ${index + 1}`;
    img.loading = 'lazy';
    dom.questionImages.appendChild(img);
  });

  dom.questionImages.classList.remove('hidden');
}

function renderPalette() {
  dom.palette.innerHTML = '';

  state.quiz.questions.forEach((question, index) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'palette-btn';

    if (index === state.currentIndex) {
      btn.classList.add('current');
    }

    if (isWrong(question)) {
      btn.classList.add('wrong');
    } else if (isAnswered(question)) {
      btn.classList.add('answered');
    }

    if (state.flagged.has(question.id)) {
      btn.classList.add('flagged');
    }

    btn.textContent = question.number;
    btn.addEventListener('click', () => {
      state.currentIndex = index;
      renderAll();
    });

    dom.palette.appendChild(btn);
  });
}

function renderExplanation(question) {
  const selectedLabels = normalizeSelectedLabels(state.answers[question.id]);
  const correctLabels = correctLabelsFor(question);
  const tutorSubmitted = Boolean(state.submittedTutor[question.id]);
  const showTutorExplanation = state.tutorMode && tutorSubmitted && selectedLabels.length > 0;
  const showReviewExplanation = state.completed;

  if (!showTutorExplanation && !showReviewExplanation) {
    dom.explanation.classList.add('hidden');
    dom.explanation.innerHTML = '';
    return;
  }

  const lines = [];
  const inFlight = Boolean(state.explanationRequestInFlight[question.id]);
  const liveError = state.explanationRequestError[question.id] || '';

  if (showReviewExplanation || showTutorExplanation) {
    question.options.forEach((option) => {
      const isCorrect = correctLabels.includes(option.label);
      const yourChoice = selectedLabels.includes(option.label) ? ' (Your choice)' : '';
      lines.push(
        `<div class="expl-item"><strong>${isCorrect ? 'Correct' : 'Incorrect'} ${option.label}${yourChoice}</strong>: ${escapeHtml(
          explanationFor(question, option.label),
        )}</div>`,
      );
    });
  }

  if (inFlight) {
    lines.unshift('<div class="expl-item"><em>Generating explanation...</em></div>');
  }
  if (liveError) {
    lines.unshift(`<div class="expl-item"><em>${escapeHtml(liveError)}</em></div>`);
  }

  dom.explanation.innerHTML = `
    <h3>Explanation</h3>
    ${lines.join('')}
    <div class="expl-source">Source: ${escapeHtml(question.explanationSource || 'none')}</div>
  `;
  dom.explanation.classList.remove('hidden');

  if ((showTutorExplanation || showReviewExplanation) && questionNeedsLiveExplanation(question)) {
    setTimeout(() => {
      requestLiveExplanation(question);
    }, 0);
  }
}

function toggleStrike(questionId, optionLabel) {
  if (state.completed) {
    return;
  }

  if (state.tutorMode && state.submittedTutor[questionId]) {
    return;
  }

  if (!state.struck[questionId]) {
    state.struck[questionId] = new Set();
  }

  if (state.struck[questionId].has(optionLabel)) {
    state.struck[questionId].delete(optionLabel);
  } else {
    state.struck[questionId].add(optionLabel);
  }

  renderAll();
}

function selectAnswer(question, optionLabel) {
  if (state.completed) {
    return;
  }

  const lockForTutor = state.tutorMode && Boolean(state.submittedTutor[question.id]);
  if (lockForTutor) {
    return;
  }

  const isMulti = isMultiSelectQuestion(question);
  const currentValue = state.tutorMode ? state.selections[question.id] : state.answers[question.id];
  const currentLabels = normalizeSelectedLabels(currentValue);

  let nextValue = optionLabel;
  if (isMulti) {
    const nextLabels = [...currentLabels];
    if (nextLabels.includes(optionLabel)) {
      const idx = nextLabels.indexOf(optionLabel);
      nextLabels.splice(idx, 1);
    } else {
      nextLabels.push(optionLabel);
    }
    nextValue = nextLabels;
  }

  if (state.tutorMode) {
    if (Array.isArray(nextValue) && nextValue.length === 0) {
      delete state.selections[question.id];
    } else {
      state.selections[question.id] = nextValue;
    }
  } else if (Array.isArray(nextValue) && nextValue.length === 0) {
    delete state.answers[question.id];
  } else {
    state.answers[question.id] = nextValue;
  }

  if (state.struck[question.id] && state.struck[question.id].has(optionLabel)) {
    state.struck[question.id].delete(optionLabel);
  }

  renderAll();
}

function renderQuestion() {
  const question = currentQuestion();
  if (!question) {
    return;
  }

  dom.examItem.textContent = `Exam Section 1: Item ${state.currentIndex + 1} of ${state.quiz.questions.length}`;
  dom.qNumber.textContent = `${question.number}.`;
  dom.qCounter.textContent = `Item ${state.currentIndex + 1} of ${state.quiz.questions.length}`;

  const flagged = state.flagged.has(question.id);
  dom.flagBtn.classList.toggle('marked', flagged);
  dom.flagBtn.textContent = flagged ? 'Marked' : 'Mark';

  if (!question.stemHtml) {
    question.stemHtml = escapeHtml(question.stem).replace(/\n/g, '<br/>');
  }

  dom.stem.innerHTML = question.stemHtml;
  renderQuestionImages(question);

  const isMulti = isMultiSelectQuestion(question);
  const submittedSelection = normalizeSelectedLabels(state.answers[question.id]);
  const pendingSelection = normalizeSelectedLabels(state.selections[question.id]);
  const activeSelection = submittedSelection.length > 0 ? submittedSelection : pendingSelection;
  const correctLabels = correctLabelsFor(question);
  const tutorSubmitted = Boolean(state.submittedTutor[question.id]);
  const reveal = state.completed || (state.tutorMode && tutorSubmitted);
  const lockForTutor = state.tutorMode && tutorSubmitted;

  dom.options.innerHTML = '';

  question.options.forEach((option) => {
    const row = document.createElement('div');
    row.className = 'option-row';

    const strikeSet = state.struck[question.id] || new Set();
    if (strikeSet.has(option.label)) {
      row.classList.add('struck');
    }

    const selectCircle = document.createElement('button');
    selectCircle.type = 'button';
    selectCircle.className = 'select-circle';
    selectCircle.classList.toggle('multi', isMulti);

    if (activeSelection.includes(option.label)) {
      selectCircle.classList.add('selected');
    }

    selectCircle.disabled = state.completed || lockForTutor;
    selectCircle.addEventListener('click', () => selectAnswer(question, option.label));

    const textButton = document.createElement('button');
    textButton.type = 'button';
    textButton.className = 'option-text';
    textButton.innerHTML = `${escapeHtml(option.label)}) ${escapeHtml(option.text)}`;
    textButton.disabled = state.completed || lockForTutor;
    textButton.addEventListener('click', () => toggleStrike(question.id, option.label));

    const badge = document.createElement('span');
    badge.className = 'option-badge';

    if (reveal && correctLabels.includes(option.label)) {
      row.classList.add('review-correct');
      badge.textContent = submittedSelection.includes(option.label) ? 'Your answer + Correct' : 'Correct';
    } else if (reveal && submittedSelection.includes(option.label) && !correctLabels.includes(option.label)) {
      row.classList.add('review-wrong');
      badge.textContent = 'Your answer';
    } else {
      badge.textContent = '';
    }

    row.appendChild(selectCircle);
    row.appendChild(textButton);
    row.appendChild(badge);
    dom.options.appendChild(row);
  });

  renderExplanation(question);

  if (state.tutorMode && !state.completed) {
    dom.submitAnswerBtn.classList.remove('hidden');

    if (tutorSubmitted) {
      dom.submitAnswerBtn.textContent = 'Submitted';
      dom.submitAnswerBtn.disabled = true;
    } else {
      dom.submitAnswerBtn.textContent = isMulti ? 'Submit Answers' : 'Submit Answer';
      dom.submitAnswerBtn.disabled = pendingSelection.length === 0;
    }
  } else {
    dom.submitAnswerBtn.classList.add('hidden');
    dom.submitAnswerBtn.disabled = true;
  }

  const atLastQuestion = state.currentIndex === state.quiz.questions.length - 1;
  dom.prevBtn.disabled = state.currentIndex === 0;
  dom.topPrevBtn.disabled = state.currentIndex === 0;
  dom.nextBtn.disabled = atLastQuestion;
  dom.topNextBtn.disabled = atLastQuestion;
  dom.nextBtn.textContent = 'Next';
  dom.finishBtn.classList.toggle('hidden', state.completed);
  dom.newTestBtn.classList.toggle('hidden', !state.completed);
}

function renderProgress() {
  const answered = state.quiz.questions.filter((question) => isAnswered(question)).length;
  dom.progressLabel.textContent = `${answered} / ${state.quiz.questions.length} answered`;
}

function renderChips() {
  const totalTimed = state.quiz.questions.length * SECONDS_PER_QUESTION;
  dom.chipTimed.textContent = state.timedMode ? `Timed (${formatClock(totalTimed)})` : 'Timed Off';
  dom.chipTutor.textContent = state.tutorMode ? 'Tutor On' : 'Tutor Off';
}

function calculateQuizStats() {
  if (!state.quiz) {
    return {
      total: 0,
      answered: 0,
      correct: 0,
      percent: 0,
    };
  }

  const total = state.quiz.questions.length;
  const answered = state.quiz.questions.filter((q) => isAnswered(q)).length;
  const correct = state.quiz.questions.filter((q) => isQuestionCorrect(q)).length;
  const percent = total ? Math.round((correct / total) * 100) : 0;

  return { total, answered, correct, percent };
}

function showEndBlockPopup() {
  const { total, correct, percent } = calculateQuizStats();
  window.alert(`Block complete\nCorrect: ${correct}/${total}\nScore: ${percent}%`);
}

function renderSummary() {
  if (!state.completed) {
    dom.summary.classList.add('hidden');
    dom.summary.innerHTML = '';
    return;
  }

  const { total, answered, correct, percent } = calculateQuizStats();

  const completionLine =
    state.completionReason === 'time'
      ? '<p><strong>Time expired.</strong> The block was automatically submitted.</p>'
      : '<p>Block submitted.</p>';

  const rows = state.quiz.questions
    .map((question, index) => {
      const selectedLabels = normalizeSelectedLabels(state.answers[question.id]);
      const correctLabels = correctLabelsFor(question);
      const selected = formatLabels(selectedLabels);
      const correctOption = formatLabels(correctLabels);

      let status = 'unanswered';
      let label = 'Unanswered';

      if (selectedLabels.length > 0) {
        if (labelsEqual(selectedLabels, correctLabels)) {
          status = 'correct';
          label = 'Correct';
        } else {
          status = 'wrong';
          label = 'Incorrect';
        }
      }

      return `
        <button type="button" class="review-row ${status}" data-index="${index}">
          <span>Q${question.number}: ${label} (You: ${selected}, Correct: ${correctOption})</span>
          <span>Open</span>
        </button>
      `;
    })
    .join('');

  dom.summary.innerHTML = `
    <h3>Block Report</h3>
    ${completionLine}
    <div class="summary-metrics">
      <span>Correct: ${correct}/${total}</span>
      <span>Answered: ${answered}/${total}</span>
      <span>Score: ${percent}%</span>
      <span>Elapsed: ${formatClock(state.elapsedSeconds)}</span>
    </div>
    <div class="review-grid">${rows}</div>
  `;

  dom.summary.classList.remove('hidden');
}

function renderAll() {
  if (!state.quiz) {
    return;
  }

  dom.quizTitle.textContent = state.quiz.title || 'Generated Quiz';
  renderQuestion();
  syncClockState();
  updateClockDisplay();
  renderChips();
  renderPalette();
  renderProgress();
  renderSummary();
}

function completeQuiz(reason = '') {
  if (state.completed) {
    return;
  }

  state.completed = true;
  state.completionReason = reason;
  stopClock();
  renderAll();
}

function resetToUploadScreen() {
  stopClock();
  clearLocalImageMap();
  state.explanationBackfillToken += 1;
  state.explanationBackfillRunning = false;
  state.quiz = null;
  state.currentIndex = 0;
  state.answers = {};
  state.selections = {};
  state.submittedTutor = {};
  state.flagged = new Set();
  state.struck = {};
  state.completed = false;
  state.completionReason = '';
  state.processing = null;
  state.secondsLeft = null;
  state.elapsedSeconds = 0;
  state.explanationRequestInFlight = {};
  state.explanationRequestTried = {};
  state.explanationRequestError = {};
  state.explanationBackfillAttempts = {};

  dom.documentInput.value = '';
  setSelectedFileName(null);
  hideUploadProgress();
  dom.quizSection.classList.add('hidden');
  dom.summary.classList.add('hidden');
  dom.uploadSection.classList.remove('hidden');
  setStatus('Upload a file to generate a new quiz.');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function canvasToBlob(canvas, type, quality) {
  return new Promise((resolve) => {
    canvas.toBlob((blob) => resolve(blob), type, quality);
  });
}

function docxImageMimeType(entryName) {
  if (/\.png$/i.test(entryName)) {
    return 'image/png';
  }
  if (/\.gif$/i.test(entryName)) {
    return 'image/gif';
  }
  if (/\.bmp$/i.test(entryName)) {
    return 'image/bmp';
  }
  if (/\.svg$/i.test(entryName)) {
    return 'image/svg+xml';
  }
  if (/\.webp$/i.test(entryName)) {
    return 'image/webp';
  }
  return 'image/jpeg';
}

function decodeXmlEntities(text) {
  return String(text || '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number.parseInt(code, 10) || 32));
}

function normalizeFallbackText(text) {
  return String(text || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function normalizeDocxTargetPath(target) {
  const parts = String(target || '')
    .replace(/\\/g, '/')
    .split('/')
    .filter(Boolean);

  const stack = [];
  for (const part of parts) {
    if (part === '.') {
      continue;
    }
    if (part === '..') {
      stack.pop();
      continue;
    }
    stack.push(part);
  }

  return stack.join('/');
}

function parseDocxRelationships(xml) {
  const relMap = {};
  if (!xml) {
    return relMap;
  }

  const regex = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*>/gi;
  let match = regex.exec(xml);

  while (match) {
    const relId = match[1];
    const target = normalizeDocxTargetPath(match[2]);
    if (relId && target) {
      relMap[relId] = target;
    }
    match = regex.exec(xml);
  }

  return relMap;
}

function insertDocxImageMarkers(xml, relMap) {
  let output = String(xml || '');

  output = output.replace(/<w:drawing[\s\S]*?<\/w:drawing>/gi, (drawingXml) => {
    const relMatch = drawingXml.match(/\br:embed="([^"]+)"/i) || drawingXml.match(/\br:link="([^"]+)"/i);
    if (!relMatch) {
      return '\n';
    }

    const target = relMap[relMatch[1]] || '';
    const fileName = target.split('/').pop();
    if (!fileName) {
      return '\n';
    }

    return `\n[IMAGE:${fileName}]\n`;
  });

  output = output.replace(/<w:pict[\s\S]*?<\/w:pict>/gi, (pictXml) => {
    const relMatch =
      pictXml.match(/\br:id="([^"]+)"/i) ||
      pictXml.match(/\bo:relid="([^"]+)"/i) ||
      pictXml.match(/\br:link="([^"]+)"/i);
    if (!relMatch) {
      return '\n';
    }

    const target = relMap[relMatch[1]] || '';
    const fileName = target.split('/').pop();
    if (!fileName) {
      return '\n';
    }

    return `\n[IMAGE:${fileName}]\n`;
  });

  return output;
}

function xmlToNormalizedDocxText(xml) {
  return normalizeFallbackText(
    decodeXmlEntities(
      String(xml || '')
        .replace(/<w:tab\/>/gi, '\t')
        .replace(/<w:br[^>]*\/>/gi, '\n')
        .replace(/<\/w:p>/gi, '\n')
        .replace(/<\/w:tr>/gi, '\n')
        .replace(/<[^>]+>/g, ''),
    ),
  );
}

async function buildDocxLocalImageMap(zip) {
  const imageMap = {};
  const mediaEntries = Object.keys(zip.files).filter((entryName) =>
    /^word\/media\/.+\.(jpe?g|png|gif|webp|bmp|svg)$/i.test(entryName),
  );

  for (const entryName of mediaEntries) {
    const entry = zip.file(entryName);
    if (!entry) {
      continue;
    }

    const bytes = await entry.async('uint8array');
    const fileName = entryName.split('/').pop();
    if (!fileName || !bytes.length) {
      continue;
    }

    const url = URL.createObjectURL(new Blob([bytes], { type: docxImageMimeType(entryName) }));
    imageMap[fileName] = url;
    imageMap[fileName.toLowerCase()] = url;
    imageMap[entryName] = url;
    imageMap[entryName.toLowerCase()] = url;
    imageMap[`media/${fileName}`] = url;
    imageMap[`media/${fileName}`.toLowerCase()] = url;
  }

  return imageMap;
}

async function extractLargeDocxPayload(file) {
  if (typeof window.JSZip === 'undefined') {
    throw new Error('Compression library unavailable in browser.');
  }

  const zip = await window.JSZip.loadAsync(await file.arrayBuffer());
  const documentEntry = zip.file('word/document.xml');
  if (!documentEntry) {
    throw new Error('Unable to read DOCX body XML.');
  }

  const relationshipsXml = (await zip.file('word/_rels/document.xml.rels')?.async('string')) || '';
  const relMap = parseDocxRelationships(relationshipsXml);
  const documentXml = await documentEntry.async('string');
  const withImageMarkers = insertDocxImageMarkers(documentXml, relMap);
  const text = xmlToNormalizedDocxText(withImageMarkers);
  const imageMap = await buildDocxLocalImageMap(zip);

  if (!text) {
    throw new Error('Could not extract text from DOCX.');
  }

  return { text, imageMap };
}

async function compressImageBytes(uint8Array, mimeType, options = {}) {
  const originalBlob = new Blob([uint8Array], { type: mimeType });
  const objectUrl = URL.createObjectURL(originalBlob);

  try {
    const image = await new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error('Image decode failed.'));
      img.src = objectUrl;
    });

    const maxDimension = Math.max(120, Number(options.maxDimension) || 1500);
    const scale = Math.min(1, maxDimension / Math.max(image.naturalWidth || 1, image.naturalHeight || 1));
    const width = Math.max(1, Math.round((image.naturalWidth || 1) * scale));
    const height = Math.max(1, Math.round((image.naturalHeight || 1) * scale));

    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext('2d');
    if (!context) {
      return null;
    }

    context.drawImage(image, 0, 0, width, height);

    const primaryType = /image\/(jpeg|png|webp)/i.test(mimeType) ? mimeType : 'image/jpeg';
    const quality = primaryType === 'image/png' ? undefined : Number(options.quality) || 0.7;
    const compressedBlob = await canvasToBlob(canvas, primaryType, quality);

    if (!compressedBlob) {
      return null;
    }

    return new Uint8Array(await compressedBlob.arrayBuffer());
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

async function compressDocxForUpload(file, pass) {
  if (typeof window.JSZip === 'undefined') {
    throw new Error('Compression library unavailable in browser.');
  }

  const zip = await window.JSZip.loadAsync(await file.arrayBuffer());
  const mediaEntries = Object.keys(zip.files).filter((entryName) =>
    /^word\/media\/.+\.(jpe?g|png|webp)$/i.test(entryName),
  );

  if (!mediaEntries.length) {
    return {
      file,
      changed: false,
      compressedImages: 0,
    };
  }

  let changed = false;
  let compressedImages = 0;

  for (const entryName of mediaEntries) {
    const entry = zip.file(entryName);
    if (!entry) {
      continue;
    }

    const original = await entry.async('uint8array');
    if (original.length < (pass?.minSourceBytes || 1)) {
      continue;
    }

    const mimeType = docxImageMimeType(entryName);
    let compressed = null;
    try {
      compressed = await compressImageBytes(original, mimeType, pass);
    } catch (error) {
      compressed = null;
    }

    if (compressed && compressed.length > 0 && compressed.length < original.length) {
      zip.file(entryName, compressed);
      changed = true;
      compressedImages += 1;
    }
  }

  if (!changed) {
    return {
      file,
      changed: false,
      compressedImages: 0,
    };
  }

  const blob = await zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 },
  });

  if (blob.size >= file.size) {
    return {
      file,
      changed: false,
      compressedImages: 0,
    };
  }

  return {
    file: new File([blob], file.name, {
      type: file.type || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      lastModified: Date.now(),
    }),
    changed: true,
    compressedImages,
  };
}

async function prepareUploadFile(rawFile) {
  const ext = extensionOf(rawFile.name);
  if (!ALLOWED_FILE_EXTENSIONS.includes(ext)) {
    throw new Error(
      `Unsupported file type "${ext || 'unknown'}". Use ${ALLOWED_FILE_EXTENSIONS.join(', ')}.`,
    );
  }

  if (rawFile.size <= FRONTEND_SAFE_UPLOAD_BYTES) {
    return {
      file: rawFile,
      note: '',
    };
  }

  if (!canCompressDocx(rawFile)) {
    throw new Error(
      `File is too large (${formatFileSize(rawFile.size)}). Vercel accepts about ${formatFileSize(
        VERCEL_HARD_REQUEST_LIMIT_BYTES,
      )} total request size, so keep uploads under ${formatFileSize(SERVER_UPLOAD_LIMIT_BYTES)}.`,
    );
  }

  setStatus(
    `File is ${formatFileSize(rawFile.size)}. Compressing DOCX images to fit Vercel upload limits...`,
  );

  let candidate = rawFile;
  const notes = [];

  for (const pass of DOCX_COMPRESSION_PASSES) {
    const result = await compressDocxForUpload(candidate, pass);
    candidate = result.file;

    if (result.changed) {
      notes.push(
        `${pass.label} compression pass adjusted ${result.compressedImages} image(s); file is now ${formatFileSize(candidate.size)}.`,
      );
    }

    if (candidate.size <= SERVER_UPLOAD_LIMIT_BYTES) {
      return {
        file: candidate,
        note: `DOCX compressed from ${formatFileSize(rawFile.size)} to ${formatFileSize(candidate.size)} before upload. ${notes.join(' ')}`.trim(),
      };
    }
  }

  throw new Error(
    `File is still too large after aggressive compression (${formatFileSize(
      candidate.size,
    )}). Try again and the app will use a large-DOCX local extraction path that preserves image references.`,
  );
}

function parseResponseText(rawBody) {
  let body = {};
  if (rawBody) {
    try {
      body = JSON.parse(rawBody);
    } catch (error) {
      if (/<!doctype html|<html/i.test(rawBody)) {
        throw new Error('Server returned a non-JSON error page. Check deployment logs and retry.');
      }
      throw new Error(rawBody.slice(0, 240));
    }
  }
  return body;
}

async function requestJsonWithProgress(url, requestOptions, onUploadProgress) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open(requestOptions?.method || 'POST', url, true);

    const headers = requestOptions?.headers || {};
    for (const [key, value] of Object.entries(headers)) {
      xhr.setRequestHeader(key, value);
    }

    if (typeof onUploadProgress === 'function' && xhr.upload) {
      xhr.upload.onprogress = (event) => {
        if (!event.lengthComputable || event.total <= 0) {
          return;
        }
        onUploadProgress(event.loaded / event.total);
      };
    }

    xhr.onerror = () => {
      reject(new Error('Network error while uploading file.'));
    };

    xhr.onload = () => {
      try {
        const body = parseResponseText(xhr.responseText || '');
        resolve({
          ok: xhr.status >= 200 && xhr.status < 300,
          status: xhr.status,
          body,
        });
      } catch (error) {
        reject(error);
      }
    };

    xhr.send(requestOptions?.body || null);
  });
}

function applyQuizToState(body, localImageMap = {}) {
  state.explanationBackfillToken += 1;
  state.explanationBackfillRunning = false;
  state.quiz = body.quiz;
  state.currentIndex = 0;
  state.answers = {};
  state.selections = {};
  state.submittedTutor = {};
  state.flagged = new Set();
  state.struck = {};
  state.completed = false;
  state.completionReason = '';
  state.processing = body.processing;
  state.localImageMap = localImageMap || {};
  state.explanationRequestInFlight = {};
  state.explanationRequestTried = {};
  state.explanationRequestError = {};
  state.explanationBackfillAttempts = {};

  state.timedMode = dom.timedMode.checked;
  state.tutorMode = dom.tutorMode.checked;

  dom.uploadSection.classList.add('hidden');
  dom.quizSection.classList.remove('hidden');
  dom.summary.classList.add('hidden');

  startClock();
}

async function uploadDocument(event) {
  event.preventDefault();

  const selectedFile = dom.documentInput.files?.[0];
  if (!selectedFile) {
    setStatus('Select a document first.', 'error');
    return;
  }

  let uploadFile = selectedFile;
  let uploadPrepNote = '';
  let localImageMap = {};
  let requestUrl = '/api/quiz/upload';
  let requestOptions = null;

  clearLocalImageMap();
  hideUploadProgress();
  setUploadProgress(0, 'Preparing upload...');

  try {
    const ext = extensionOf(selectedFile.name);
    if (ext === '.docx' && selectedFile.size > FRONTEND_SAFE_UPLOAD_BYTES) {
      setStatus(
        `Large DOCX detected (${formatFileSize(
          selectedFile.size,
        )}). Extracting text and image references locally to bypass Vercel upload limits...`,
      );

      const extracted = await extractLargeDocxPayload(selectedFile);
      localImageMap = extracted.imageMap;
      uploadPrepNote =
        'Large DOCX was parsed locally for upload safety. Embedded images are mapped by reference and should appear in related questions.';

      requestUrl = '/api/quiz/from-text';
      requestOptions = {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          text: extracted.text,
          fileName: selectedFile.name,
          timedMode: dom.timedMode.checked ? 'true' : 'false',
          tutorMode: dom.tutorMode.checked ? 'true' : 'false',
        }),
      };
    } else {
      const prepared = await prepareUploadFile(selectedFile);
      uploadFile = prepared.file;
      uploadPrepNote = prepared.note;

      const formData = new FormData();
      formData.append('document', uploadFile);
      formData.append('timedMode', dom.timedMode.checked ? 'true' : 'false');
      formData.append('tutorMode', dom.tutorMode.checked ? 'true' : 'false');

      requestOptions = {
        method: 'POST',
        body: formData,
      };
    }
  } catch (error) {
    hideUploadProgress();
    setStatus(error.message || 'Unable to prepare upload file.', 'error');
    return;
  }

  setStatus('Uploading file and building quiz...');
  dom.uploadBtn.disabled = true;
  setUploadProgress(3, 'Uploading file...');

  try {
    const response = await requestJsonWithProgress(requestUrl, requestOptions, (fraction) => {
      const percent = Math.max(3, Math.min(92, Math.round(fraction * 92)));
      setUploadProgress(percent, 'Uploading file...');
    });

    setUploadProgress(96, 'Processing questions and explanations...');
    const body = response.body;

    if (!response.ok) {
      throw new Error(body.error || `Upload failed (HTTP ${response.status}).`);
    }

    applyQuizToState(body, localImageMap);
    setUploadProgress(100, 'Quiz ready.');

    const parseInfo = body.processing?.parsing;
    const geminiInfo = body.processing?.gemini;
    const parts = [];

    if (parseInfo) {
      parts.push(
        `${parseInfo.totalQuestions} questions parsed, ${parseInfo.answersMapped} answers linked, ${parseInfo.explanationsMapped} explanations mapped.`,
      );
    }

    if (geminiInfo?.attempted) {
      parts.push(
        `Gemini updated ${geminiInfo.updatedQuestions} question(s)${geminiInfo.failedChunks ? ` with ${geminiInfo.failedChunks} failed chunk(s)` : ''}.`,
      );
      if (geminiInfo.reason) {
        parts.push(geminiInfo.reason);
      }
    } else if (geminiInfo?.reason) {
      parts.push(`Gemini skipped: ${geminiInfo.reason}`);
    }

    if (state.timedMode) {
      parts.push(`Timed block set to ${formatClock(state.quiz.questions.length * SECONDS_PER_QUESTION)}.`);
    }

    if (uploadPrepNote) {
      parts.push(uploadPrepNote);
    }

    setStatus(parts.join(' '), 'success');
    renderAll();
    void backfillMissingExplanations();
  } catch (error) {
    hideUploadProgress();
    setStatus(error.message || 'Unable to process file.', 'error');
  } finally {
    if (!state.quiz) {
      hideUploadProgress();
    }
    dom.uploadBtn.disabled = false;
  }
}

function goPrevious() {
  if (!state.quiz || state.currentIndex <= 0) {
    return;
  }

  state.currentIndex -= 1;
  renderAll();
}

function goNext() {
  if (!state.quiz) {
    return;
  }

  if (state.currentIndex >= state.quiz.questions.length - 1) {
    return;
  }

  state.currentIndex += 1;
  renderAll();
}

function toggleFlag() {
  const question = currentQuestion();
  if (!question) {
    return;
  }

  if (state.flagged.has(question.id)) {
    state.flagged.delete(question.id);
  } else {
    state.flagged.add(question.id);
  }

  renderAll();
}

function onStemMouseUp() {
  setTimeout(() => {
    applySelectionHighlight();
  }, 0);
}

function onStemClick(event) {
  const highlight = event.target.closest('.user-highlight');
  if (!highlight || !dom.stem.contains(highlight)) {
    return;
  }

  removeHighlightNode(highlight);
  saveStemMarkup();
}

function onSummaryClick(event) {
  const row = event.target.closest('.review-row');
  if (!row) {
    return;
  }

  const index = Number.parseInt(row.dataset.index || '-1', 10);
  if (!Number.isInteger(index) || index < 0 || index >= state.quiz.questions.length) {
    return;
  }

  state.currentIndex = index;
  renderAll();
  window.scrollTo({ top: dom.quizSection.offsetTop, behavior: 'smooth' });
}

function onDocumentInputChange() {
  const file = dom.documentInput.files?.[0] || null;
  setSelectedFileName(file);
}

function onDropzoneDragEnter(event) {
  if (!hasFileTransfer(event)) {
    return;
  }
  event.preventDefault();
  setDropzoneActive(true);
}

function onDropzoneDragOver(event) {
  if (!hasFileTransfer(event)) {
    return;
  }
  event.preventDefault();
  event.dataTransfer.dropEffect = 'copy';
  setDropzoneActive(true);
}

function onDropzoneDragLeave(event) {
  if (!event.currentTarget.contains(event.relatedTarget)) {
    setDropzoneActive(false);
  }
}

function onDropzoneDrop(event) {
  if (!hasFileTransfer(event)) {
    return;
  }
  event.preventDefault();
  setDropzoneActive(false);
  const file = event.dataTransfer.files?.[0];
  if (!file) {
    return;
  }
  setInputFile(file);
}

function onPickFileClick() {
  dom.documentInput?.click();
}

function submitCurrentTutorAnswer() {
  if (!state.quiz || !state.tutorMode || state.completed) {
    return;
  }

  const question = currentQuestion();
  if (!question) {
    return;
  }

  if (state.submittedTutor[question.id]) {
    return;
  }

  const selectionLabels = normalizeSelectedLabels(state.selections[question.id]);
  if (selectionLabels.length === 0) {
    return;
  }

  if (isMultiSelectQuestion(question)) {
    state.answers[question.id] = selectionLabels;
  } else {
    state.answers[question.id] = selectionLabels[0];
  }
  state.submittedTutor[question.id] = true;
  renderAll();
}

function endBlock() {
  if (!state.quiz || state.completed) {
    return;
  }
  showEndBlockPopup();
  completeQuiz('manual');
}

function makeNewTest() {
  resetToUploadScreen();
}

dom.uploadForm.addEventListener('submit', uploadDocument);
dom.documentInput.addEventListener('change', onDocumentInputChange);
dom.pickFileBtn.addEventListener('click', onPickFileClick);
dom.fileDropzone.addEventListener('dragenter', onDropzoneDragEnter);
dom.fileDropzone.addEventListener('dragover', onDropzoneDragOver);
dom.fileDropzone.addEventListener('dragleave', onDropzoneDragLeave);
dom.fileDropzone.addEventListener('drop', onDropzoneDrop);
dom.prevBtn.addEventListener('click', goPrevious);
dom.nextBtn.addEventListener('click', goNext);
dom.topPrevBtn.addEventListener('click', goPrevious);
dom.topNextBtn.addEventListener('click', goNext);
dom.flagBtn.addEventListener('click', toggleFlag);
dom.finishBtn.addEventListener('click', endBlock);
dom.newTestBtn.addEventListener('click', makeNewTest);
dom.submitAnswerBtn.addEventListener('click', submitCurrentTutorAnswer);
dom.stem.addEventListener('mouseup', onStemMouseUp);
dom.stem.addEventListener('touchend', onStemMouseUp, { passive: true });
dom.stem.addEventListener('click', onStemClick);
dom.summary.addEventListener('click', onSummaryClick);

window.addEventListener('beforeunload', () => {
  stopClock();
  clearLocalImageMap();
});

setSelectedFileName(null);
hideUploadProgress();

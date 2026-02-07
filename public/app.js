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

const dom = {
  uploadSection: document.getElementById('upload-section'),
  uploadForm: document.getElementById('upload-form'),
  documentInput: document.getElementById('document-input'),
  timedMode: document.getElementById('timed-mode'),
  tutorMode: document.getElementById('tutor-mode'),
  useGemini: document.getElementById('use-gemini'),
  uploadBtn: document.getElementById('upload-btn'),
  uploadStatus: document.getElementById('upload-status'),
  quizSection: document.getElementById('quiz-section'),
  quizTitle: document.getElementById('quiz-title'),
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

function isAnswered(question) {
  return Boolean(state.answers[question.id]);
}

function isWrong(question) {
  if (!state.completed) {
    return false;
  }

  const selected = state.answers[question.id];
  if (!selected || !question.correctOption) {
    return false;
  }

  return selected !== question.correctOption;
}

function explanationFor(question, label) {
  return (
    question.explanations?.[label] ||
    question.sourceExplanation ||
    'No explanation available for this option.'
  );
}

function safeImageSrc(src) {
  const trimmed = String(src || '').trim();
  if (!trimmed) {
    return '';
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
  const selected = state.answers[question.id];
  const tutorSubmitted = Boolean(state.submittedTutor[question.id]);
  const showTutorExplanation = state.tutorMode && tutorSubmitted && selected;
  const showReviewExplanation = state.completed;

  if (!showTutorExplanation && !showReviewExplanation) {
    dom.explanation.classList.add('hidden');
    dom.explanation.innerHTML = '';
    return;
  }

  const lines = [];

  if (showReviewExplanation || showTutorExplanation) {
    question.options.forEach((option) => {
      const isCorrect = question.correctOption === option.label;
      const yourChoice = selected === option.label ? ' (Your choice)' : '';
      lines.push(
        `<div class="expl-item"><strong>${isCorrect ? 'Correct' : 'Incorrect'} ${option.label}${yourChoice}</strong>: ${escapeHtml(
          explanationFor(question, option.label),
        )}</div>`,
      );
    });
  }

  dom.explanation.innerHTML = `
    <h3>Explanation</h3>
    ${lines.join('')}
    <div class="expl-source">Source: ${escapeHtml(question.explanationSource || 'none')}</div>
  `;
  dom.explanation.classList.remove('hidden');
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

  if (state.tutorMode) {
    state.selections[question.id] = optionLabel;
  } else {
    state.answers[question.id] = optionLabel;
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

  const submittedSelection = state.answers[question.id] || null;
  const pendingSelection = state.selections[question.id] || null;
  const activeSelection = submittedSelection || pendingSelection;
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

    if (activeSelection === option.label) {
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

    if (reveal && question.correctOption === option.label) {
      row.classList.add('review-correct');
      badge.textContent = submittedSelection === option.label ? 'Your answer + Correct' : 'Correct';
    } else if (
      reveal &&
      submittedSelection === option.label &&
      question.correctOption &&
      submittedSelection !== question.correctOption
    ) {
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
      dom.submitAnswerBtn.textContent = 'Submit Answer';
      dom.submitAnswerBtn.disabled = !Boolean(pendingSelection);
    }
  } else {
    dom.submitAnswerBtn.classList.add('hidden');
    dom.submitAnswerBtn.disabled = true;
  }

  const atLastQuestion = state.currentIndex === state.quiz.questions.length - 1;
  dom.prevBtn.disabled = state.currentIndex === 0;
  dom.nextBtn.disabled = atLastQuestion;
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

function renderSummary() {
  if (!state.completed) {
    dom.summary.classList.add('hidden');
    dom.summary.innerHTML = '';
    return;
  }

  const total = state.quiz.questions.length;
  const answered = state.quiz.questions.filter((q) => state.answers[q.id]).length;
  const correct = state.quiz.questions.filter(
    (q) => q.correctOption && state.answers[q.id] === q.correctOption,
  ).length;

  const completionLine =
    state.completionReason === 'time'
      ? '<p><strong>Time expired.</strong> The block was automatically submitted.</p>'
      : '<p>Block submitted.</p>';

  const rows = state.quiz.questions
    .map((question, index) => {
      const selected = state.answers[question.id] || '-';
      const correctOption = question.correctOption || '-';

      let status = 'unanswered';
      let label = 'Unanswered';

      if (selected !== '-') {
        if (selected === correctOption) {
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
      <span>Score: ${total ? Math.round((correct / total) * 100) : 0}%</span>
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

  dom.documentInput.value = '';
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
  if (/\.webp$/i.test(entryName)) {
    return 'image/webp';
  }
  return 'image/jpeg';
}

function baseNameWithoutExtension(fileName) {
  const name = String(fileName || 'document').trim();
  const dot = name.lastIndexOf('.');
  return (dot > 0 ? name.slice(0, dot) : name).replace(/[^\w.-]+/g, '_').slice(0, 90) || 'document';
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

async function extractDocxTextFallbackFile(file) {
  if (typeof window.JSZip === 'undefined') {
    throw new Error('Compression library unavailable in browser.');
  }

  const zip = await window.JSZip.loadAsync(await file.arrayBuffer());
  const xmlEntries = Object.keys(zip.files).filter((entryName) =>
    /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes)\.xml$/i.test(entryName),
  );

  if (!xmlEntries.length) {
    throw new Error('Unable to read text from DOCX fallback.');
  }

  const chunks = [];

  for (const entryName of xmlEntries) {
    const entry = zip.file(entryName);
    if (!entry) {
      continue;
    }

    const xml = await entry.async('string');
    const text = decodeXmlEntities(
      xml
        .replace(/<w:tab\/>/gi, '\t')
        .replace(/<w:br[^>]*\/>/gi, '\n')
        .replace(/<\/w:p>/gi, '\n')
        .replace(/<\/w:tr>/gi, '\n')
        .replace(/<[^>]+>/g, ''),
    );

    if (text.trim()) {
      chunks.push(text);
    }
  }

  const mergedText = normalizeFallbackText(chunks.join('\n\n'));
  if (!mergedText) {
    throw new Error('Could not extract text from DOCX fallback.');
  }

  return new File([mergedText], `${baseNameWithoutExtension(file.name)}-text-only.txt`, {
    type: 'text/plain',
    lastModified: Date.now(),
  });
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

  setStatus(
    `Compressed file is still ${formatFileSize(
      candidate.size,
    )}. Switching to text-only fallback to fit Vercel limits (images will be skipped).`,
  );

  const textFallback = await extractDocxTextFallbackFile(rawFile);
  if (textFallback.size <= SERVER_UPLOAD_LIMIT_BYTES) {
    return {
      file: textFallback,
      note: `The original DOCX remained too large for Vercel with embedded images. Uploaded a text-only fallback (${formatFileSize(
        textFallback.size,
      )}) so parsing can continue. Split the source into smaller docs if you need all images preserved.`,
    };
  }

  throw new Error(
    `File is still too large even after aggressive compression and text fallback. Size is ${formatFileSize(
      rawFile.size,
    )}. Split the file into multiple smaller documents and upload one at a time.`,
  );
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

  try {
    const prepared = await prepareUploadFile(selectedFile);
    uploadFile = prepared.file;
    uploadPrepNote = prepared.note;
  } catch (error) {
    setStatus(error.message || 'Unable to prepare upload file.', 'error');
    return;
  }

  setStatus('Processing document and building quiz...');
  dom.uploadBtn.disabled = true;

  const formData = new FormData();
  formData.append('document', uploadFile);
  formData.append('timedMode', dom.timedMode.checked ? 'true' : 'false');
  formData.append('tutorMode', dom.tutorMode.checked ? 'true' : 'false');
  formData.append('useGemini', dom.useGemini.checked ? 'true' : 'false');

  try {
    const response = await fetch('/api/quiz/upload', {
      method: 'POST',
      body: formData,
    });

    const rawBody = await response.text();
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

    if (!response.ok) {
      throw new Error(body.error || `Upload failed (HTTP ${response.status}).`);
    }

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

    state.timedMode = dom.timedMode.checked;
    state.tutorMode = dom.tutorMode.checked;

    dom.uploadSection.classList.add('hidden');
    dom.quizSection.classList.remove('hidden');
    dom.summary.classList.add('hidden');

    startClock();

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
  } catch (error) {
    setStatus(error.message || 'Unable to process file.', 'error');
  } finally {
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

  const selection = state.selections[question.id];
  if (!selection) {
    return;
  }

  state.answers[question.id] = selection;
  state.submittedTutor[question.id] = true;
  renderAll();
}

function endBlock() {
  if (!state.quiz || state.completed) {
    return;
  }
  completeQuiz('manual');
}

function makeNewTest() {
  resetToUploadScreen();
}

dom.uploadForm.addEventListener('submit', uploadDocument);
dom.prevBtn.addEventListener('click', goPrevious);
dom.nextBtn.addEventListener('click', goNext);
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
});

const { GoogleGenAI } = require('@google/genai');

const DEFAULT_MODEL_CHAIN = ['gemma-3-27b-it', 'gemma-3-12b-it'];
const EXPLANATION_PLACEHOLDER_PATTERNS = [
  'could not be generated at this time',
  'not available in the source',
  'enable gemini with a valid api key',
];

function isPlaceholderExplanationText(text) {
  const normalized = String(text || '').trim().toLowerCase();
  if (!normalized) {
    return true;
  }
  return EXPLANATION_PLACEHOLDER_PATTERNS.some((pattern) => normalized.includes(pattern));
}

function chunkArray(items, size) {
  const chunks = [];
  for (let i = 0; i < items.length; i += size) {
    chunks.push(items.slice(i, i + size));
  }
  return chunks;
}

function extractJson(text) {
  if (!text) {
    return null;
  }

  const trimmed = text.trim();

  if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
    try {
      return JSON.parse(trimmed);
    } catch (error) {
      // Continue to fenced fallback.
    }
  }

  const fenced = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenced) {
    try {
      return JSON.parse(fenced[1].trim());
    } catch (error) {
      return null;
    }
  }

  const firstBrace = trimmed.indexOf('[');
  const lastBrace = trimmed.lastIndexOf(']');
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    try {
      return JSON.parse(trimmed.slice(firstBrace, lastBrace + 1));
    } catch (error) {
      return null;
    }
  }

  return null;
}

function parseModelChain(value) {
  if (Array.isArray(value)) {
    return value.map((item) => String(item || '').trim()).filter(Boolean);
  }
  return String(value || '')
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
}

function buildModelChain(options = {}) {
  const fromOption = parseModelChain(options.modelChain);
  const fromEnv = parseModelChain(process.env.GEMINI_MODEL_CHAIN);
  const baseChain = fromOption.length > 0 ? fromOption : fromEnv.length > 0 ? fromEnv : DEFAULT_MODEL_CHAIN;
  const preferred = String(options.model || process.env.GEMINI_MODEL || '').trim();
  const chain = preferred ? [preferred, ...baseChain] : [...baseChain];

  const unique = [];
  for (const model of chain) {
    if (!model || unique.includes(model)) {
      continue;
    }
    unique.push(model);
  }
  return unique;
}

function errorText(error) {
  return [
    error?.code,
    error?.status,
    error?.error?.status,
    error?.message,
    error?.error?.message,
    error?.details,
    error?.error?.details,
  ]
    .filter(Boolean)
    .map((value) => String(value).toLowerCase())
    .join(' | ');
}

function isRateLimitError(error) {
  const text = errorText(error);
  return (
    text.includes('429') ||
    text.includes('resource_exhausted') ||
    text.includes('rate limit') ||
    text.includes('quota')
  );
}

function isAuthError(error) {
  const text = errorText(error);
  return (
    text.includes('401') ||
    text.includes('unauthenticated') ||
    text.includes('invalid api key') ||
    text.includes('api key not valid')
  );
}

function isJsonModeUnsupportedError(error) {
  const text = errorText(error);
  return (
    (text.includes('invalid_argument') || text.includes('400') || text.includes('bad request')) &&
    (text.includes('responsemimetype') ||
      text.includes('response mime') ||
      text.includes('application/json') ||
      text.includes('json mode is not enabled') ||
      text.includes('json mode'))
  );
}

function isFallbackCandidateError(error) {
  if (isRateLimitError(error) || isAuthError(error)) {
    return true;
  }

  const text = errorText(error);
  return (
    text.includes('404') ||
    text.includes('model not found') ||
    text.includes('unsupported model') ||
    text.includes('not found') ||
    text.includes('403') ||
    text.includes('permission_denied') ||
    text.includes('503') ||
    text.includes('500') ||
    text.includes('unavailable') ||
    text.includes('overloaded') ||
    text.includes('internal') ||
    text.includes('timeout') ||
    text.includes('timed out') ||
    text.includes('fetch failed') ||
    text.includes('network')
  );
}

function normalizeOptionLabels(value) {
  if (Array.isArray(value)) {
    const labels = [];
    for (const item of value) {
      const label = String(item || '').trim().toUpperCase();
      if (/^[A-F]$/.test(label) && !labels.includes(label)) {
        labels.push(label);
      }
    }
    return labels;
  }

  const single = String(value || '').trim().toUpperCase();
  return /^[A-F]$/.test(single) ? [single] : [];
}

function knownCorrectLabels(question) {
  const byArray = normalizeOptionLabels(question.correctOptions);
  if (byArray.length > 0) {
    return byArray;
  }
  return normalizeOptionLabels(question.correctOption);
}

function hasSufficientExplanations(question) {
  if (knownCorrectLabels(question).length === 0) {
    return false;
  }

  if (!question.options || question.options.length < 2) {
    return false;
  }

  for (const option of question.options) {
    const explanation = question.explanations?.[option.label] || '';
    const normalized = explanation.trim();
    if (normalized.length < 25) {
      return false;
    }

    if (isPlaceholderExplanationText(explanation)) {
      return false;
    }
  }

  return true;
}

function buildPromptPayload(question) {
  return {
    number: question.number,
    stem: question.stem,
    options: question.options,
    knownCorrectOptions: knownCorrectLabels(question),
    knownCorrectOption: question.correctOption || null,
    knownExplanationForCorrect: question.sourceExplanation || null,
  };
}

function buildPrompt(chunk) {
  const payload = chunk.map(buildPromptPayload);

  return [
    'You are helping build an NBME-style medical practice quiz.',
    'For each question, provide concise teaching explanations for every option.',
    'Only infer answer choice labels when none are supplied in knownCorrectOptions/knownCorrectOption.',
    'Return ONLY JSON as an array with this exact shape:',
    '[{"number":1,"correctOption":"A","correctOptions":["A"],"explanations":{"A":"...","B":"...","C":"...","D":"..."}}]',
    'Rules:',
    '- Use only option labels provided.',
    '- Keep each explanation practical and educational (1-3 sentences).',
    '- Explain why the correct option is right and why each incorrect option is wrong.',
    '- If knownCorrectOptions/knownCorrectOption is provided, preserve it.',
    '',
    JSON.stringify(payload),
  ].join('\n');
}

function mergeAiResultIntoQuestion(question, generated) {
  if (!generated || typeof generated !== 'object') {
    return;
  }

  const hasExistingCorrect = knownCorrectLabels(question).length > 0;

  if (!hasExistingCorrect) {
    const generatedCorrectOptions = normalizeOptionLabels(generated.correctOptions);
    const generatedCorrectOption = normalizeOptionLabels(generated.correctOption);
    const resolvedCorrect = generatedCorrectOptions.length > 0 ? generatedCorrectOptions : generatedCorrectOption;
    if (resolvedCorrect.length > 0) {
      question.correctOptions = resolvedCorrect;
      question.correctOption = resolvedCorrect[0];
    }
  }

  const byOption = generated.explanations;
  if (!byOption || typeof byOption !== 'object') {
    return;
  }

  let addedAny = false;

  for (const option of question.options) {
    const next = byOption[option.label];
    if (!next || typeof next !== 'string') {
      continue;
    }

    const current = question.explanations[option.label] || '';
    if (current.trim().length >= 25 && !isPlaceholderExplanationText(current)) {
      continue;
    }

    question.explanations[option.label] = next.trim();
    addedAny = true;
  }

  if (addedAny) {
    if (question.explanationSource === 'document') {
      question.explanationSource = 'mixed';
    } else {
      question.explanationSource = 'gemini';
    }
  }
}

async function generateChunk(ai, model, chunk, timeoutMs = 12000) {
  const prompt = buildPrompt(chunk);
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  let response;
  try {
    try {
      response = await ai.models.generateContent({
        model,
        contents: prompt,
        config: {
          responseMimeType: 'application/json',
          temperature: 0.2,
          abortSignal: controller.signal,
        },
      });
    } catch (jsonModeError) {
      if (!isJsonModeUnsupportedError(jsonModeError)) {
        throw jsonModeError;
      }

      response = await ai.models.generateContent({
        model,
        contents: prompt,
        config: {
          temperature: 0.2,
          abortSignal: controller.signal,
        },
      });
    }
  } catch (error) {
    if (controller.signal.aborted) {
      const timeoutError = new Error(`Gemini chunk timed out after ${timeoutMs}ms.`);
      timeoutError.code = 'GEMINI_TIMEOUT';
      throw timeoutError;
    }
    if (isRateLimitError(error)) {
      const rateLimitError = new Error('Gemini rate limit reached for current model.');
      rateLimitError.code = 'GEMINI_RATE_LIMIT';
      throw rateLimitError;
    }
    if (isAuthError(error)) {
      const authError = new Error('Gemini authentication failed.');
      authError.code = 'GEMINI_AUTH';
      throw authError;
    }
    if (isFallbackCandidateError(error)) {
      const modelError = new Error('Gemini model call failed.');
      modelError.code = 'GEMINI_MODEL_ERROR';
      throw modelError;
    }
    throw error;
  } finally {
    clearTimeout(timer);
  }

  const parsed = extractJson(response.text);

  if (!Array.isArray(parsed)) {
    const outputError = new Error('Gemini returned non-JSON output.');
    outputError.code = 'GEMINI_MODEL_ERROR';
    throw outputError;
  }

  return parsed;
}

async function enrichQuestionsWithGemini(questions, options = {}) {
  const apiKey = options.apiKey || process.env.GEMINI_API_KEY;
  const models = buildModelChain(options);
  const initialModel = models[0] || 'gemma-3-27b-it';
  const chunkSize = Number.parseInt(String(options.chunkSize || process.env.GEMINI_CHUNK_SIZE || '3'), 10);
  const perChunkTimeoutMs = Number.parseInt(
    String(options.perChunkTimeoutMs || process.env.GEMINI_CHUNK_TIMEOUT_MS || '12000'),
    10,
  );
  const maxMilliseconds = Number.parseInt(
    String(options.maxMilliseconds || process.env.GEMINI_MAX_MS || '35000'),
    10,
  );
  const maxQuestions = Number.parseInt(
    String(options.maxQuestions || process.env.GEMINI_MAX_QUESTIONS || '40'),
    10,
  );

  if (!apiKey) {
    return {
      attempted: false,
      updatedQuestions: 0,
      reason: 'GEMINI_API_KEY is not configured.',
    };
  }

  const targets = questions.filter((q) => !hasSufficientExplanations(q));
  if (targets.length === 0) {
    return {
      attempted: false,
      updatedQuestions: 0,
      reason: 'All questions already contain sufficient explanations.',
    };
  }

  const limitedTargets = Number.isInteger(maxQuestions) && maxQuestions > 0 ? targets.slice(0, maxQuestions) : targets;
  const skippedQuestions = Math.max(0, targets.length - limitedTargets.length);

  const ai = new GoogleGenAI({ apiKey });
  const pendingChunks = chunkArray(limitedTargets, Math.max(1, chunkSize));

  let updatedQuestions = 0;
  let failures = 0;
  let processedQuestions = 0;
  let timedOut = false;
  let exhaustedModels = false;
  let authFailed = false;
  let modelIndex = 0;
  let activeModel = initialModel;
  let fallbackCount = 0;
  let rateLimitFallbackCount = 0;
  let modelErrorFallbackCount = 0;
  const triedModels = [activeModel];
  const startedAt = Date.now();

  while (pendingChunks.length > 0) {
    const chunk = pendingChunks.shift();
    if (!chunk || chunk.length === 0) {
      continue;
    }

    while (true) {
      activeModel = models[Math.min(modelIndex, models.length - 1)] || initialModel;

      const elapsedMs = Date.now() - startedAt;
      const remainingMs = maxMilliseconds - elapsedMs;

      if (remainingMs <= 1500) {
        timedOut = true;
        pendingChunks.length = 0;
        break;
      }

      const timeoutMs = Math.max(1500, Math.min(perChunkTimeoutMs, remainingMs - 250));

      try {
        const generatedItems = await generateChunk(ai, activeModel, chunk, timeoutMs);
        const generatedByNumber = new Map();

        for (const item of generatedItems) {
          if (!item || typeof item.number !== 'number') {
            continue;
          }
          generatedByNumber.set(item.number, item);
        }

        for (const question of chunk) {
          const before = JSON.stringify(question.explanations);
          mergeAiResultIntoQuestion(question, generatedByNumber.get(question.number));
          const after = JSON.stringify(question.explanations);
          if (before !== after) {
            updatedQuestions += 1;
          }
        }

        processedQuestions += chunk.length;
        break;
      } catch (error) {
        if (error && error.code === 'GEMINI_TIMEOUT') {
          if (chunk.length > 1) {
            const mid = Math.ceil(chunk.length / 2);
            pendingChunks.unshift(chunk.slice(mid));
            pendingChunks.unshift(chunk.slice(0, mid));
            break;
          }

          failures += 1;
          timedOut = true;
          break;
        }

        const canFallback = modelIndex < models.length - 1;
        const shouldFallback =
          error &&
          (error.code === 'GEMINI_RATE_LIMIT' ||
            error.code === 'GEMINI_MODEL_ERROR' ||
            (error.code !== 'GEMINI_AUTH' && isFallbackCandidateError(error)));

        if (shouldFallback && canFallback) {
          if (error.code === 'GEMINI_RATE_LIMIT') {
            rateLimitFallbackCount += 1;
          } else {
            modelErrorFallbackCount += 1;
          }
          modelIndex += 1;
          fallbackCount += 1;
          const nextModel = models[modelIndex];
          if (nextModel && !triedModels.includes(nextModel)) {
            triedModels.push(nextModel);
          }
          continue;
        }

        if (error && error.code === 'GEMINI_AUTH') {
          failures += 1;
          authFailed = true;
          pendingChunks.length = 0;
          break;
        }

        if (shouldFallback && !canFallback) {
          if (chunk.length > 1 && error.code !== 'GEMINI_RATE_LIMIT') {
            const mid = Math.ceil(chunk.length / 2);
            pendingChunks.unshift(chunk.slice(mid));
            pendingChunks.unshift(chunk.slice(0, mid));
            break;
          }

          failures += 1;
          exhaustedModels = true;
          pendingChunks.length = 0;
          break;
        }

        failures += 1;
        break;
      }
    }
  }

  const reasonParts = [];
  if (skippedQuestions > 0) {
    reasonParts.push(`Limited to first ${limitedTargets.length} question(s) to fit runtime.`);
  }
  if (timedOut) {
    reasonParts.push('Stopped early due to runtime budget.');
  }
  if (fallbackCount > 0) {
    reasonParts.push(`Fallback switched models ${fallbackCount} time(s); active model: ${activeModel}.`);
  }
  if (rateLimitFallbackCount > 0) {
    reasonParts.push(`Rate-limit fallbacks: ${rateLimitFallbackCount}.`);
  }
  if (modelErrorFallbackCount > 0) {
    reasonParts.push(`Model-error fallbacks: ${modelErrorFallbackCount}.`);
  }
  if (authFailed) {
    reasonParts.push('Gemini authentication failed.');
  }
  if (exhaustedModels) {
    reasonParts.push('All fallback models failed or were unavailable.');
  }
  if (failures > 0) {
    reasonParts.push(`${failures} chunk(s) failed.`);
  }

  return {
    attempted: true,
    model: activeModel,
    modelChain: models,
    triedModels,
    rateLimitFallbacks: rateLimitFallbackCount,
    rateLimitFallbackCount,
    modelErrorFallbackCount,
    updatedQuestions,
    failedChunks: failures,
    processedQuestions,
    skippedQuestions,
    timedOut,
    reason: reasonParts.join(' ') || '',
  };
}

module.exports = {
  enrichQuestionsWithGemini,
};

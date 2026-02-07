const { GoogleGenAI } = require('@google/genai');

const DEFAULT_MODEL_CHAIN = ['gemini-2.5-flash-lite', 'gemini-3.0-flash', 'gemini-2.5-flash', 'gemma-3-12b-it'];

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

function isRateLimitError(error) {
  const pieces = [
    error?.code,
    error?.status,
    error?.error?.status,
    error?.message,
    error?.error?.message,
    error?.details,
  ]
    .filter(Boolean)
    .map((value) => String(value).toLowerCase());

  return pieces.some((value) => {
    return (
      value.includes('429') ||
      value.includes('resource_exhausted') ||
      value.includes('rate limit') ||
      value.includes('quota')
    );
  });
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
    if (explanation.trim().length < 25) {
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
    if (current.trim().length >= 25) {
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
    response = await ai.models.generateContent({
      model,
      contents: prompt,
      config: {
        responseMimeType: 'application/json',
        temperature: 0.2,
        abortSignal: controller.signal,
      },
    });
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
    throw error;
  } finally {
    clearTimeout(timer);
  }

  const parsed = extractJson(response.text);

  if (!Array.isArray(parsed)) {
    throw new Error('Gemini returned non-JSON output.');
  }

  return parsed;
}

async function enrichQuestionsWithGemini(questions, options = {}) {
  const apiKey = options.apiKey || process.env.GEMINI_API_KEY;
  const models = buildModelChain(options);
  const initialModel = models[0] || 'gemini-2.5-flash-lite';
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
  const chunks = chunkArray(limitedTargets, Math.max(1, chunkSize));

  let updatedQuestions = 0;
  let failures = 0;
  let processedQuestions = 0;
  let timedOut = false;
  let exhaustedModels = false;
  let modelIndex = 0;
  let activeModel = initialModel;
  let fallbackCount = 0;
  const triedModels = [activeModel];
  const startedAt = Date.now();

  chunkLoop: for (const chunk of chunks) {
    while (true) {
      activeModel = models[Math.min(modelIndex, models.length - 1)] || initialModel;

      const elapsedMs = Date.now() - startedAt;
      const remainingMs = maxMilliseconds - elapsedMs;

      if (remainingMs <= 1500) {
        timedOut = true;
        break chunkLoop;
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
          failures += 1;
          timedOut = true;
          break chunkLoop;
        }

        if (error && error.code === 'GEMINI_RATE_LIMIT' && modelIndex < models.length - 1) {
          modelIndex += 1;
          fallbackCount += 1;
          const nextModel = models[modelIndex];
          if (nextModel && !triedModels.includes(nextModel)) {
            triedModels.push(nextModel);
          }
          continue;
        }

        if (error && error.code === 'GEMINI_RATE_LIMIT' && modelIndex >= models.length - 1) {
          failures += 1;
          exhaustedModels = true;
          break chunkLoop;
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
    reasonParts.push(`Rate-limit fallback used ${fallbackCount} time(s); active model: ${activeModel}.`);
  }
  if (exhaustedModels) {
    reasonParts.push('All fallback models were rate-limited.');
  }
  if (failures > 0) {
    reasonParts.push(`${failures} chunk(s) failed.`);
  }

  return {
    attempted: true,
    model: activeModel,
    modelChain: models,
    triedModels,
    rateLimitFallbacks: fallbackCount,
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

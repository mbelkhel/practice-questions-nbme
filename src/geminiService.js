const { GoogleGenAI } = require('@google/genai');

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

function hasSufficientExplanations(question) {
  if (!question.correctOption) {
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
    knownCorrectOption: question.correctOption || null,
    knownExplanationForCorrect: question.sourceExplanation || null,
  };
}

function buildPrompt(chunk) {
  const payload = chunk.map(buildPromptPayload);

  return [
    'You are helping build an NBME-style medical practice quiz.',
    'For each question, identify the best answer choice and provide concise teaching explanations for every option.',
    'Return ONLY JSON as an array with this exact shape:',
    '[{"number":1,"correctOption":"A","explanations":{"A":"...","B":"...","C":"...","D":"..."}}]',
    'Rules:',
    '- Use only option labels provided.',
    '- Keep each explanation practical and educational (1-3 sentences).',
    '- Explain why the correct option is right and why each incorrect option is wrong.',
    '- If knownCorrectOption is provided, keep it unless clearly impossible from the stem.',
    '',
    JSON.stringify(payload),
  ].join('\n');
}

function mergeAiResultIntoQuestion(question, generated) {
  if (!generated || typeof generated !== 'object') {
    return;
  }

  if (generated.correctOption && /^[A-F]$/.test(generated.correctOption)) {
    question.correctOption = generated.correctOption;
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

async function generateChunk(ai, model, chunk) {
  const prompt = buildPrompt(chunk);

  const response = await ai.models.generateContent({
    model,
    contents: prompt,
  });

  const parsed = extractJson(response.text);

  if (!Array.isArray(parsed)) {
    throw new Error('Gemini returned non-JSON output.');
  }

  return parsed;
}

async function enrichQuestionsWithGemini(questions, options = {}) {
  const apiKey = options.apiKey || process.env.GEMINI_API_KEY;
  const model = options.model || process.env.GEMINI_MODEL || 'gemini-2.5-flash-lite';

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

  const ai = new GoogleGenAI({ apiKey });
  const chunks = chunkArray(targets, 8);

  let updatedQuestions = 0;
  let failures = 0;

  for (const chunk of chunks) {
    try {
      const generatedItems = await generateChunk(ai, model, chunk);
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
    } catch (error) {
      failures += 1;
    }
  }

  return {
    attempted: true,
    model,
    updatedQuestions,
    failedChunks: failures,
  };
}

module.exports = {
  enrichQuestionsWithGemini,
};

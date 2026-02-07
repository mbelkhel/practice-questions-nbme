const { normalizeWhitespace } = require('./fileTextExtractor');

const OPTION_LABELS = ['A', 'B', 'C', 'D', 'E', 'F'];
const IMAGE_PATH_REGEX = /\.(?:png|jpe?g|gif|webp|bmp|svg)(?:\?[^)\s]*)?$/i;

function isOptionLine(line) {
  return /^\s*[A-F][\).:]\s+/.test(line);
}

function maybePushImage(images, rawSource) {
  if (!rawSource) {
    return;
  }

  const source = String(rawSource).trim().replace(/[),.;]+$/, '');
  if (!source) {
    return;
  }

  const isDataUrl = /^data:image\//i.test(source);
  const isHttp = /^https?:\/\//i.test(source) && IMAGE_PATH_REGEX.test(source);
  const isRelativePath = /^(?:\/|\.\/|\.\.\/)/.test(source) && IMAGE_PATH_REGEX.test(source);
  const isFileName = /^[^\s/]+?\.(?:png|jpe?g|gif|webp|bmp|svg)$/i.test(source);

  if (!isDataUrl && !isHttp && !isRelativePath && !isFileName) {
    return;
  }

  if (!images.includes(source)) {
    images.push(source);
  }
}

function extractImageRefs(text) {
  const images = [];
  let cleaned = String(text || '');

  cleaned = cleaned.replace(/!\[[^\]]*]\(([^)]+)\)/gi, (match, src) => {
    maybePushImage(images, src);
    return '';
  });

  cleaned = cleaned.replace(/\[IMAGE:([^\]]+)\]/gi, (match, src) => {
    maybePushImage(images, src);
    return '';
  });

  cleaned = cleaned.replace(/(?:^|\n)\s*(?:image|figure)\s*[:\-]\s*([^\n]+)/gi, (match, srcLine) => {
    const candidate = srcLine.trim().split(/\s+/)[0];
    maybePushImage(images, candidate);
    return '\n';
  });

  cleaned = cleaned.replace(
    /(?:^|\n)\s*((?:https?:\/\/|\/|\.\/|\.\.\/)?[^\s]+\.(?:png|jpe?g|gif|webp|bmp|svg)(?:\?[^\s]+)?)\s*(?=\n|$)/gi,
    (match, src) => {
      maybePushImage(images, src);
      return '\n';
    },
  );

  return {
    text: normalizeWhitespace(cleaned),
    images,
  };
}

function parseInlineOptions(blockText) {
  const optionRegex = /\(([A-F])\)\s*([\s\S]*?)(?=\([A-F]\)|$)/g;
  const matches = [...blockText.matchAll(optionRegex)];

  if (matches.length < 2) {
    return { stem: blockText.trim(), options: [] };
  }

  const firstMatch = matches[0];
  const stem = blockText.slice(0, firstMatch.index).trim();
  const options = matches.map((match) => ({
    label: match[1],
    text: normalizeWhitespace(match[2]),
  }));

  return { stem, options };
}

function parseUnlabeledTrailingOptions(lines) {
  const nonEmpty = lines.map((line) => line.trim()).filter(Boolean);

  if (nonEmpty.length < 3) {
    return { stem: nonEmpty.join(' '), options: [] };
  }

  for (let i = nonEmpty.length - 2; i >= 0; i -= 1) {
    const stemCandidate = nonEmpty.slice(0, i + 1);
    const optionCandidate = nonEmpty.slice(i + 1);

    const stemTail = stemCandidate[stemCandidate.length - 1] || '';
    const validOptionCount = optionCandidate.length >= 2 && optionCandidate.length <= 6;
    const optionShapeLooksRight = optionCandidate.every((line) => line.length <= 140 && !/[?]$/.test(line));

    if (validOptionCount && optionShapeLooksRight && /[:?]$/.test(stemTail)) {
      return {
        stem: normalizeWhitespace(stemCandidate.join('\n')),
        options: optionCandidate.map((text, idx) => ({
          label: OPTION_LABELS[idx],
          text: normalizeWhitespace(text),
        })),
      };
    }
  }

  return { stem: normalizeWhitespace(nonEmpty.join('\n')), options: [] };
}

function parseQuestionBlock(number, blockText) {
  const text = blockText.trim();
  const lines = text.split('\n');

  let firstOptionLine = -1;
  for (let i = 0; i < lines.length; i += 1) {
    if (isOptionLine(lines[i])) {
      firstOptionLine = i;
      break;
    }
  }

  let stem = text;
  let options = [];
  let embeddedAnswer = null;

  if (firstOptionLine >= 0) {
    const stemLines = lines.slice(0, firstOptionLine);
    const optionLines = lines.slice(firstOptionLine);

    stem = normalizeWhitespace(stemLines.join('\n'));

    let current = null;
    for (const rawLine of optionLines) {
      const line = rawLine.trim();
      const optionMatch = line.match(/^([A-F])[\).:]\s*(.*)$/);

      if (optionMatch) {
        if (current) {
          current.text = normalizeWhitespace(current.text);
          options.push(current);
        }

        current = {
          label: optionMatch[1],
          text: optionMatch[2] || '',
        };

        continue;
      }

      const answerLine = line.match(/^(?:correct\s*answer|answer)\s*[:\-]\s*([A-F])/i);
      if (answerLine) {
        embeddedAnswer = answerLine[1].toUpperCase();
        continue;
      }

      if (current) {
        current.text = `${current.text} ${line}`.trim();
      } else {
        stem = `${stem} ${line}`.trim();
      }
    }

    if (current) {
      current.text = normalizeWhitespace(current.text);
      options.push(current);
    }
  }

  if (options.length < 2) {
    const inlineParsed = parseInlineOptions(text);
    stem = inlineParsed.stem;
    options = inlineParsed.options;
  }

  if (options.length < 2) {
    const unlabeledParsed = parseUnlabeledTrailingOptions(lines);
    stem = unlabeledParsed.stem;
    options = unlabeledParsed.options;
  }

  const answerInBlock = text.match(/(?:correct\s*answer|answer)\s*[:\-]?\s*([A-F])/i);
  if (!embeddedAnswer && answerInBlock) {
    embeddedAnswer = answerInBlock[1].toUpperCase();
  }

  const extractedStem = extractImageRefs(stem);
  const imageRefs = [...extractedStem.images];
  stem = extractedStem.text;

  options = options.map((option) => {
    const extractedOption = extractImageRefs(option.text);
    for (const source of extractedOption.images) {
      maybePushImage(imageRefs, source);
    }

    return {
      ...option,
      text: extractedOption.text,
    };
  });

  return {
    id: `q-${number}`,
    number,
    stem: normalizeWhitespace(stem),
    stemHtml: null,
    images: imageRefs,
    options,
    correctOption: embeddedAnswer,
    explanations: {},
    sourceExplanation: '',
    explanationSource: 'none',
  };
}

function splitQuestionAndAnswerSections(rawText) {
  const text = normalizeWhitespace(rawText);
  const lower = text.toLowerCase();

  const headingPatterns = [
    /\n\s*answer key\b/g,
    /\n\s*answers and explanations\b/g,
    /\n\s*answers\b/g,
    /\n\s*explanations\b/g,
    /\n\s*rationales\b/g,
  ];

  let answerStart = -1;
  for (const pattern of headingPatterns) {
    let match = pattern.exec(lower);
    while (match) {
      const idx = match.index;
      if (idx > text.length * 0.2) {
        if (answerStart === -1 || idx < answerStart) {
          answerStart = idx;
        }
      }
      match = pattern.exec(lower);
    }
  }

  if (answerStart === -1) {
    const lines = text.split('\n');
    for (let i = Math.floor(lines.length * 0.35); i < lines.length; i += 1) {
      let answerLikeLines = 0;
      for (let j = i; j < Math.min(i + 8, lines.length); j += 1) {
        if (/^\s*(?:q(?:uestion)?\s*)?\d{1,3}\s*[\).:-]?\s*(?:answer\s*[:\-]\s*)?[A-F]\b/i.test(lines[j])) {
          answerLikeLines += 1;
        }
      }
      if (answerLikeLines >= 4) {
        answerStart = lines.slice(0, i).join('\n').length;
        break;
      }
    }
  }

  if (answerStart === -1) {
    return {
      questionSection: text,
      answerSection: '',
    };
  }

  return {
    questionSection: text.slice(0, answerStart).trim(),
    answerSection: text.slice(answerStart).trim(),
  };
}

function parseQuestions(questionSection) {
  const questions = [];
  const explicitRegex = /(?:^|\n)\s*(?:question\s*|q\s*)(\d{1,3})\s*[\).:-]?\s*\n+([\s\S]*?)(?=\n\s*(?:question\s*|q\s*)\d{1,3}\s*[\).:-]?\s*\n+|$)/gi;
  const numericRegex = /(?:^|\n)\s*(\d{1,3})\s*[\).]\s+([\s\S]*?)(?=\n\s*\d{1,3}\s*[\).]\s+|$)/g;

  let matches = [...questionSection.matchAll(explicitRegex)];
  if (matches.length === 0) {
    matches = [...questionSection.matchAll(numericRegex)];
  }

  for (const match of matches) {
    const number = Number.parseInt(match[1], 10);
    const blockText = normalizeWhitespace(match[2]);
    const parsed = parseQuestionBlock(number, blockText);

    if (parsed.stem && parsed.options.length >= 2) {
      questions.push(parsed);
    }
  }

  return questions;
}

function parseAnswerSection(answerSection) {
  const answerMap = new Map();

  if (!answerSection) {
    return answerMap;
  }

  const simpleKeyRegex = /(?:^|\n)\s*(?:q(?:uestion)?\s*)?(\d{1,3})\s*[\).:-]\s*([A-F])\b/gi;
  for (const match of answerSection.matchAll(simpleKeyRegex)) {
    const number = Number.parseInt(match[1], 10);
    const option = match[2].toUpperCase();

    if (!answerMap.has(number)) {
      answerMap.set(number, {
        correctOption: option,
        explanation: '',
      });
    } else if (!answerMap.get(number).correctOption) {
      answerMap.get(number).correctOption = option;
    }
  }

  const blockRegex = /(?:^|\n)\s*(?:q(?:uestion)?\s*)?(\d{1,3})\s*[\).:-]?\s*(?:answer\s*[:\-]?\s*)?([\s\S]*?)(?=\n\s*(?:q(?:uestion)?\s*)?\d{1,3}\s*[\).:-]?\s*(?:answer\s*[:\-]?\s*)|$)/gi;
  for (const match of answerSection.matchAll(blockRegex)) {
    const number = Number.parseInt(match[1], 10);
    const rawBody = normalizeWhitespace(match[2]);

    if (!rawBody || rawBody.length < 2) {
      continue;
    }

    let correctOption = null;
    let explanation = rawBody;

    const answerPrefixed = explanation.match(/(?:^|\b)(?:answer|ans|correct\s*answer)\s*[:\-]?\s*([A-F])\b/i);
    if (answerPrefixed) {
      correctOption = answerPrefixed[1].toUpperCase();
    } else {
      const leadingOption = explanation.match(/^\(?([A-F])\)?[\).:\-]/i);
      if (leadingOption) {
        correctOption = leadingOption[1].toUpperCase();
      }
    }

    const explanationOnly = explanation.match(/explanation\s*:\s*([\s\S]*)/i);
    if (explanationOnly) {
      explanation = explanationOnly[1].trim();
    } else {
      explanation = explanation
        .replace(/^(?:answer|ans|correct\s*answer)\s*[:\-]?\s*/i, '')
        .replace(/^\(?[A-F]\)?[\).:\-]\s*/i, '')
        .trim();
    }

    const existing = answerMap.get(number) || {
      correctOption: null,
      explanation: '',
    };

    if (correctOption) {
      existing.correctOption = correctOption;
    }

    if (explanation && explanation.length > existing.explanation.length) {
      existing.explanation = explanation;
    }

    answerMap.set(number, existing);
  }

  return answerMap;
}

function applyAnswersToQuestions(questions, answerMap) {
  for (const question of questions) {
    const answer = answerMap.get(question.number);

    if (!answer) {
      continue;
    }

    if (answer.correctOption) {
      question.correctOption = answer.correctOption;
    }

    if (answer.explanation) {
      question.sourceExplanation = answer.explanation;
      question.explanationSource = 'document';

      if (question.correctOption) {
        question.explanations[question.correctOption] = answer.explanation;
      }
    }
  }
}

function buildQuizFromText(rawText) {
  const { questionSection, answerSection } = splitQuestionAndAnswerSections(rawText);
  const questions = parseQuestions(questionSection);

  const answerMap = parseAnswerSection(answerSection);
  applyAnswersToQuestions(questions, answerMap);

  const inferredTitle = questionSection
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
    .slice(0, 3)
    .find((line) => !/^(question|q)\s*\d+/i.test(line));

  return {
    title: (inferredTitle || 'Generated Quiz').slice(0, 120),
    questions,
    parsing: {
      detectedAnswerSection: Boolean(answerSection),
      totalQuestions: questions.length,
      answersMapped: questions.filter((q) => q.correctOption).length,
      explanationsMapped: questions.filter((q) => Object.keys(q.explanations).length > 0).length,
    },
  };
}

module.exports = {
  buildQuizFromText,
};

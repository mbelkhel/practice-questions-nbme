const { normalizeWhitespace } = require('./fileTextExtractor');

const OPTION_LABELS = ['A', 'B', 'C', 'D', 'E', 'F'];
const IMAGE_PATH_REGEX = /\.(?:png|jpe?g|gif|webp|bmp|svg)(?:\?[^)\s]*)?$/i;
const MULTI_SELECT_STEM_REGEX =
  /\b(select all that apply|choose all that apply|which .* are correct|all of the following are true)\b/i;

function isOptionLine(line) {
  return /^\s*[A-F][\).:]\s+/.test(line);
}

function normalizeOptionDisplayText(text) {
  return normalizeWhitespace(String(text || '').replace(/^[☐☑✓]\s*/u, '').trim());
}

function normalizeBooleanToken(text) {
  const normalized = normalizeWhitespace(String(text || '').replace(/^[☐☑✓]\s*/u, ''));
  if (/^true$/i.test(normalized)) {
    return 'TRUE';
  }
  if (/^false$/i.test(normalized)) {
    return 'FALSE';
  }
  return '';
}

function isTrueFalseQuestionByOptions(question) {
  if (!question || !Array.isArray(question.options) || question.options.length !== 2) {
    return false;
  }

  const first = normalizeBooleanToken(question.options[0].text);
  const second = normalizeBooleanToken(question.options[1].text);
  return Boolean(first && second && first !== second);
}

function pickOptionForBoolean(question, boolValue) {
  if (!question || !Array.isArray(question.options)) {
    return null;
  }

  const target = String(boolValue || '').toUpperCase();
  if (!target) {
    return null;
  }

  const found = question.options.find((option) => normalizeBooleanToken(option.text) === target);
  return found ? found.label : null;
}

function splitLetterAnswerToken(token) {
  const cleaned = String(token || '').trim().toUpperCase().replace(/\s+AND\s+/g, ',');
  if (!/^[A-F](?:\s*[,/;&+]\s*[A-F])*$/i.test(cleaned)) {
    return [];
  }

  return cleaned
    .split(/[,/;&+]/)
    .map((part) => part.trim())
    .filter((part) => /^[A-F]$/.test(part));
}

function splitNumericAnswerToken(token) {
  const cleaned = String(token || '').trim();
  if (!/^\d(?:\s*,\s*\d)*$/.test(cleaned)) {
    return [];
  }

  return cleaned
    .split(',')
    .map((part) => Number.parseInt(part.trim(), 10))
    .filter((value) => Number.isInteger(value) && value >= 1 && value <= OPTION_LABELS.length)
    .map((value) => OPTION_LABELS[value - 1]);
}

function parseGeneralAnswerToken(rawToken) {
  const token = normalizeWhitespace(String(rawToken || '').replace(/^[\-•]\s*/, ''));
  if (!token) {
    return {
      correctOptions: [],
      booleanValue: null,
      rawToken: token,
    };
  }

  const boolToken = normalizeBooleanToken(token);
  if (boolToken) {
    return {
      correctOptions: [],
      booleanValue: boolToken,
      rawToken: token,
    };
  }

  const letterOptions = splitLetterAnswerToken(token);
  if (letterOptions.length) {
    return {
      correctOptions: letterOptions,
      booleanValue: null,
      rawToken: token,
    };
  }

  const numericOptions = splitNumericAnswerToken(token);
  if (numericOptions.length) {
    return {
      correctOptions: numericOptions,
      booleanValue: null,
      rawToken: token,
    };
  }

  return {
    correctOptions: [],
    booleanValue: null,
    rawToken: token,
  };
}

function looksLikeAnswerTokenLine(line) {
  const token = normalizeWhitespace(String(line || '').replace(/^[\-•]\s*/, ''));
  if (!token) {
    return false;
  }

  if (normalizeBooleanToken(token)) {
    return true;
  }

  if (splitLetterAnswerToken(token).length > 0) {
    return true;
  }

  if (splitNumericAnswerToken(token).length > 0) {
    return true;
  }

  return false;
}

function inferQuestionType(question) {
  if (Array.isArray(question.correctOptions) && question.correctOptions.length > 1) {
    return 'multi_select';
  }

  if (isTrueFalseQuestionByOptions(question)) {
    return 'true_false';
  }

  if (MULTI_SELECT_STEM_REGEX.test(question.stem || '')) {
    return 'multi_select';
  }

  return 'single_select';
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
    text: normalizeOptionDisplayText(match[2]),
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
    const validOptionCount = optionCandidate.length >= 2 && optionCandidate.length <= OPTION_LABELS.length;
    const optionShapeLooksRight = optionCandidate.every((line) => line.length <= 180 && !/[?]$/.test(line));

    if (validOptionCount && optionShapeLooksRight && /[:?]$/.test(stemTail)) {
      return {
        stem: normalizeWhitespace(stemCandidate.join('\n')),
        options: optionCandidate.map((text, idx) => ({
          label: OPTION_LABELS[idx],
          text: normalizeOptionDisplayText(text),
        })),
      };
    }
  }

  return { stem: normalizeWhitespace(nonEmpty.join('\n')), options: [] };
}

function createQuestionFromParts(number, stem, options, embeddedAnswer = null, type = 'single_select') {
  const extractedStem = extractImageRefs(stem);
  const imageRefs = [...extractedStem.images];
  const cleanStem = extractedStem.text;

  const cleanOptions = options.map((option) => {
    const extractedOption = extractImageRefs(option.text);
    for (const source of extractedOption.images) {
      maybePushImage(imageRefs, source);
    }

    return {
      ...option,
      text: normalizeOptionDisplayText(extractedOption.text),
    };
  });

  const cleanCorrectOption = embeddedAnswer && /^[A-F]$/.test(embeddedAnswer) ? embeddedAnswer : null;

  return {
    id: `q-${number}`,
    number,
    type,
    stem: normalizeWhitespace(cleanStem),
    stemHtml: null,
    images: imageRefs,
    options: cleanOptions,
    correctOption: cleanCorrectOption,
    correctOptions: cleanCorrectOption ? [cleanCorrectOption] : [],
    explanations: {},
    sourceExplanation: '',
    explanationSource: 'none',
    rawAnswerToken: '',
  };
}

function isImageMarkerLine(line) {
  return /^\[IMAGE:/i.test(line);
}

function isLikelyQuestionStartLine(line) {
  const trimmed = line.trim();
  if (!trimmed || isImageMarkerLine(trimmed)) {
    return false;
  }

  if (/^(?:question\s+\d+|q\d+)/i.test(trimmed)) {
    return true;
  }

  if (/^(?:a|an)\s+\d{1,3}-year-old\b/i.test(trimmed)) {
    return true;
  }

  if (/^(?:what|which|in the|next step|true or false|place the)\b/i.test(trimmed)) {
    return true;
  }

  if (/[?:]$/.test(trimmed) && trimmed.length >= 35) {
    return true;
  }

  return false;
}

function looksLikeOptionText(line) {
  const trimmed = line.trim();
  if (!trimmed || isImageMarkerLine(trimmed)) {
    return false;
  }

  if (/^answers?$/i.test(trimmed)) {
    return false;
  }

  if (isLikelyQuestionStartLine(trimmed)) {
    return false;
  }

  if (trimmed.length > 240) {
    return false;
  }

  return true;
}

function stripLeadingOptionMarker(text) {
  return String(text || '').replace(/^\(?[A-F]\)?[).:\-]\s+/i, '').trim();
}

function parseEmbeddedOptionFromStem(stemText) {
  const stem = normalizeWhitespace(stemText);
  if (!stem) {
    return null;
  }

  const colonIndex = Math.max(stem.lastIndexOf(':'), stem.lastIndexOf('：'));
  if (colonIndex < 0 || colonIndex >= stem.length - 8) {
    return null;
  }

  const before = normalizeWhitespace(stem.slice(0, colonIndex));
  const optionText = normalizeOptionDisplayText(stem.slice(colonIndex + 1));

  if (!before || !optionText) {
    return null;
  }

  if (optionText.length < 18 || optionText.length > 260) {
    return null;
  }

  if (/[?]$/.test(optionText) || /^[A-F][).:]/i.test(optionText)) {
    return null;
  }

  if (isLikelyQuestionStartLine(optionText)) {
    return null;
  }

  if (!/\s/.test(optionText)) {
    return null;
  }

  return {
    stem: before,
    optionText,
  };
}

function relabelSequentialOptions(options) {
  return options.slice(0, OPTION_LABELS.length).map((option, idx) => ({
    label: OPTION_LABELS[idx],
    text: normalizeOptionDisplayText(option.text),
  }));
}

function tryRecoverLeadingOptionFromStem(question, missingLabels) {
  if (!question || !Array.isArray(question.options)) {
    return false;
  }

  if (!Array.isArray(missingLabels) || missingLabels.length === 0) {
    return false;
  }

  if (question.options.length < 2 || question.options.length >= OPTION_LABELS.length) {
    return false;
  }

  const expectedMissingLabel = OPTION_LABELS[question.options.length];
  if (!missingLabels.includes(expectedMissingLabel)) {
    return false;
  }

  const embedded = parseEmbeddedOptionFromStem(question.stem);
  if (!embedded) {
    return false;
  }

  question.stem = embedded.stem;
  question.options = relabelSequentialOptions([
    { text: embedded.optionText },
    ...question.options.map((option) => ({ text: option.text })),
  ]);
  return true;
}

function expandTrueFalseClusterQuestion(stem, optionLines, numberStart) {
  if (!/^true\s*or\s*false$/i.test((stem || '').trim())) {
    return [];
  }

  const cleaned = optionLines
    .map((line) => normalizeOptionDisplayText(stripLeadingOptionMarker(line)))
    .filter(Boolean);

  const questions = [];
  let cursor = 0;
  let number = numberStart;

  while (cursor + 2 < cleaned.length) {
    const statement = cleaned[cursor];
    const firstBool = normalizeBooleanToken(cleaned[cursor + 1]);
    const secondBool = normalizeBooleanToken(cleaned[cursor + 2]);

    if (!statement || !firstBool || !secondBool || firstBool === secondBool) {
      break;
    }

    const options = [
      { label: 'A', text: firstBool === 'TRUE' ? 'True' : 'False' },
      { label: 'B', text: secondBool === 'TRUE' ? 'True' : 'False' },
    ];

    questions.push(createQuestionFromParts(number, statement, options, null, 'true_false'));
    number += 1;
    cursor += 3;
  }

  if (questions.length >= 2 && cursor >= cleaned.length) {
    return questions;
  }

  return [];
}

function parseUnnumberedQuestionBlocks(questionSection) {
  const rawLines = questionSection.split('\n').map((line) => line.trim()).filter(Boolean);
  if (!rawLines.length) {
    return [];
  }

  let lines = rawLines;
  if (rawLines.length > 1 && !isLikelyQuestionStartLine(rawLines[0]) && isLikelyQuestionStartLine(rawLines[1])) {
    lines = rawLines.slice(1);
  }

  const questions = [];
  let cursor = 0;
  let number = 1;
  let pendingImageLines = [];

  while (cursor < lines.length) {
    while (cursor < lines.length && !isLikelyQuestionStartLine(lines[cursor])) {
      if (isImageMarkerLine(lines[cursor])) {
        pendingImageLines.push(lines[cursor]);
      }
      cursor += 1;
    }

    if (cursor >= lines.length) {
      break;
    }

    const stemLines = [...pendingImageLines, lines[cursor]];
    pendingImageLines = [];
    cursor += 1;

    while (cursor < lines.length) {
      const nextLine = lines[cursor];

      if (isImageMarkerLine(nextLine)) {
        stemLines.push(nextLine);
        cursor += 1;
        continue;
      }

      if (looksLikeOptionText(nextLine)) {
        break;
      }

      if (isLikelyQuestionStartLine(nextLine) && stemLines.some((line) => /[?:]$/.test(line))) {
        break;
      }

      stemLines.push(nextLine);
      cursor += 1;
    }

    const optionLines = [];
    while (cursor < lines.length) {
      const nextLine = lines[cursor];

      if (isImageMarkerLine(nextLine)) {
        // Images encountered after answer choices usually belong to the following question.
        if (optionLines.length >= 2) {
          pendingImageLines.push(nextLine);
          cursor += 1;
          break;
        }
        stemLines.push(nextLine);
        cursor += 1;
        continue;
      }

      if (optionLines.length >= 2 && isLikelyQuestionStartLine(nextLine)) {
        break;
      }

      if (!looksLikeOptionText(nextLine)) {
        if (optionLines.length >= 2) {
          break;
        }
        cursor += 1;
        continue;
      }

      optionLines.push(nextLine);
      cursor += 1;

      if (optionLines.length >= 30) {
        break;
      }
    }

    if (optionLines.length >= 2) {
      const stem = normalizeWhitespace(stemLines.join('\n'));
      const expandedTrueFalse = expandTrueFalseClusterQuestion(stem, optionLines, number);
      if (expandedTrueFalse.length > 0) {
        questions.push(...expandedTrueFalse);
        number += expandedTrueFalse.length;
        continue;
      }

      const options = optionLines.slice(0, OPTION_LABELS.length).map((text, idx) => ({
        label: OPTION_LABELS[idx],
        text: normalizeOptionDisplayText(stripLeadingOptionMarker(text)),
      }));

      const question = createQuestionFromParts(number, stem, options, null);
      questions.push(question);
      number += 1;
      continue;
    }
  }

  return questions;
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
          current.text = normalizeOptionDisplayText(current.text);
          options.push(current);
        }

        current = {
          label: optionMatch[1],
          text: normalizeOptionDisplayText(optionMatch[2] || ''),
        };

        continue;
      }

      const answerLine = line.match(/^(?:correct\s*answer|answer)\s*[:\-]\s*([A-F])/i);
      if (answerLine) {
        embeddedAnswer = answerLine[1].toUpperCase();
        continue;
      }

      if (current) {
        current.text = normalizeOptionDisplayText(`${current.text} ${line}`.trim());
      } else {
        stem = `${stem} ${line}`.trim();
      }
    }

    if (current) {
      current.text = normalizeOptionDisplayText(current.text);
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

  return createQuestionFromParts(number, stem, options, embeddedAnswer);
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
  const explicitRegex =
    /(?:^|\n)\s*(?:question\s*|q\s*)(\d{1,3})\s*[\).:-]?\s*\n+([\s\S]*?)(?=\n\s*(?:question\s*|q\s*)\d{1,3}\s*[\).:-]?\s*\n+|$)/gi;
  const numericRegex = /(?:^|\n)\s*(\d{1,3})\s*[\).]\s+([\s\S]*?)(?=\n\s*\d{1,3}\s*[\).]\s+|$)/g;

  let matches = [...questionSection.matchAll(explicitRegex)];
  if (matches.length === 0) {
    matches = [...questionSection.matchAll(numericRegex)];
  }

  if (matches.length === 0) {
    return parseUnnumberedQuestionBlocks(questionSection);
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

function parseSequentialAnswerSection(answerSection, questions) {
  const answerMap = new Map();

  if (!answerSection || !questions.length) {
    return answerMap;
  }

  const lines = normalizeWhitespace(answerSection)
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean);

  const startIdx = lines.findIndex((line) => /^answers?$/i.test(line));
  const candidateLines = (startIdx >= 0 ? lines.slice(startIdx + 1) : lines).map((line) => line.trim());

  let questionIdx = 0;
  for (const line of candidateLines) {
    if (questionIdx >= questions.length) {
      break;
    }

    if (!looksLikeAnswerTokenLine(line)) {
      continue;
    }

    const question = questions[questionIdx];
    const parsed = parseGeneralAnswerToken(line);

    answerMap.set(question.number, {
      correctOption: parsed.correctOptions[0] || null,
      correctOptions: parsed.correctOptions,
      booleanValue: parsed.booleanValue,
      explanation: '',
      rawAnswerToken: parsed.rawToken,
    });

    questionIdx += 1;
  }

  return answerMap;
}

function parseAnswerSection(answerSection) {
  const answerMap = new Map();

  if (!answerSection) {
    return answerMap;
  }

  const simpleKeyRegex = /(?:^|\n)\s*(?:q(?:uestion)?\s*)?(\d{1,3})\s*[\).:-]\s*([^\n]+)/gi;
  for (const match of answerSection.matchAll(simpleKeyRegex)) {
    const number = Number.parseInt(match[1], 10);
    const parsed = parseGeneralAnswerToken(match[2]);

    const existing = answerMap.get(number) || {
      correctOption: null,
      correctOptions: [],
      booleanValue: null,
      explanation: '',
      rawAnswerToken: '',
    };

    if (existing.correctOptions.length === 0 && parsed.correctOptions.length > 0) {
      existing.correctOptions = parsed.correctOptions;
      existing.correctOption = parsed.correctOptions[0] || null;
    }

    if (!existing.booleanValue && parsed.booleanValue) {
      existing.booleanValue = parsed.booleanValue;
    }

    if (!existing.rawAnswerToken && parsed.rawToken) {
      existing.rawAnswerToken = parsed.rawToken;
    }

    answerMap.set(number, existing);
  }

  const blockRegex =
    /(?:^|\n)\s*(?:q(?:uestion)?\s*)?(\d{1,3})\s*[\).:-]?\s*(?:answer\s*[:\-]?\s*)?([\s\S]*?)(?=\n\s*(?:q(?:uestion)?\s*)?\d{1,3}\s*[\).:-]?\s*(?:answer\s*[:\-]?\s*)|$)/gi;
  for (const match of answerSection.matchAll(blockRegex)) {
    const number = Number.parseInt(match[1], 10);
    const rawBody = normalizeWhitespace(match[2]);

    if (!rawBody || rawBody.length < 2) {
      continue;
    }

    const parsed = parseGeneralAnswerToken(rawBody);
    let explanation = rawBody;

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
      correctOptions: [],
      booleanValue: null,
      explanation: '',
      rawAnswerToken: '',
    };

    if (existing.correctOptions.length === 0 && parsed.correctOptions.length > 0) {
      existing.correctOptions = parsed.correctOptions;
      existing.correctOption = parsed.correctOptions[0] || null;
    }

    if (!existing.booleanValue && parsed.booleanValue) {
      existing.booleanValue = parsed.booleanValue;
    }

    if (explanation && explanation.length > existing.explanation.length) {
      existing.explanation = explanation;
    }

    if (!existing.rawAnswerToken && parsed.rawToken) {
      existing.rawAnswerToken = parsed.rawToken;
    }

    answerMap.set(number, existing);
  }

  return answerMap;
}

function applyAnswersToQuestions(questions, answerMap) {
  for (const question of questions) {
    const answer = answerMap.get(question.number);

    if (answer) {
      const declaredAnswerLabels = [];
      if (Array.isArray(answer.correctOptions) && answer.correctOptions.length > 0) {
        for (const label of answer.correctOptions) {
          const normalized = String(label || '').toUpperCase();
          if (/^[A-F]$/.test(normalized)) {
            declaredAnswerLabels.push(normalized);
          }
        }
      } else if (answer.correctOption) {
        const normalized = String(answer.correctOption || '').toUpperCase();
        if (/^[A-F]$/.test(normalized)) {
          declaredAnswerLabels.push(normalized);
        }
      }

      if (declaredAnswerLabels.length > 0) {
        const availableLabels = new Set(question.options.map((option) => option.label));
        const missingLabels = declaredAnswerLabels.filter((label) => !availableLabels.has(label));
        if (missingLabels.length > 0) {
          tryRecoverLeadingOptionFromStem(question, missingLabels);
        }
      }

      let resolvedCorrectOptions = [];

      if (Array.isArray(answer.correctOptions) && answer.correctOptions.length > 0) {
        resolvedCorrectOptions = answer.correctOptions
          .map((label) => String(label || '').toUpperCase())
          .filter((label) => question.options.some((option) => option.label === label));
      } else if (answer.correctOption) {
        const label = String(answer.correctOption || '').toUpperCase();
        if (question.options.some((option) => option.label === label)) {
          resolvedCorrectOptions = [label];
        }
      } else if (answer.booleanValue) {
        const mapped = pickOptionForBoolean(question, answer.booleanValue);
        if (mapped) {
          resolvedCorrectOptions = [mapped];
        }
      }

      question.correctOptions = resolvedCorrectOptions;
      question.correctOption = resolvedCorrectOptions[0] || null;
      question.rawAnswerToken = answer.rawAnswerToken || '';

      if (answer.explanation) {
        question.sourceExplanation = answer.explanation;
        question.explanationSource = 'document';

        if (question.correctOption) {
          question.explanations[question.correctOption] = answer.explanation;
        }
      }
    }

    question.type = inferQuestionType(question);
  }
}

function buildQuizFromText(rawText) {
  const { questionSection, answerSection } = splitQuestionAndAnswerSections(rawText);
  const questions = parseQuestions(questionSection);

  const answerMap = parseAnswerSection(answerSection);
  const hasDirectMappedAnswers = [...answerMap.values()].some((answer) => {
    return (
      (Array.isArray(answer.correctOptions) && answer.correctOptions.length > 0) ||
      Boolean(answer.correctOption) ||
      Boolean(answer.booleanValue)
    );
  });

  if (!hasDirectMappedAnswers) {
    answerMap.clear();
    const sequentialAnswerMap = parseSequentialAnswerSection(answerSection, questions);
    for (const [number, answer] of sequentialAnswerMap.entries()) {
      answerMap.set(number, answer);
    }
  }

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
      answersMapped: questions.filter((q) => Array.isArray(q.correctOptions) && q.correctOptions.length > 0).length,
      explanationsMapped: questions.filter((q) => Object.keys(q.explanations).length > 0).length,
    },
  };
}

module.exports = {
  buildQuizFromText,
};

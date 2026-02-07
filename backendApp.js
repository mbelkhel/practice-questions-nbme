require('dotenv').config();

const fs = require('fs/promises');
const fsSync = require('fs');
const os = require('os');
const path = require('path');
const express = require('express');
const multer = require('multer');

const { extractTextFromDocument } = require('./src/fileTextExtractor');
const { buildQuizFromText } = require('./src/quizParser');
const { enrichQuestionsWithGemini } = require('./src/geminiService');

const app = express();
const PORT = Number.parseInt(process.env.PORT || '3000', 10);

const runningOnVercel = Boolean(process.env.VERCEL || process.env.VERCEL_ENV || process.env.NOW_REGION);
const uploadDir = runningOnVercel ? path.join(os.tmpdir(), 'uploads') : path.join(__dirname, 'uploads');
const maxUploadBytes = runningOnVercel ? Math.floor(4.2 * 1024 * 1024) : 25 * 1024 * 1024;
fsSync.mkdirSync(uploadDir, { recursive: true });

function toBool(value, defaultValue = false) {
  if (value === undefined || value === null || value === '') {
    return defaultValue;
  }

  if (typeof value === 'boolean') {
    return value;
  }

  const normalized = String(value).trim().toLowerCase();
  return ['1', 'true', 'yes', 'on'].includes(normalized);
}

function allowedFile(file) {
  const ext = path.extname(file.originalname || '').toLowerCase();
  return ['.pdf', '.docx', '.doc', '.txt', '.md'].includes(ext);
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname || '');
    const base = path.basename(file.originalname || 'upload', ext).replace(/[^a-z0-9-_]+/gi, '_').slice(0, 80);
    cb(null, `${Date.now()}-${base || 'document'}${ext}`);
  },
});

const upload = multer({
  storage,
  limits: {
    fileSize: maxUploadBytes,
  },
  fileFilter: (req, file, cb) => {
    if (!allowedFile(file)) {
      cb(new Error('Unsupported file type. Upload PDF, DOCX, DOC, TXT, or MD.'));
      return;
    }
    cb(null, true);
  },
});

app.use(express.json({ limit: '1mb' }));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/health', (req, res) => {
  res.json({
    ok: true,
    timestamp: new Date().toISOString(),
  });
});

app.post('/api/quiz/upload', upload.single('document'), async (req, res) => {
  let uploadedPath = null;

  try {
    if (!req.file) {
      res.status(400).json({ error: 'No file uploaded.' });
      return;
    }

    uploadedPath = req.file.path;

    const text = await extractTextFromDocument(uploadedPath, req.file.originalname);
    if (!text || text.trim().length < 30) {
      res.status(400).json({ error: 'The uploaded file appears empty or unreadable.' });
      return;
    }

    const quiz = buildQuizFromText(text);

    if (!quiz.questions.length) {
      res.status(400).json({
        error: 'No valid multiple-choice questions were detected. Check document format (Question N + answer choices).',
      });
      return;
    }

    const useGemini = toBool(req.body.useGemini, false);
    let gemini = {
      attempted: false,
      updatedQuestions: 0,
      reason: 'Disabled by request.',
    };

    if (useGemini) {
      gemini = await enrichQuestionsWithGemini(quiz.questions, {
        model: process.env.GEMINI_MODEL || 'gemini-2.5-flash-lite',
      });
    }

    for (const question of quiz.questions) {
      for (const option of question.options) {
        if (!question.explanations[option.label] || question.explanations[option.label].trim().length < 5) {
          if (question.sourceExplanation) {
            question.explanations[option.label] = question.sourceExplanation;
            if (question.explanationSource === 'none') {
              question.explanationSource = 'document';
            }
          } else {
            question.explanations[option.label] = 'Explanation not available in source. Enable Gemini with a valid API key to auto-generate rationale.';
          }
        }
      }
    }

    const tutorModeDefault = toBool(req.body.tutorMode, true);
    const timedModeDefault = toBool(req.body.timedMode, false);
    const timedSecondsPerQuestion = 90;

    res.json({
      quiz,
      processing: {
        fileName: req.file.originalname,
        parsing: quiz.parsing,
        gemini,
      },
      defaults: {
        tutorMode: tutorModeDefault,
        timedMode: timedModeDefault,
        timedSecondsPerQuestion,
        timedSecondsTotal: quiz.questions.length * timedSecondsPerQuestion,
      },
    });
  } catch (error) {
    res.status(500).json({
      error: error.message || 'Failed to process document.',
    });
  } finally {
    if (uploadedPath) {
      try {
        await fs.unlink(uploadedPath);
      } catch (error) {
        // No-op: upload cleanup should not fail the request.
      }
    }
  }
});

app.use((error, req, res, next) => {
  if (error && error.code === 'LIMIT_FILE_SIZE') {
    const limitMb = Math.floor(maxUploadBytes / (1024 * 1024));
    res.status(413).json({
      error: `Uploaded file is too large. Max size is ${limitMb}MB in this deployment.`,
    });
    return;
  }

  res.status(400).json({
    error: error.message || 'Upload failed.',
  });
});

module.exports = {
  app,
  PORT,
};

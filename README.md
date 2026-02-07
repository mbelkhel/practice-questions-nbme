# NBME-Style Local Quiz App

Local web app to upload practice-question documents (`.pdf`, `.docx`, `.doc`) and run them in an NBME-like quiz interface.

## Features

- Uploads and parses question banks from document files.
- Detects question blocks and answer/explanation sections (including `Question 1` + `Q1. Answer:` style).
- Links explanations from the bottom answer section back to each question.
- NBME-like solving tools:
  - timed mode at `90 sec/question` (auto)
  - always-on stopwatch
  - tutor mode with `Submit Answer` per question (selection is editable until submit)
  - defaults on load: `Tutor ON`, `Timed OFF`, `Use Gemini OFF`
  - question flagging
  - option strikeout by clicking answer text
  - circle selector for final answer choice
  - drag-to-highlight in stem, click-highlight to remove only that highlight
  - question palette navigation
- Stem image support:
  - inline markers such as `[IMAGE:https://...jpg]`
  - Markdown images like `![desc](https://...png)`
  - `.docx` embedded images are extracted into inline image markers when available
- Uses Gemini only when answers/explanations are missing or insufficient.
- Fills per-option rationale for correct and incorrect answers.

## Project Structure

- `server.js`: Express API + static frontend host
- `backendApp.js`: shared Express app setup (local + Vercel serverless)
- `api/index.js`: Vercel serverless entrypoint
- `src/fileTextExtractor.js`: PDF/DOCX/DOC/TXT extraction
- `src/quizParser.js`: question/answer parsing + linking
- `src/geminiService.js`: explanation generation fallback
- `public/index.html`: app UI
- `public/styles.css`: UI styles
- `public/app.js`: quiz client behavior

## Setup

```bash
npm install
cp .env.example .env
```

Set your Gemini key in `.env`:

```bash
GEMINI_API_KEY=your_key_here
```

## Run

```bash
npm start
```

Open:

- <http://localhost:3000>

## Dev Mode

```bash
npm run dev
```

## Deploy To Vercel (Free Hobby)

1. Push this folder to a GitHub repository.
2. In Vercel, click **Add New Project** and import that repo.
3. In Vercel project settings, add environment variables:
   - `GEMINI_API_KEY`
   - `GEMINI_MODEL` (optional, default is `gemini-2.5-flash-lite`)
4. Deploy.
5. Share the generated `https://<project>.vercel.app` URL.

This repo includes:

- `vercel.json` routing config
- `api/index.js` serverless API entry
- `backendApp.js` to reuse the same backend logic in Vercel

## Gemini Notes (Free Tier)

The app defaults to:

- `GEMINI_MODEL=gemini-2.5-flash-lite`

As of **February 5, 2026**, Google AI docs list this model as available in free tier limits for Gemini API usage.

If you want a different model, update `GEMINI_MODEL` in `.env`.

## API Endpoint

`POST /api/quiz/upload` (`multipart/form-data`)

Fields:

- `document` (required)
- `timedMode` (`true`/`false`)
- `tutorMode` (`true`/`false`)
- `useGemini` (`true`/`false`)

Returns parsed quiz JSON + parsing/generation metadata.

## Quick Verification With Your Sample

```bash
curl -s -F document=@"/Users/malachybelkhelladi/Downloads/Rheumatoid arthritis test yourself questions1.docx" \
  -F timedMode=true -F tutorMode=true -F useGemini=false \
  http://localhost:3000/api/quiz/upload
```

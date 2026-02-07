const fs = require('fs/promises');
const path = require('path');
const mammoth = require('mammoth');
const pdfParse = require('pdf-parse');
const WordExtractor = require('word-extractor');

function normalizeWhitespace(text) {
  return text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/[\u00A0\t]+/g, ' ')
    .replace(/[ ]{2,}/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function decodeHtmlEntities(text) {
  return text
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
}

function htmlToTextWithImageMarkers(html) {
  const withMarkers = String(html || '')
    .replace(/<img[^>]*src=["']([^"']+)["'][^>]*>/gi, '\n[IMAGE:$1]\n')
    .replace(/<(?:\/p|\/div|\/h[1-6]|\/li|\/tr|\/table|\/ul|\/ol)>/gi, '\n')
    .replace(/<br\s*\/?\s*>/gi, '\n')
    .replace(/<[^>]+>/g, '');

  return normalizeWhitespace(decodeHtmlEntities(withMarkers));
}

async function extractFromPdf(filePath) {
  const buffer = await fs.readFile(filePath);
  const parsed = await pdfParse(buffer);
  return normalizeWhitespace(parsed.text || '');
}

async function extractFromDocx(filePath) {
  const [rawResult, htmlResult] = await Promise.all([
    mammoth.extractRawText({ path: filePath }),
    mammoth.convertToHtml({
      path: filePath,
      convertImage: mammoth.images.inline((element) =>
        element.read('base64').then((imageBase64) => ({
          src: `data:${element.contentType};base64,${imageBase64}`,
        })),
      ),
    }),
  ]);

  const rawText = normalizeWhitespace(rawResult.value || '');
  const htmlText = htmlToTextWithImageMarkers(htmlResult.value || '');
  const hasImageMarkers = /\[IMAGE:/i.test(htmlText);

  if (hasImageMarkers) {
    return htmlText;
  }

  if (rawText.length >= htmlText.length) {
    return rawText;
  }

  return htmlText;
}

async function extractFromDoc(filePath) {
  const extractor = new WordExtractor();
  const extracted = await extractor.extract(filePath);
  return normalizeWhitespace(extracted.getBody() || '');
}

async function extractFromPlainText(filePath) {
  const text = await fs.readFile(filePath, 'utf8');
  return normalizeWhitespace(text);
}

async function extractTextFromDocument(filePath, originalName = '') {
  const ext = path.extname(originalName || filePath).toLowerCase();

  if (ext === '.pdf') {
    return extractFromPdf(filePath);
  }

  if (ext === '.docx') {
    return extractFromDocx(filePath);
  }

  if (ext === '.doc') {
    try {
      return await extractFromDoc(filePath);
    } catch (error) {
      throw new Error('Unable to parse .doc file. Convert to .docx if parsing fails.');
    }
  }

  if (ext === '.txt' || ext === '.md') {
    return extractFromPlainText(filePath);
  }

  throw new Error('Unsupported file type. Upload PDF, DOCX, DOC, TXT, or MD.');
}

module.exports = {
  extractTextFromDocument,
  normalizeWhitespace,
};

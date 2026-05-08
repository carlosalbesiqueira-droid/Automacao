const fs = require('node:fs/promises');
const path = require('node:path');
const { safeFileName } = require('../utils/fs');

async function fetchNoteDocumentFromUrl(noteUrl, outputDirectory, fallbackName = 'nota') {
  const noteUrlCandidate = resolveSupportedNoteUrl(noteUrl);
  if (!noteUrlCandidate) {
    throw new Error('Link da NF nao suportado para captura automatica.');
  }

  const attemptUrls = buildFetchAttemptUrls(noteUrlCandidate);
  let lastError = null;

  for (const attemptUrl of attemptUrls) {
    try {
      const response = await fetch(attemptUrl, {
        headers: {
          'user-agent': 'Mozilla/5.0 AutomacaoY3'
        }
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }

      const contentType = String(response.headers.get('content-type') || '').toLowerCase();
      const buffer = Buffer.from(await response.arrayBuffer());
      const fileName = resolveFileName(attemptUrl, fallbackName, contentType);
      const filePath = path.join(outputDirectory, fileName);

      await fs.writeFile(filePath, buffer);

      if (contentType.includes('pdf')) {
        return {
          kind: 'pdf',
          sourceUrl: noteUrlCandidate,
          resolvedUrl: attemptUrl,
          fileName,
          filePath,
          content: buffer,
        };
      }

      const html = buffer.toString('utf8');
      return {
        kind: 'html',
        sourceUrl: noteUrlCandidate,
        resolvedUrl: attemptUrl,
        fileName,
        filePath,
        content: buffer,
        text: extractTextFromHtml(html),
      };
    } catch (error) {
      lastError = error;
    }
  }

  throw new Error(`Falha ao baixar a NF pelo link informado: ${lastError?.message || 'erro desconhecido'}.`);
}

function resolveSupportedNoteUrl(noteUrl) {
  const text = String(noteUrl || '').trim();
  if (!text) {
    return '';
  }

  try {
    const url = new URL(text);
    if (/nfse\.salvador\.ba\.gov\.br/i.test(url.hostname) && /notaprint\.aspx$/i.test(url.pathname)) {
      url.hash = '';
      return url.toString();
    }

    if (/notaprintpdf\.aspx$/i.test(url.pathname)) {
      url.hash = '';
      return url.toString();
    }

    if (/notaprint\.aspx$/i.test(url.pathname)) {
      url.hash = '';
      return url.toString();
    }
  } catch {
    return '';
  }

  return '';
}

function buildFetchAttemptUrls(noteUrl) {
  const urls = [];
  const directUrl = String(noteUrl || '').trim();
  if (!directUrl) {
    return urls;
  }

  const pdfUrl = buildPdfUrlFromNoteUrl(directUrl);
  if (pdfUrl && pdfUrl !== directUrl) {
    urls.push(pdfUrl);
  }

  urls.push(directUrl);
  return Array.from(new Set(urls));
}

function buildPdfUrlFromNoteUrl(noteUrl) {
  const text = String(noteUrl || '').trim();
  if (!text) {
    return '';
  }

  try {
    const url = new URL(text);
    if (/nfse\.salvador\.ba\.gov\.br/i.test(url.hostname)) {
      return text;
    }

    if (/notaprintpdf\.aspx$/i.test(url.pathname)) {
      url.hash = '';
      return url.toString();
    }

    if (/notaprint\.aspx$/i.test(url.pathname)) {
      url.pathname = url.pathname.replace(/notaprint\.aspx$/i, 'notaprintpdf.aspx');
      url.hash = '';
      return url.toString();
    }
  } catch {
    return '';
  }

  return '';
}

function resolveFileName(documentUrl, fallbackName, contentType = '') {
  try {
    const url = new URL(documentUrl);
    const nf = url.searchParams.get('nf');
    if (nf) {
      const extension = contentType.includes('pdf') ? 'pdf' : 'html';
      return safeFileName(`nota_${nf}.${extension}`);
    }
  } catch {
    // usa fallback abaixo
  }

  const extension = contentType.includes('pdf') ? 'pdf' : 'html';
  return safeFileName(`${fallbackName || 'nota'}.${extension}`);
}

function extractTextFromHtml(html) {
  return String(html || '')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<[^>]+>/g, '\n')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&#39;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\r/g, '\n')
    .replace(/\n{2,}/g, '\n')
    .trim();
}

module.exports = {
  buildPdfUrlFromNoteUrl,
  extractTextFromHtml,
  fetchNoteDocumentFromUrl,
  resolveSupportedNoteUrl,
};

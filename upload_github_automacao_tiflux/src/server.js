const path = require('node:path');
const fs = require('node:fs/promises');
const { execFile } = require('node:child_process');
const { Readable } = require('node:stream');
const { promisify } = require('node:util');
const dotenv = require('dotenv');
const express = require('express');
const multer = require('multer');
const { previewEmailSource, previewComparisonSource, processEmailAutomation, processEmailComparison } = require('./process-email');
const { composeFollowUpPayload } = require('./email/follow-up');
const { buildOutlookWebComposeUrl, sendEmailWithOutlook } = require('./email/outlook');
const { ensureDirectory, makeTimestamp, safeFileName } = require('./utils/fs');

dotenv.config();

const app = express();
const execFileAsync = promisify(execFile);
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024
  }
});

const PORT = Number(process.env.PORT || process.env.WEB_PORT || 3210);
const BOT_FATURAS_API_BASE = String(process.env.BOT_FATURAS_API_BASE || 'http://127.0.0.1:8321').replace(/\/+$/, '');
const PUBLIC_DIR = path.resolve(__dirname, '../public');
const runs = new Map();

app.set('trust proxy', true);
app.use((request, response, next) => {
  response.header('Access-Control-Allow-Origin', '*');
  response.header('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  response.header('Access-Control-Allow-Headers', 'Content-Type');

  if (request.method === 'OPTIONS') {
    response.sendStatus(204);
    return;
  }

  next();
});
app.use('/bot-faturas-api', proxyBotFaturasApi);
app.use(express.json({ limit: '20mb' }));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(PUBLIC_DIR));

app.get('/api/health', (_request, response) => {
  response.json({
    ok: true,
    app: 'automacao-y3-nfse',
    mode: 'web'
  });
});

app.get('/api/health/full', async (_request, response) => {
  try {
    const apiResponse = await fetch(`${BOT_FATURAS_API_BASE}/health`);
    const apiPayload = await apiResponse.json();
    response.status(apiResponse.ok && apiPayload?.ok ? 200 : 503).json({
      ok: Boolean(apiResponse.ok && apiPayload?.ok),
      app: 'automacao-y3-nfse',
      mode: 'web+bot-faturas-api',
      botFaturasApi: apiPayload,
    });
  } catch (error) {
    response.status(503).json({
      ok: false,
      app: 'automacao-y3-nfse',
      mode: 'web+bot-faturas-api',
      error: `API interna indisponivel: ${error.message}`,
      target: BOT_FATURAS_API_BASE,
    });
  }
});

app.post('/api/faturas/validate', optionalUpload, async (request, response) => {
  try {
    const input = await buildFaturasInput(request);
    const payload = await runBotFaturasService('validate', input);
    response.json(payload);
  } catch (error) {
    response.status(400).json({
      ok: false,
      error: error.message
    });
  }
});

app.post('/api/faturas/run', optionalUpload, async (request, response) => {
  try {
    const input = await buildFaturasInput(request);
    const payload = await runBotFaturasService('run', input);
    const runId = `faturas-${makeTimestamp().replace(/[^0-9]/g, '')}`;

    runs.set(runId, {
      id: runId,
      kind: 'bot-faturas',
      createdAt: new Date().toISOString(),
      files: {
        archivePath: payload.files?.archive_path || '',
        reportPath: payload.files?.report_path || '',
        updatedSourcePath: payload.files?.updated_source_path || '',
        snapshotPath: payload.files?.snapshot_path || '',
      },
      summary: payload.summary,
      source: payload.summary?.source || {},
    });

    response.json({
      ...payload,
      runId,
      files: {
        executionRoot: payload.files?.execution_root || '',
        archiveName: payload.files?.archive_path ? path.basename(payload.files.archive_path) : '',
        archiveUrl: payload.files?.archive_path ? `/api/faturas/runs/${runId}/file/archive` : '',
        reportName: payload.files?.report_path ? path.basename(payload.files.report_path) : '',
        reportUrl: payload.files?.report_path ? `/api/faturas/runs/${runId}/file/report` : '',
        updatedSourceName: payload.files?.updated_source_path ? path.basename(payload.files.updated_source_path) : '',
        updatedSourceUrl: payload.files?.updated_source_path ? `/api/faturas/runs/${runId}/file/updated-source` : '',
        snapshotName: payload.files?.snapshot_path ? path.basename(payload.files.snapshot_path) : '',
        snapshotUrl: payload.files?.snapshot_path ? `/api/faturas/runs/${runId}/file/snapshot` : '',
      },
    });
  } catch (error) {
    response.status(400).json({
      ok: false,
      error: error.message
    });
  }
});

app.post('/api/process', optionalUpload, async (request, response) => {
  try {
    const input = buildProcessInput(request);
    validateWebInput(input);

    const result = await processEmailAutomation(input);
    const runId = result.timestamp;

    runs.set(runId, {
      id: runId,
      createdAt: new Date().toISOString(),
      workbookPath: result.workbookPath,
      generalFilePath: result.generalFilePath,
      pendingFilePath: result.pendingFilePath,
      differenceFilePath: result.differenceFilePath,
      emailSummaryPath: result.emailSummaryPath,
      outputDirectory: result.outputDirectory,
      pendingRows: result.pendingRows,
      followUpRows: result.followUpRows || result.pendingRows,
      followUpDraft: result.followUpDraft,
    });

    response.json({
      ok: true,
      runId,
      summary: {
        subject: result.email.subject,
        from: result.email.from,
        date: result.email.date,
        spreadsheetName: result.spreadsheetAttachment?.filename || '',
        pdfCount: result.pdfAttachments.length,
        finalRowCount: result.finalRows.length,
        pendingRowCount: result.pendingRows.length,
        tableCount: result.selectedTables.length,
        warnings: result.warnings
      },
      generalTable: result.generalTable
        ? {
            id: result.generalTable.id,
            title: result.generalTable.title,
            headers: result.generalTable.headers,
            rows: result.generalTable.rows.slice(0, 60),
            rowCount: result.generalTable.rowCount,
            columnCount: result.generalTable.columnCount,
            truncated: result.generalTable.rows.length > 60
          }
        : null,
      pendingTable: result.pendingTable
        ? {
            id: result.pendingTable.id,
            title: result.pendingTable.title,
            headers: result.pendingTable.headers,
            rows: result.pendingTable.rows.slice(0, 60),
            rowCount: result.pendingTable.rowCount,
            columnCount: result.pendingTable.columnCount,
            truncated: result.pendingTable.rows.length > 60
          }
        : null,
      differenceTable: result.differenceTable
        ? {
            id: result.differenceTable.id,
            title: result.differenceTable.title,
            headers: result.differenceTable.headers,
            rows: result.differenceTable.rows.slice(0, 60),
            rowCount: result.differenceTable.rowCount,
            columnCount: result.differenceTable.columnCount,
            truncated: result.differenceTable.rows.length > 60
          }
        : null,
      files: {
        outputDirectory: result.outputDirectory,
        workbookName: path.basename(result.workbookPath),
        workbookUrl: `/api/runs/${runId}/file/workbook`,
        generalFileName: result.generalFilePath ? path.basename(result.generalFilePath) : '',
        generalFileUrl: result.generalFilePath ? `/api/runs/${runId}/file/general` : '',
        pendingFileName: result.pendingFilePath ? path.basename(result.pendingFilePath) : '',
        pendingFileUrl: result.pendingFilePath ? `/api/runs/${runId}/file/pending` : '',
        differenceFileName: result.differenceFilePath ? path.basename(result.differenceFilePath) : '',
        differenceFileUrl: result.differenceFilePath ? `/api/runs/${runId}/file/difference` : '',
        summaryName: path.basename(result.emailSummaryPath),
        summaryUrl: `/api/runs/${runId}/file/summary`
      },
      followUp: result.followUpDraft
        ? {
            available: Boolean((result.followUpRows || result.pendingRows || []).length),
            from: result.followUpDraft.from,
            to: result.followUpDraft.to,
            cc: result.followUpDraft.cc,
            subject: result.followUpDraft.subject,
            message: result.followUpDraft.message,
            pendingCount: result.followUpDraft.pendingCount,
            attachmentNames: result.followUpDraft.attachmentNames,
            mode: result.followUpDraft.mode,
            itemLabel: result.followUpDraft.itemLabel,
          }
        : {
            available: false,
            from: '',
            to: '',
            cc: '',
            subject: '',
            message: '',
            pendingCount: 0,
            attachmentNames: [],
            mode: 'pending',
            itemLabel: 'pendencia(s)',
          }
    });
  } catch (error) {
    response.status(400).json({
      ok: false,
      error: error.message
    });
  }
});

app.post('/api/runs/:runId/follow-up/outlook', async (request, response) => {
  try {
    const run = runs.get(request.params.runId);
    if (!run) {
      response.status(404).json({ ok: false, error: 'Execução não encontrada.' });
      return;
    }

    const followUpRows = Array.isArray(run.followUpRows) ? run.followUpRows : (Array.isArray(run.pendingRows) ? run.pendingRows : []);
    const baseAttachmentPath = run.followUpDraft?.attachmentPaths?.[0] || run.pendingFilePath || run.differenceFilePath || '';
    const signatureImage = await persistSignatureImage(run, request.body.signatureImage);

    const action = request.body.action === 'send' ? 'send' : 'draft';
    const draft = composeFollowUpPayload({
      from: request.body.from || run.followUpDraft?.from || '',
      to: request.body.to || run.followUpDraft?.to || '',
      cc: request.body.cc || run.followUpDraft?.cc || '',
      subject: request.body.subject || run.followUpDraft?.subject || '',
      message: request.body.message || run.followUpDraft?.message || '',
      pendingRows: followUpRows,
      attachmentPaths: [baseAttachmentPath].filter(Boolean),
      mode: run.followUpDraft?.mode || 'pending',
      signatureImage,
    });

    const outlookResult = await sendEmailWithOutlook({
      from: draft.from,
      to: draft.to,
      cc: draft.cc,
      subject: draft.subject,
      htmlBody: draft.htmlBody,
      textBody: draft.textBody,
      attachmentPaths: draft.attachmentPaths,
      inlineAttachments: draft.inlineAttachments,
      sendNow: action === 'send',
    });

    run.followUpDraft = {
      ...run.followUpDraft,
      from: draft.from,
      to: draft.to,
      cc: draft.cc,
      subject: draft.subject,
      message: draft.message,
      attachmentNames: draft.attachmentNames,
      pendingCount: draft.pendingCount,
      signatureImageName: draft.signatureImageName,
    };

    response.json({
      ok: true,
      action: outlookResult.action || action,
      message: outlookResult.message || (action === 'send'
        ? 'E-mail enviado pelo Outlook com a planilha anexa.'
        : 'Rascunho aberto no Outlook com a planilha anexa.'),
      followUp: {
        from: draft.from,
        to: draft.to,
        cc: draft.cc,
        subject: draft.subject,
        pendingCount: draft.pendingCount,
        attachmentNames: draft.attachmentNames,
        signatureImageName: draft.signatureImageName,
        fallback: outlookResult.fallback || '',
      },
    });
  } catch (error) {
    response.status(400).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post('/api/runs/:runId/follow-up/outlook-web', async (request, response) => {
  try {
    const run = runs.get(request.params.runId);
    if (!run) {
      response.status(404).json({ ok: false, error: 'Execucao nao encontrada.' });
      return;
    }

    const followUpRows = Array.isArray(run.followUpRows) ? run.followUpRows : (Array.isArray(run.pendingRows) ? run.pendingRows : []);
    const baseAttachmentPath = run.followUpDraft?.attachmentPaths?.[0] || run.pendingFilePath || run.differenceFilePath || '';
    const signatureImage = await persistSignatureImage(run, request.body.signatureImage);
    const baseUrl = resolvePublicBaseUrl(request);
    const attachmentInfo = resolveFollowUpAttachmentInfo(run, baseAttachmentPath);

    const draft = composeFollowUpPayload({
      from: request.body.from || run.followUpDraft?.from || '',
      to: request.body.to || run.followUpDraft?.to || '',
      cc: request.body.cc || run.followUpDraft?.cc || '',
      subject: request.body.subject || run.followUpDraft?.subject || '',
      message: request.body.message || run.followUpDraft?.message || '',
      pendingRows: followUpRows,
      attachmentPaths: [baseAttachmentPath].filter(Boolean),
      mode: run.followUpDraft?.mode || 'pending',
      signatureImage,
    });

    const attachmentUrl = attachmentInfo.relativeUrl ? new URL(attachmentInfo.relativeUrl, baseUrl).toString() : '';
    const signatureUrl = signatureImage?.path
      ? new URL(`/api/runs/${run.id}/file/signature/${encodeURIComponent(path.basename(signatureImage.path))}`, baseUrl).toString()
      : '';
    const webBody = buildOutlookWebBody(draft.textBody, {
      attachmentUrl,
      attachmentName: attachmentInfo.name,
      signatureUrl,
      signatureName: signatureImage?.fileName || '',
    });
    const composeUrl = buildOutlookWebComposeUrl({
      to: draft.to,
      cc: draft.cc,
      subject: draft.subject,
      body: webBody,
    });

    run.followUpDraft = {
      ...run.followUpDraft,
      from: draft.from,
      to: draft.to,
      cc: draft.cc,
      subject: draft.subject,
      message: draft.message,
      attachmentNames: draft.attachmentNames,
      pendingCount: draft.pendingCount,
      signatureImageName: draft.signatureImageName,
    };

    response.json({
      ok: true,
      action: 'draft-web',
      message: attachmentUrl
        ? 'Outlook Web aberto com a mensagem preenchida. A planilha de apoio sera baixada neste aparelho para voce anexar antes de enviar.'
        : 'Outlook Web aberto com a mensagem preenchida.',
      composeUrl,
      attachmentUrl,
      attachmentName: attachmentInfo.name,
      signatureUrl,
      signatureName: signatureImage?.fileName || '',
      followUp: {
        from: draft.from,
        to: draft.to,
        cc: draft.cc,
        subject: draft.subject,
        pendingCount: draft.pendingCount,
        attachmentNames: draft.attachmentNames,
        signatureImageName: draft.signatureImageName,
        fallback: 'outlook-web',
      },
    });
  } catch (error) {
    response.status(400).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post('/api/process-compare', optionalUpload, async (request, response) => {
  try {
    const input = buildCompareInput(request);
    validateCompareInput(input);

    const result = await processEmailComparison(input);
    const runId = result.timestamp;

    runs.set(runId, {
      id: runId,
      createdAt: new Date().toISOString(),
      workbookPath: result.workbookPath,
      generalFilePath: result.generalFilePath,
      pendingFilePath: result.pendingFilePath,
      differenceFilePath: result.differenceFilePath,
      emailSummaryPath: result.emailSummaryPath,
      outputDirectory: result.outputDirectory,
      pendingRows: result.pendingRows,
      followUpRows: result.followUpRows || result.pendingRows,
      followUpDraft: result.followUpDraft,
    });

    response.json({
      ok: true,
      runId,
      summary: {
        subject: `Pedido x Retorno`,
        requestSubject: result.requestEmail.subject,
        responseSubject: result.responseEmail.subject,
        pdfCount: result.pdfAttachments.length,
        requestedRowCount: result.requestedRows.length,
        finalRowCount: result.finalRows.length,
        pendingRowCount: result.pendingRows.length,
        differenceRowCount: result.differenceSummary?.totalCount || 0,
        requestOnlyCount: result.differenceSummary?.requestOnlyCount || 0,
        responseOnlyCount: result.differenceSummary?.responseOnlyCount || 0,
        tableCount: result.selectedTables.length,
        warnings: result.warnings
      },
      generalTable: result.generalTable
        ? {
            id: result.generalTable.id,
            title: result.generalTable.title,
            headers: result.generalTable.headers,
            rows: result.generalTable.rows.slice(0, 60),
            rowCount: result.generalTable.rowCount,
            columnCount: result.generalTable.columnCount,
            truncated: result.generalTable.rows.length > 60
          }
        : null,
      pendingTable: result.pendingTable
        ? {
            id: result.pendingTable.id,
            title: result.pendingTable.title,
            headers: result.pendingTable.headers,
            rows: result.pendingTable.rows.slice(0, 60),
            rowCount: result.pendingTable.rowCount,
            columnCount: result.pendingTable.columnCount,
            truncated: result.pendingTable.rows.length > 60
          }
        : null,
      differenceTable: result.differenceTable
        ? {
            id: result.differenceTable.id,
            title: result.differenceTable.title,
            headers: result.differenceTable.headers,
            rows: result.differenceTable.rows.slice(0, 60),
            rowCount: result.differenceTable.rowCount,
            columnCount: result.differenceTable.columnCount,
            truncated: result.differenceTable.rows.length > 60
          }
        : null,
      files: {
        outputDirectory: result.outputDirectory,
        workbookName: path.basename(result.workbookPath),
        workbookUrl: `/api/runs/${runId}/file/workbook`,
        generalFileName: result.generalFilePath ? path.basename(result.generalFilePath) : '',
        generalFileUrl: result.generalFilePath ? `/api/runs/${runId}/file/general` : '',
        pendingFileName: result.pendingFilePath ? path.basename(result.pendingFilePath) : '',
        pendingFileUrl: result.pendingFilePath ? `/api/runs/${runId}/file/pending` : '',
        differenceFileName: result.differenceFilePath ? path.basename(result.differenceFilePath) : '',
        differenceFileUrl: result.differenceFilePath ? `/api/runs/${runId}/file/difference` : '',
        summaryName: path.basename(result.emailSummaryPath),
        summaryUrl: `/api/runs/${runId}/file/summary`
      },
      followUp: result.followUpDraft
        ? {
            available: Boolean((result.followUpRows || result.pendingRows || []).length),
            from: result.followUpDraft.from,
            to: result.followUpDraft.to,
            cc: result.followUpDraft.cc,
            subject: result.followUpDraft.subject,
            message: result.followUpDraft.message,
            pendingCount: result.followUpDraft.pendingCount,
            attachmentNames: result.followUpDraft.attachmentNames,
            mode: result.followUpDraft.mode,
            itemLabel: result.followUpDraft.itemLabel,
          }
        : {
            available: false,
            from: '',
            to: '',
            cc: '',
            subject: '',
            message: '',
            pendingCount: 0,
            attachmentNames: [],
            mode: 'pending',
            itemLabel: 'pendencia(s)',
          }
    });
  } catch (error) {
    response.status(400).json({
      ok: false,
      error: error.message
    });
  }
});

app.post('/api/preview-email', optionalUpload, async (request, response) => {
  try {
    const input = buildProcessInput(request);
    validateWebInput(input);

    const preview = await previewEmailSource(input);
    response.json({
      ok: true,
      preview
    });
  } catch (error) {
    response.status(400).json({
      ok: false,
      error: error.message
    });
  }
});

app.post('/api/preview-compare-side', optionalUpload, async (request, response) => {
  try {
    const side = request.body.side === 'request' ? 'request' : 'response';
    const fileField = side === 'request' ? 'requestEmlFile' : 'responseEmlFile';
    const uploadFile = getUploadedFile(request, fileField);

    if (!uploadFile?.buffer) {
      throw new Error(`Envie o arquivo .eml no portal "${side === 'request' ? 'o que pedimos' : 'o que o cliente mandou'}".`);
    }

    const preview = await previewComparisonSource(buildCompareSideInput(uploadFile, side));

    response.json({
      ok: true,
      preview
    });
  } catch (error) {
    response.status(400).json({
      ok: false,
      error: error.message
    });
  }
});

app.get('/api/runs/:runId/file/:kind', (request, response) => {
  const run = runs.get(request.params.runId);
  if (!run) {
    response.status(404).json({ ok: false, error: 'Execução não encontrada.' });
    return;
  }

  let filePath = run.workbookPath;
  if (request.params.kind === 'summary') {
    filePath = run.emailSummaryPath;
  } else if (request.params.kind === 'general') {
    filePath = run.generalFilePath;
  } else if (request.params.kind === 'pending') {
    filePath = run.pendingFilePath;
  } else if (request.params.kind === 'difference') {
    filePath = run.differenceFilePath;
  }

  if (!filePath) {
    response.status(404).json({ ok: false, error: 'Arquivo nao encontrado para esta execucao.' });
    return;
  }

  response.download(filePath, path.basename(filePath));
});

app.get('/api/faturas/runs/:runId/file/:kind', (request, response) => {
  const run = runs.get(request.params.runId);
  if (!run || run.kind !== 'bot-faturas') {
    response.status(404).json({ ok: false, error: 'Execucao do bot de faturas nao encontrada.' });
    return;
  }

  let filePath = '';
  if (request.params.kind === 'archive') {
    filePath = run.files.archivePath;
  } else if (request.params.kind === 'report') {
    filePath = run.files.reportPath;
  } else if (request.params.kind === 'updated-source') {
    filePath = run.files.updatedSourcePath;
  } else if (request.params.kind === 'snapshot') {
    filePath = run.files.snapshotPath;
  }

  if (!filePath) {
    response.status(404).json({ ok: false, error: 'Arquivo nao encontrado para esta execucao do bot de faturas.' });
    return;
  }

  response.download(filePath, path.basename(filePath));
});

app.get('/api/runs/:runId/file/signature/:fileName', (request, response) => {
  const run = runs.get(request.params.runId);
  if (!run) {
    response.status(404).json({ ok: false, error: 'Execucao nao encontrada.' });
    return;
  }

  const signatureDirectory = path.join(run.outputDirectory, 'assinatura');
  const filePath = path.join(signatureDirectory, path.basename(request.params.fileName || ''));

  response.download(filePath, path.basename(filePath), (error) => {
    if (error && !response.headersSent) {
      response.status(404).json({ ok: false, error: 'Arquivo de assinatura nao encontrado.' });
    }
  });
});

app.use((_request, response) => {
  if (_request.method !== 'GET') {
    response.status(404).json({
      ok: false,
      error: 'Rota não encontrada.'
    });
    return;
  }

  response.sendFile(path.join(PUBLIC_DIR, 'index.html'));
});

app.listen(PORT, () => {
  console.log(`Painel Y3 disponível em http://localhost:${PORT}`);
});

async function proxyBotFaturasApi(request, response) {
  try {
    const targetPath = request.originalUrl.replace(/^\/bot-faturas-api/, '') || '/';
    const targetUrl = new URL(targetPath, `${BOT_FATURAS_API_BASE}/`);
    const headers = { ...request.headers };
    delete headers.host;
    delete headers.connection;
    delete headers['content-length'];
    delete headers['accept-encoding'];
    headers['x-forwarded-host'] = request.get('host');
    headers['x-forwarded-proto'] = request.protocol;

    const apiResponse = await fetch(targetUrl, {
      method: request.method,
      headers,
      body: ['GET', 'HEAD'].includes(request.method) ? undefined : request,
      duplex: ['GET', 'HEAD'].includes(request.method) ? undefined : 'half',
      redirect: 'manual',
    });

    response.status(apiResponse.status);
    apiResponse.headers.forEach((value, key) => {
      if (key.toLowerCase() === 'transfer-encoding') {
        return;
      }
      response.setHeader(key, value);
    });

    if (!apiResponse.body) {
      response.end();
      return;
    }

    Readable.fromWeb(apiResponse.body).pipe(response);
  } catch (error) {
    response.status(502).json({
      ok: false,
      error: `Falha ao conectar com a API interna do bot de faturas: ${error.message}`,
      target: BOT_FATURAS_API_BASE,
    });
  }
}

async function persistSignatureImage(run, signatureImage) {
  if (!signatureImage?.base64) {
    return null;
  }

  const originalName = safeFileName(signatureImage.name || 'assinatura.png');
  const extension = path.extname(originalName) || extensionFromMimeType(signatureImage.type);
  const finalName = extension
    ? ensureExtension(originalName, extension)
    : `${path.basename(originalName, path.extname(originalName)) || 'assinatura'}.png`;
  const signatureDirectory = path.join(run.outputDirectory, 'assinatura');
  await ensureDirectory(signatureDirectory);

  const stamp = makeTimestamp();
  const filePath = path.join(signatureDirectory, `${stamp}_${finalName}`);
  await fs.writeFile(filePath, Buffer.from(String(signatureImage.base64 || ''), 'base64'));

  return {
    path: filePath,
    cid: `assinatura-y3-${stamp.replace(/[^a-zA-Z0-9]/g, '')}`,
    fileName: path.basename(filePath),
    type: signatureImage.type || '',
  };
}

function resolvePublicBaseUrl(request) {
  const forwardedProto = String(request.headers['x-forwarded-proto'] || '').split(',')[0].trim();
  const protocol = forwardedProto || request.protocol || 'http';
  const host = request.get('host');
  return `${protocol}://${host}`;
}

function resolveFollowUpAttachmentInfo(run, baseAttachmentPath) {
  if (!baseAttachmentPath) {
    return { relativeUrl: '', name: '' };
  }

  if (run.pendingFilePath && baseAttachmentPath === run.pendingFilePath) {
    return {
      relativeUrl: `/api/runs/${run.id}/file/pending`,
      name: path.basename(run.pendingFilePath),
    };
  }

  if (run.differenceFilePath && baseAttachmentPath === run.differenceFilePath) {
    return {
      relativeUrl: `/api/runs/${run.id}/file/difference`,
      name: path.basename(run.differenceFilePath),
    };
  }

  if (run.generalFilePath && baseAttachmentPath === run.generalFilePath) {
    return {
      relativeUrl: `/api/runs/${run.id}/file/general`,
      name: path.basename(run.generalFilePath),
    };
  }

  if (run.workbookPath && baseAttachmentPath === run.workbookPath) {
    return {
      relativeUrl: `/api/runs/${run.id}/file/workbook`,
      name: path.basename(run.workbookPath),
    };
  }

  return {
    relativeUrl: '',
    name: path.basename(baseAttachmentPath),
  };
}

function buildOutlookWebBody(baseText, { attachmentUrl = '', attachmentName = '', signatureUrl = '', signatureName = '' } = {}) {
  return String(baseText || '').trim();
}

function buildProcessInput(request) {
  const mode = request.body.mode === 'upload' ? 'upload' : 'imap';
  const base = {
    output: request.body.output || 'output',
    selectedAttachmentIds: normalizeListValue(request.body.selectedAttachmentIds),
    attachmentSelectionApplied: request.body.attachmentSelectionApplied === 'true',
    selectedTableIds: normalizeListValue(request.body.selectedTableIds),
    tableSelectionApplied: request.body.tableSelectionApplied === 'true',
    connection: {
      provider: request.body.provider || process.env.MAIL_PROVIDER || 'outlook',
      user: request.body.emailUser || '',
      pass: request.body.emailPass || '',
      host: request.body.customHost || '',
      port: request.body.customPort || '',
      secure: request.body.customSecure || ''
    }
  };

  if (mode === 'upload') {
    const uploadFile = getUploadedFile(request, 'emlFile');
    return {
      ...base,
      emlBuffer: uploadFile?.buffer,
      emlFilename: uploadFile?.originalname || 'upload.eml'
    };
  }

  return {
    ...base,
    subject: request.body.subject,
    mailbox: request.body.mailbox || process.env.MAILBOX || 'INBOX',
    unseenOnly: request.body.unseenOnly === 'true' || request.body.unseenOnly === 'on'
  };
}

function buildCompareInput(request) {
  const requestFile = getUploadedFile(request, 'requestEmlFile');
  const responseFile = getUploadedFile(request, 'responseEmlFile');

  return {
    output: request.body.output || 'output',
    requestSource: {
      ...(isSpreadsheetFile(requestFile)
        ? {
            spreadsheetBuffer: requestFile?.buffer,
            spreadsheetFilename: requestFile?.originalname || 'pedido.xlsx',
          }
        : {
            emlBuffer: requestFile?.buffer,
            emlFilename: requestFile?.originalname || 'pedido.eml',
          }),
      selectedAttachmentIds: normalizeListValue(request.body.requestSelectedAttachmentIds),
      attachmentSelectionApplied: request.body.requestAttachmentSelectionApplied === 'true',
      selectedTableIds: normalizeListValue(request.body.requestSelectedTableIds),
      tableSelectionApplied: request.body.requestTableSelectionApplied === 'true'
    },
    responseSource: {
      ...(isSpreadsheetFile(responseFile)
        ? {
            spreadsheetBuffer: responseFile?.buffer,
            spreadsheetFilename: responseFile?.originalname || 'retorno.xlsx',
          }
        : {
            emlBuffer: responseFile?.buffer,
            emlFilename: responseFile?.originalname || 'retorno.eml',
          }),
      selectedAttachmentIds: normalizeListValue(request.body.responseSelectedAttachmentIds),
      attachmentSelectionApplied: request.body.responseAttachmentSelectionApplied === 'true',
      selectedTableIds: normalizeListValue(request.body.responseSelectedTableIds),
      tableSelectionApplied: request.body.responseTableSelectionApplied === 'true'
    }
  };
}

function validateWebInput(input) {
  if (input.emlBuffer) {
    return;
  }

  if (input.subject) {
    return;
  }

  throw new Error('Informe o assunto do e-mail ou envie um arquivo .eml.');
}

function validateCompareInput(input) {
  if (!input.requestSource?.emlBuffer && !input.requestSource?.spreadsheetBuffer) {
    throw new Error('Envie um arquivo .eml ou uma planilha no portal "o que pedimos".');
  }

  if (!input.responseSource?.emlBuffer && !input.responseSource?.spreadsheetBuffer) {
    throw new Error('Envie um arquivo .eml ou uma planilha no portal "o que o cliente mandou".');
  }
}

function optionalUpload(request, response, next) {
  const contentType = request.headers['content-type'] || '';
  if (!contentType.includes('multipart/form-data')) {
    next();
    return;
  }

  upload.fields([
    { name: 'emlFile', maxCount: 1 },
    { name: 'requestEmlFile', maxCount: 1 },
    { name: 'responseEmlFile', maxCount: 1 },
    { name: 'faturasSpreadsheetFile', maxCount: 1 }
  ])(request, response, next);
}

function getUploadedFile(request, fieldName) {
  if (request.file && request.file.fieldname === fieldName) {
    return request.file;
  }

  if (request.files && Array.isArray(request.files[fieldName]) && request.files[fieldName][0]) {
    return request.files[fieldName][0];
  }

  return null;
}

function normalizeListValue(value) {
  if (Array.isArray(value)) {
    return value;
  }

  if (value == null || value === '') {
    return [];
  }

  return [value];
}

function buildCompareSideInput(uploadFile, side) {
  if (isSpreadsheetFile(uploadFile)) {
    return {
      spreadsheetBuffer: uploadFile?.buffer,
      spreadsheetFilename: uploadFile?.originalname || `${side}.xlsx`,
    };
  }

  return {
    emlBuffer: uploadFile?.buffer,
    emlFilename: uploadFile?.originalname || `${side}.eml`,
  };
}

function isSpreadsheetFile(file) {
  const name = String(file?.originalname || '').toLowerCase();
  return /\.(xlsx|xlsm|csv|xls)$/i.test(name);
}

function extensionFromMimeType(value) {
  const mime = String(value || '').toLowerCase();
  if (mime.includes('png')) {
    return '.png';
  }
  if (mime.includes('jpeg') || mime.includes('jpg')) {
    return '.jpg';
  }
  if (mime.includes('gif')) {
    return '.gif';
  }
  if (mime.includes('webp')) {
    return '.webp';
  }
  if (mime.includes('bmp')) {
    return '.bmp';
  }
  return '';
}

function ensureExtension(fileName, extension) {
  if (String(fileName).toLowerCase().endsWith(String(extension).toLowerCase())) {
    return fileName;
  }
  return `${fileName}${extension}`;
}

async function buildFaturasInput(request) {
  const sourceType = request.body.sourceType === 'upload' ? 'upload' : 'google-sheet';
  const limitValue = Number(request.body.limit || '');
  const spreadsheetFile = getUploadedFile(request, 'faturasSpreadsheetFile');

  const input = {
    sourceType,
    inputFile: '',
    limit: Number.isFinite(limitValue) && limitValue > 0 ? limitValue : undefined,
    ids: normalizeListValue((request.body.ids || '').split(',').map((item) => item.trim()).filter(Boolean)),
    includeReview: request.body.includeReview === 'true' || request.body.includeReview === 'on',
    dryRun: request.body.dryRun === 'true' || request.body.dryRun === 'on',
    headless: !(request.body.showBrowser === 'true' || request.body.showBrowser === 'on'),
  };

  if (sourceType === 'upload') {
    if (!spreadsheetFile?.buffer) {
      throw new Error('Envie uma planilha .xlsx, .xlsm, .xls ou .csv para usar o bot por upload.');
    }

    input.inputFile = await persistBotFaturasUpload(spreadsheetFile);
  }

  return input;
}

async function persistBotFaturasUpload(file) {
  const uploadsDirectory = path.resolve(__dirname, '../output/bot_faturas/_uploads');
  await ensureDirectory(uploadsDirectory);
  const stamp = makeTimestamp().replace(/[^0-9]/g, '');
  const originalName = safeFileName(file?.originalname || 'bot_faturas.xlsx');
  const targetPath = path.join(uploadsDirectory, `${stamp}_${originalName}`);
  await fs.writeFile(targetPath, file.buffer);
  return targetPath;
}

async function runBotFaturasService(command, input) {
  const pythonExecutable = await resolvePythonExecutable();
  const scriptPath = path.resolve(__dirname, '../scripts/bot_faturas_service.py');
  const args = [scriptPath, command];

  if (input.inputFile) {
    args.push('--input-file', input.inputFile);
  }
  if (command === 'run' && input.limit) {
    args.push('--limit', String(input.limit));
  }
  if (command === 'run' && input.ids?.length) {
    args.push('--ids', input.ids.join(','));
  }
  if (command === 'run' && input.includeReview) {
    args.push('--include-review');
  }
  if (command === 'run' && input.dryRun) {
    args.push('--dry-run');
  }
  if (command === 'run' && typeof input.headless === 'boolean') {
    args.push(input.headless ? '--headless' : '--no-headless');
  }

  const { stdout, stderr } = await execFileAsync(pythonExecutable, args, {
    cwd: path.resolve(__dirname, '..'),
    env: {
      ...process.env,
      PYTHONIOENCODING: 'utf-8',
    },
    maxBuffer: 20 * 1024 * 1024,
  });

  const rawOutput = String(stdout || '').trim();
  if (!rawOutput) {
    throw new Error(String(stderr || 'O servico do bot de faturas nao retornou resposta.').trim());
  }

  try {
    return JSON.parse(rawOutput);
  } catch {
    throw new Error(String(stderr || rawOutput || 'Falha ao interpretar a resposta do bot de faturas.').trim());
  }
}

async function resolvePythonExecutable() {
  const localVenv = path.resolve(__dirname, '../.venv/Scripts/python.exe');
  try {
    await fs.access(localVenv);
    return localVenv;
  } catch {
    return process.platform === 'win32' ? 'python' : 'python3';
  }
}

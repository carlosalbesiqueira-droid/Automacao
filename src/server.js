const path = require('node:path');
const dotenv = require('dotenv');
const express = require('express');
const multer = require('multer');
const { previewEmailSource, previewComparisonSource, processEmailAutomation, processEmailComparison } = require('./process-email');
const { composeFollowUpPayload } = require('./email/follow-up');
const { sendEmailWithOutlook } = require('./email/outlook');

dotenv.config();

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024
  }
});

const PORT = Number(process.env.WEB_PORT || 3210);
const PUBLIC_DIR = path.resolve(__dirname, '../public');
const runs = new Map();

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
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
app.use(express.static(PUBLIC_DIR));

app.get('/api/health', (_request, response) => {
  response.json({
    ok: true,
    app: 'automacao-y3-nfse',
    mode: 'web'
  });
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
      emailSummaryPath: result.emailSummaryPath,
      outputDirectory: result.outputDirectory,
      pendingRows: result.pendingRows,
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
      files: {
        outputDirectory: result.outputDirectory,
        workbookName: path.basename(result.workbookPath),
        workbookUrl: `/api/runs/${runId}/file/workbook`,
        generalFileName: result.generalFilePath ? path.basename(result.generalFilePath) : '',
        generalFileUrl: result.generalFilePath ? `/api/runs/${runId}/file/general` : '',
        pendingFileName: result.pendingFilePath ? path.basename(result.pendingFilePath) : '',
        pendingFileUrl: result.pendingFilePath ? `/api/runs/${runId}/file/pending` : '',
        summaryName: path.basename(result.emailSummaryPath),
        summaryUrl: `/api/runs/${runId}/file/summary`
      },
      followUp: result.followUpDraft
        ? {
            available: Boolean(result.pendingRows.length),
            to: result.followUpDraft.to,
            cc: result.followUpDraft.cc,
            subject: result.followUpDraft.subject,
            message: result.followUpDraft.message,
            pendingCount: result.followUpDraft.pendingCount,
            attachmentNames: result.followUpDraft.attachmentNames,
          }
        : {
            available: false,
            to: '',
            cc: '',
            subject: '',
            message: '',
            pendingCount: 0,
            attachmentNames: [],
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
      response.status(404).json({ ok: false, error: 'Execucao nao encontrada.' });
      return;
    }

    if (!Array.isArray(run.pendingRows) || !run.pendingRows.length) {
      throw new Error('Nao existem pendencias remanescentes para cobrar nesta execucao.');
    }

    const action = request.body.action === 'send' ? 'send' : 'draft';
    const draft = composeFollowUpPayload({
      to: request.body.to || run.followUpDraft?.to || '',
      cc: request.body.cc || run.followUpDraft?.cc || '',
      subject: request.body.subject || run.followUpDraft?.subject || '',
      message: request.body.message || run.followUpDraft?.message || '',
      pendingRows: run.pendingRows,
      attachmentPaths: [run.pendingFilePath].filter(Boolean),
    });

    const outlookResult = await sendEmailWithOutlook({
      to: draft.to,
      cc: draft.cc,
      subject: draft.subject,
      htmlBody: draft.htmlBody,
      attachmentPaths: draft.attachmentPaths,
      sendNow: action === 'send',
    });

    run.followUpDraft = {
      ...run.followUpDraft,
      to: draft.to,
      cc: draft.cc,
      subject: draft.subject,
      message: draft.message,
      attachmentNames: draft.attachmentNames,
      pendingCount: draft.pendingCount,
    };

    response.json({
      ok: true,
      action: outlookResult.action || action,
      message: action === 'send'
        ? 'E-mail enviado pelo Outlook com a planilha de pendencias anexada.'
        : 'Rascunho aberto no Outlook com a planilha de pendencias anexada.',
      followUp: {
        to: draft.to,
        cc: draft.cc,
        subject: draft.subject,
        pendingCount: draft.pendingCount,
        attachmentNames: draft.attachmentNames,
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
      emailSummaryPath: result.emailSummaryPath,
      outputDirectory: result.outputDirectory,
      pendingRows: result.pendingRows,
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
      files: {
        outputDirectory: result.outputDirectory,
        workbookName: path.basename(result.workbookPath),
        workbookUrl: `/api/runs/${runId}/file/workbook`,
        generalFileName: result.generalFilePath ? path.basename(result.generalFilePath) : '',
        generalFileUrl: result.generalFilePath ? `/api/runs/${runId}/file/general` : '',
        pendingFileName: result.pendingFilePath ? path.basename(result.pendingFilePath) : '',
        pendingFileUrl: result.pendingFilePath ? `/api/runs/${runId}/file/pending` : '',
        summaryName: path.basename(result.emailSummaryPath),
        summaryUrl: `/api/runs/${runId}/file/summary`
      },
      followUp: result.followUpDraft
        ? {
            available: Boolean(result.pendingRows.length),
            to: result.followUpDraft.to,
            cc: result.followUpDraft.cc,
            subject: result.followUpDraft.subject,
            message: result.followUpDraft.message,
            pendingCount: result.followUpDraft.pendingCount,
            attachmentNames: result.followUpDraft.attachmentNames,
          }
        : {
            available: false,
            to: '',
            cc: '',
            subject: '',
            message: '',
            pendingCount: 0,
            attachmentNames: [],
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
  }

  if (!filePath) {
    response.status(404).json({ ok: false, error: 'Arquivo nao encontrado para esta execucao.' });
    return;
  }

  response.download(filePath, path.basename(filePath));
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
    { name: 'responseEmlFile', maxCount: 1 }
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

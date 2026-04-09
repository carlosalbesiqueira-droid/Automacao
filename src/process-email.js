const path = require('node:path');
const { CANONICAL_COLUMNS } = require('./config/canonical-columns');
const {
  fetchEmailBySubject,
  loadEmailFromBuffer,
  loadEmailFromEml,
  saveAttachments,
  summarizeEmail
} = require('./email/fetch-email');
const { buildCombinedTable } = require('./email/table-utils');
const { buildFollowUpDraft } = require('./email/follow-up');
const { collapseDuplicateRows, mergeSpreadsheetRowsWithInvoices } = require('./normalize/merge');
const { buildPendingReport, pickRequestTable } = require('./normalize/pending');
const { writeOutputWorkbook } = require('./output/write-workbook');
const { fetchNoteDocumentFromUrl } = require('./parsers/note-link');
const { parseInvoicePdfText, parsePdfAttachment } = require('./parsers/pdf');
const {
  normalizeEmailTable,
  normalizeSpreadsheetAttachment
} = require('./parsers/spreadsheet');
const { ensureDirectory, makeTimestamp, writeTextFile } = require('./utils/fs');
const { digitsOnly, formatDate, normalizeKey } = require('./utils/normalizers');

async function processEmailAutomation(input) {
  const timestamp = makeTimestamp();
  const outputDirectory = path.resolve(input.output || 'output', timestamp);
  const attachmentsDirectory = path.join(outputDirectory, 'anexos');
  const warnings = [];

  await ensureDirectory(attachmentsDirectory);

  const email = await resolveEmailSource(input);
  const selectedTables = filterTablesBySelection(
    email.tables || [],
    input.selectedTableIds,
    input.tableSelectionApplied,
  );

  const selectedAttachments = filterAttachmentsBySelection(
    email.attachments,
    input.selectedAttachmentIds,
    input.attachmentSelectionApplied,
  );
  const savedAttachments = await saveAttachments(selectedAttachments, attachmentsDirectory);
  const spreadsheetAttachments = savedAttachments.filter((attachment) => attachment.kind === 'spreadsheet');
  const spreadsheetAttachment = spreadsheetAttachments[0] || null;
  const pdfAttachments = savedAttachments.filter((attachment) => attachment.kind === 'pdf');

  if (!spreadsheetAttachment && !selectedTables.length && !pdfAttachments.length) {
    throw new Error('Selecione ao menos uma tabela, uma planilha anexada ou um PDF de nota para gerar o arquivo geral.');
  }

  if (spreadsheetAttachments.length > 1) {
    warnings.push(
      `Mais de uma planilha foi selecionada. A automacao usou apenas "${spreadsheetAttachment.filename}" como base principal.`,
    );
  }

  if (!pdfAttachments.length) {
    warnings.push('Nenhum PDF de nota fiscal foi encontrado. A saida sera gerada apenas com base na planilha anexada e/ou nas tabelas selecionadas.');
  }

  const spreadsheetRows = spreadsheetAttachment
    ? await normalizeSpreadsheetAttachment(spreadsheetAttachment, warnings)
    : [];

  const tableRows = collapseDuplicateRows(
    selectedTables.flatMap((table) => normalizeEmailTable(table, warnings)),
  );

  if (!spreadsheetAttachment && selectedTables.length) {
    warnings.push('Nenhuma planilha anexada foi selecionada. A automacao usou as tabelas do corpo do e-mail como base para o arquivo geral.');
  }

  const invoices = [];
  for (const attachment of pdfAttachments) {
    const invoice = await parsePdfAttachment(attachment, warnings);
    invoices.push(invoice);
  }

  const baseRows = buildBaseRows(spreadsheetRows, tableRows, invoices, warnings);
  const appendUnmatchedInvoices = true;
  const finalRows = baseRows.length
    ? mergeSpreadsheetRowsWithInvoices(baseRows, invoices, warnings, {
        appendUnmatchedInvoices,
        warnUnmatchedRows: !(spreadsheetRows.length && tableRows.length)
      })
    : mergeSpreadsheetRowsWithInvoices([], invoices, warnings, {
        appendUnmatchedInvoices
      });

  await enrichRowsFromNoteLinks(finalRows, attachmentsDirectory, warnings);
  harmonizeRowsWithOfficialLink(finalRows);
  const normalizedFinalRows = collapseDuplicateRows(finalRows);

  const generalTable = buildGeneralTable(normalizedFinalRows, selectedTables);
  const requestTable = pickRequestTable(selectedTables);
  const requestedRows = requestTable
    ? collapseDuplicateRows(normalizeEmailTable(requestTable, []))
    : [];
  const { pendingRows, pendingTable } = buildPendingReport(requestedRows, normalizedFinalRows);

  if (selectedTables.length && !requestTable) {
    warnings.push('Nao foi possivel identificar automaticamente a tabela principal de pendencias solicitadas.');
  }

  const { workbookPath, generalFilePath, pendingFilePath } = await writeOutputWorkbook({
    outputDirectory,
    email,
    rows: normalizedFinalRows,
    pendingRows,
    warnings,
    selectedTables,
    generalTable,
    pendingTable,
  });

  const emailSummaryPath = path.join(outputDirectory, 'resumo_email.txt');
  await writeTextFile(emailSummaryPath, summarizeEmail(email));

  return {
    timestamp,
    email,
    spreadsheetAttachment,
    pdfAttachments,
    finalRows: normalizedFinalRows,
    generalTable,
    pendingRows,
    pendingTable,
    warnings,
    outputDirectory,
    workbookPath,
    generalFilePath,
    pendingFilePath,
    emailSummaryPath,
    selectedTables
  };
}

async function processEmailComparison(input) {
  const timestamp = makeTimestamp();
  const outputDirectory = path.resolve(input.output || 'output', timestamp);
  const requestDirectory = path.join(outputDirectory, 'pedido');
  const responseDirectory = path.join(outputDirectory, 'retorno');
  const warnings = [];

  await ensureDirectory(requestDirectory);
  await ensureDirectory(responseDirectory);

  const requestEmail = await resolveComparisonSource(input.requestSource);
  const responseEmail = await resolveComparisonSource(input.responseSource);

  const requestContext = await prepareRequestContext(
    requestEmail,
    input.requestSource,
    requestDirectory,
    warnings,
  );

  const responseContext = await prepareResponseContext(
    responseEmail,
    input.responseSource,
    responseDirectory,
    warnings,
  );

  const { pendingRows, pendingTable } = buildPendingReport(
    requestContext.requestedRows,
    responseContext.finalRows,
  );

  const combinedSelectedTables = [
    ...requestContext.selectedTables.map((table) => ({
      ...table,
      title: prefixTableTitle('Pedido', table.title),
    })),
    ...responseContext.selectedTables.map((table) => ({
      ...table,
      title: prefixTableTitle('Retorno', table.title),
    })),
  ];

  const comparisonEmailSummary = buildComparisonEmailSummary(requestEmail, responseEmail);

  const { workbookPath, generalFilePath, pendingFilePath } = await writeOutputWorkbook({
    outputDirectory,
    email: comparisonEmailSummary,
    rows: responseContext.finalRows,
    pendingRows,
    warnings,
    selectedTables: combinedSelectedTables,
    generalTable: responseContext.generalTable,
    pendingTable,
  });

  const emailSummaryPath = path.join(outputDirectory, 'resumo_comparativo.txt');
  await writeTextFile(
    emailSummaryPath,
    [
      'Pedido',
      summarizeEmail(requestEmail),
      '',
      'Retorno',
      summarizeEmail(responseEmail),
    ].join('\n\n'),
  );

  const followUpDraft = buildFollowUpDraft({
    pendingRows,
    requestEmail,
    responseEmail,
    pendingFilePath,
  });

  return {
    timestamp,
    requestEmail,
    responseEmail,
    requestedRows: requestContext.requestedRows,
    finalRows: responseContext.finalRows,
    generalTable: responseContext.generalTable,
    pendingRows,
    pendingTable,
    warnings,
    outputDirectory,
    workbookPath,
    generalFilePath,
    pendingFilePath,
    emailSummaryPath,
    selectedTables: combinedSelectedTables,
    pdfAttachments: responseContext.pdfAttachments,
    followUpDraft,
  };
}

async function previewEmailSource(input) {
  const email = await resolveEmailSource(input);

  return {
    source: email.source,
    provider: email.provider || '',
    mailbox: email.mailbox,
    subject: email.subject || '(sem assunto)',
    from: email.from || '',
    to: email.to || '',
    date: email.date || null,
    text: (email.text || '').trim().slice(0, 12000),
    tableCount: email.tables?.length || 0,
    tables: (email.tables || []).map((table) => ({
      id: table.id,
      title: table.title,
      headers: table.headers,
      rows: table.rows.slice(0, 30),
      rowCount: table.rowCount,
      columnCount: table.columnCount,
      truncated: table.rows.length > 30
    })),
    attachmentCount: email.attachments.length,
    attachments: email.attachments.map((attachment) => ({
      id: attachment.id,
      filename: attachment.filename,
      contentType: attachment.contentType,
      size: attachment.size,
      kind: attachment.kind
    }))
  };
}

async function previewComparisonSource(input) {
  if (input.spreadsheetBuffer) {
    const warnings = [];
    const rows = collapseDuplicateRows(await normalizeSpreadsheetAttachment({
      filename: input.spreadsheetFilename || 'planilha.xlsx',
      content: input.spreadsheetBuffer,
    }, warnings));
    const previewTable = buildRowsPreviewTable(rows, `Planilha ${input.spreadsheetFilename || ''}`.trim());

    return {
      source: 'spreadsheet',
      provider: '',
      mailbox: 'upload de planilha',
      subject: input.spreadsheetFilename || 'planilha enviada',
      from: '',
      to: '',
      date: null,
      text: [
        `Arquivo de planilha carregado diretamente: ${input.spreadsheetFilename || 'planilha enviada'}`,
        `${rows.length} linha(s) uteis encontradas apos a normalizacao.`,
        warnings.length ? `Avisos detectados: ${warnings.length}.` : 'Nenhum aviso automatico nesta leitura.',
      ].join('\n'),
      tableCount: previewTable ? 1 : 0,
      tables: previewTable ? [previewTable] : [],
      attachmentCount: 0,
      attachments: [],
    };
  }

  return previewEmailSource(input);
}

async function resolveEmailSource(input) {
  if (input.emlBuffer) {
    return loadEmailFromBuffer(input.emlBuffer, {
      sourcePath: input.emlFilename || 'upload.eml',
      mailbox: 'upload .eml'
    });
  }

  if (input.emlPath) {
    return loadEmailFromEml(input.emlPath);
  }

  return fetchEmailBySubject({
    subject: input.subject,
    mailbox: input.mailbox,
    unseenOnly: Boolean(input.unseenOnly),
    connection: input.connection || {}
  });
}

async function resolveComparisonSource(input) {
  if (input.spreadsheetBuffer) {
    return {
      source: 'spreadsheet',
      provider: '',
      mailbox: 'upload de planilha',
      subject: input.spreadsheetFilename || 'planilha enviada',
      from: '',
      to: '',
      date: null,
      text: `Arquivo de planilha carregado diretamente: ${input.spreadsheetFilename || 'planilha enviada'}`,
      sourcePath: input.spreadsheetFilename || 'planilha.xlsx',
      attachments: [],
      tables: [],
    };
  }

  return resolveEmailSource(input);
}

async function prepareRequestContext(email, sourceInput, targetDirectory, warnings) {
  if (sourceInput.spreadsheetBuffer) {
    const requestedRows = collapseDuplicateRows(await normalizeSpreadsheetAttachment({
      filename: sourceInput.spreadsheetFilename || 'pedido.xlsx',
      content: sourceInput.spreadsheetBuffer,
    }, warnings));

    if (!requestedRows.length) {
      throw new Error('Nao foi possivel identificar as pendencias solicitadas na planilha do portal "o que pedimos".');
    }

    return {
      selectedTables: [buildRowsPreviewTable(requestedRows, 'Planilha do pedido')].filter(Boolean),
      requestedRows,
    };
  }

  const selectedTables = filterTablesBySelection(
    email.tables || [],
    sourceInput.selectedTableIds,
    sourceInput.tableSelectionApplied,
  );

  const selectedAttachments = filterAttachmentsBySelection(
    email.attachments,
    sourceInput.selectedAttachmentIds,
    sourceInput.attachmentSelectionApplied,
  );

  const savedAttachments = await saveAttachments(selectedAttachments, targetDirectory);
  const spreadsheetAttachments = savedAttachments.filter((attachment) => attachment.kind === 'spreadsheet');
  const spreadsheetAttachment = spreadsheetAttachments[0] || null;

  if (!spreadsheetAttachment && !selectedTables.length) {
    throw new Error('No portal "o que pedimos", selecione ao menos uma tabela ou uma planilha com as pendencias solicitadas.');
  }

  if (spreadsheetAttachments.length > 1) {
    warnings.push(
      `Mais de uma planilha foi selecionada no portal de pedido. A automacao usou apenas "${spreadsheetAttachment.filename}" como base principal do que foi solicitado.`,
    );
  }

  const spreadsheetRows = spreadsheetAttachment
    ? await normalizeSpreadsheetAttachment(spreadsheetAttachment, warnings)
    : [];

  const tableRows = collapseDuplicateRows(
    selectedTables.flatMap((table) => normalizeEmailTable(table, warnings)),
  );

  const requestedRows = collapseDuplicateRows([...spreadsheetRows, ...tableRows]);
  if (!requestedRows.length) {
    throw new Error('Nao foi possivel identificar as pendencias solicitadas no portal "o que pedimos".');
  }

  return {
    selectedTables,
    requestedRows,
  };
}

async function prepareResponseContext(email, sourceInput, targetDirectory, warnings) {
  if (sourceInput.spreadsheetBuffer) {
    const spreadsheetRows = collapseDuplicateRows(await normalizeSpreadsheetAttachment({
      filename: sourceInput.spreadsheetFilename || 'retorno.xlsx',
      content: sourceInput.spreadsheetBuffer,
    }, warnings));

    if (!spreadsheetRows.length) {
      throw new Error('Nao foi possivel identificar dados uteis na planilha do portal "o que o cliente mandou".');
    }

    await enrichRowsFromNoteLinks(spreadsheetRows, targetDirectory, warnings);
    harmonizeRowsWithOfficialLink(spreadsheetRows);
    const finalRows = collapseDuplicateRows(spreadsheetRows);

    return {
      selectedTables: [buildRowsPreviewTable(finalRows, 'Planilha do retorno')].filter(Boolean),
      pdfAttachments: [],
      finalRows,
      generalTable: buildGeneralTable(finalRows, []),
    };
  }

  const selectedTables = filterTablesBySelection(
    email.tables || [],
    sourceInput.selectedTableIds,
    sourceInput.tableSelectionApplied,
  );

  const selectedAttachments = filterAttachmentsBySelection(
    email.attachments,
    sourceInput.selectedAttachmentIds,
    sourceInput.attachmentSelectionApplied,
  );

  const savedAttachments = await saveAttachments(selectedAttachments, targetDirectory);
  const spreadsheetAttachments = savedAttachments.filter((attachment) => attachment.kind === 'spreadsheet');
  const spreadsheetAttachment = spreadsheetAttachments[0] || null;
  const pdfAttachments = savedAttachments.filter((attachment) => attachment.kind === 'pdf');

  if (!spreadsheetAttachment && !selectedTables.length && !pdfAttachments.length) {
    throw new Error('No portal "o que o cliente mandou", selecione ao menos uma tabela, uma planilha anexada ou um PDF oficial.');
  }

  if (spreadsheetAttachments.length > 1) {
    warnings.push(
      `Mais de uma planilha foi selecionada no portal de retorno. A automacao usou apenas "${spreadsheetAttachment.filename}" como base principal do cliente.`,
    );
  }

  if (!pdfAttachments.length) {
    warnings.push('No portal de retorno, nenhum PDF oficial foi encontrado. O comparativo vai considerar apenas links oficiais e a planilha do cliente.');
  }

  const spreadsheetRows = spreadsheetAttachment
    ? await normalizeSpreadsheetAttachment(spreadsheetAttachment, warnings)
    : [];

  const tableRows = collapseDuplicateRows(
    selectedTables.flatMap((table) => normalizeEmailTable(table, warnings)),
  );

  const invoices = [];
  for (const attachment of pdfAttachments) {
    const invoice = await parsePdfAttachment(attachment, warnings);
    invoices.push(invoice);
  }

  const baseRows = buildBaseRows(spreadsheetRows, tableRows, invoices, warnings);
  const appendUnmatchedInvoices = true;
  const finalRows = baseRows.length
    ? mergeSpreadsheetRowsWithInvoices(baseRows, invoices, warnings, {
        appendUnmatchedInvoices,
        warnUnmatchedRows: !(spreadsheetRows.length && tableRows.length),
      })
    : mergeSpreadsheetRowsWithInvoices([], invoices, warnings, {
        appendUnmatchedInvoices,
      });

  await enrichRowsFromNoteLinks(finalRows, targetDirectory, warnings);
  harmonizeRowsWithOfficialLink(finalRows);
  const normalizedFinalRows = collapseDuplicateRows(finalRows);
  const generalTable = buildGeneralTable(normalizedFinalRows, selectedTables);

  return {
    selectedTables,
    pdfAttachments,
    finalRows: normalizedFinalRows,
    generalTable,
  };
}

function buildBaseRows(spreadsheetRows, tableRows, invoices, warnings) {
  if (spreadsheetRows.length && tableRows.length) {
    return buildCrossCheckedBaseRows(spreadsheetRows, tableRows, invoices, warnings);
  }

  if (!spreadsheetRows.length) {
    return tableRows;
  }

  return collapseDuplicateRows([...spreadsheetRows, ...tableRows]);
}

function buildCrossCheckedBaseRows(spreadsheetRows, tableRows, invoices, warnings) {
  const baseRows = [];
  const usedSpreadsheetIndexes = new Set();

  tableRows.forEach((tableRow) => {
    const spreadsheetIndex = findBestSpreadsheetMatchIndex(tableRow, spreadsheetRows, usedSpreadsheetIndexes);
    const invoiceIndex = findBestInvoiceSupportIndex(tableRow, invoices);

    if (spreadsheetIndex < 0 && invoiceIndex < 0) {
      warnings.push(`Linha da tabela ignorada por falta de confirmacao na planilha ou nos PDFs: ${buildRowIdentity(tableRow)}.`);
      return;
    }

    let mergedRow = { ...tableRow };
    if (spreadsheetIndex >= 0) {
      mergedRow = mergeCanonicalRows({ ...spreadsheetRows[spreadsheetIndex] }, tableRow);
      usedSpreadsheetIndexes.add(spreadsheetIndex);
    }

    baseRows.push(mergedRow);
  });

  spreadsheetRows.forEach((spreadsheetRow, index) => {
    if (usedSpreadsheetIndexes.has(index)) {
      return;
    }

    baseRows.push({ ...spreadsheetRow });
  });

  return collapseDuplicateRows(baseRows);
}

function findBestSpreadsheetMatchIndex(tableRow, spreadsheetRows, usedSpreadsheetIndexes) {
  let bestIndex = -1;
  let bestScore = 0;

  spreadsheetRows.forEach((spreadsheetRow, index) => {
    if (usedSpreadsheetIndexes.has(index)) {
      return;
    }

    const score = scoreSpreadsheetSupport(tableRow, spreadsheetRow);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestScore >= 60 ? bestIndex : -1;
}

function findBestInvoiceSupportIndex(row, invoices) {
  let bestIndex = -1;
  let bestScore = 0;

  invoices.forEach((invoice, index) => {
    const score = scoreInvoiceSupport(row, invoice);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestScore >= 60 ? bestIndex : -1;
}

function scoreSpreadsheetSupport(tableRow, spreadsheetRow) {
  let score = 0;

  if (documentNumbersMatch(tableRow.rps, spreadsheetRow.rps)) {
    score += 100;
  }

  if (documentNumbersMatch(tableRow.rps, spreadsheetRow.dps)) {
    score += 90;
  }

  if (tableRow.tomador && spreadsheetRow.tomador && normalizeKey(tableRow.tomador) === normalizeKey(spreadsheetRow.tomador)) {
    score += 20;
  }

  if (tableRow.cnpj && spreadsheetRow.cnpj && tableRow.cnpj === spreadsheetRow.cnpj) {
    score += 20;
  }

  return score;
}

function scoreInvoiceSupport(row, invoice) {
  let score = 0;

  if (documentNumbersMatch(row.rps, invoice.rps)) {
    score += 100;
  }

  if (documentNumbersMatch(row.rps, invoice.dps)) {
    score += 90;
  }

  if (row.tomador && invoice.tomador && normalizeKey(row.tomador) === normalizeKey(invoice.tomador)) {
    score += 20;
  }

  if (row.cnpj && invoice.cnpj && row.cnpj === invoice.cnpj) {
    score += 20;
  }

  return score;
}

function mergeCanonicalRows(primaryRow, supplementalRow) {
  const merged = { ...primaryRow };

  for (const column of CANONICAL_COLUMNS) {
    const key = column.key;
    const primaryValue = merged[key];
    const supplementalValue = supplementalRow[key];

    if ((primaryValue == null || primaryValue === '') && supplementalValue != null && supplementalValue !== '') {
      merged[key] = supplementalValue;
      continue;
    }

    if (shouldPreferSupplementalValue(key, primaryValue, supplementalValue)) {
      merged[key] = supplementalValue;
    }
  }

  if (!merged.arquivoPdf && supplementalRow.arquivoPdf) {
    merged.arquivoPdf = supplementalRow.arquivoPdf;
  }

  if (!merged.notaUrl && supplementalRow.notaUrl) {
    merged.notaUrl = supplementalRow.notaUrl;
  }

  return merged;
}

function shouldPreferSupplementalValue(key, primaryValue, supplementalValue) {
  if (supplementalValue == null || supplementalValue === '') {
    return false;
  }

  if (key === 'rps') {
    return String(supplementalValue).length > String(primaryValue || '').length;
  }

  if (key === 'dps' || key === 'nfse' || key === 'numeroObra') {
    return String(supplementalValue).length > String(primaryValue || '').length;
  }

  if (key === 'tomador' || key === 'intermediario' || key === 'cartaDe') {
    return alphanumericRichness(supplementalValue) > alphanumericRichness(primaryValue);
  }

  if (key === 'cnpj') {
    return !primaryValue && Boolean(supplementalValue);
  }

  return false;
}

function buildRowIdentity(row) {
  return row.rps || row.nfse || row.dps || row.tomador || 'sem identificador';
}

async function enrichRowsFromNoteLinks(rows, attachmentsDirectory, warnings) {
  const cache = new Map();

  for (const row of rows) {
    if (!row.notaUrl || !needsNoteLinkEnrichment(row)) {
      continue;
    }

    try {
      let parsedNote = cache.get(row.notaUrl);
      if (!parsedNote) {
        const downloaded = await fetchNoteDocumentFromUrl(row.notaUrl, attachmentsDirectory, row.nfse || row.rps || 'nota');
        if (downloaded.kind === 'pdf') {
          parsedNote = await parsePdfAttachment({
            filename: downloaded.fileName,
            savedPath: downloaded.filePath,
            content: downloaded.content
          }, warnings);
        } else {
          parsedNote = parseInvoicePdfText(downloaded.text || '', downloaded.fileName);
          parsedNote.arquivoPdf = '';
          parsedNote.arquivoPdfNome = '';
        }

        parsedNote.notaUrl = row.notaUrl;
        cache.set(row.notaUrl, parsedNote);
      }

      const enrichedRow = mergeCanonicalRows(row, parsedNote);
      Object.assign(row, enrichedRow);
      if (!row.arquivoPdf && parsedNote.arquivoPdf) {
        row.arquivoPdf = parsedNote.arquivoPdf;
      }
    } catch (error) {
      warnings.push(`Nao foi possivel complementar a linha ${buildRowIdentity(row)} pelo link da NF: ${error.message}`);
    }
  }
}

function needsNoteLinkEnrichment(row) {
  return !row.tomador
    || !row.cnpj
    || row.valorServico == null
    || row.valorServico === ''
    || row.issDevido == null
    || row.issDevido === ''
    || row.issPagar == null
    || row.issPagar === '';
}

function harmonizeRowsWithOfficialLink(rows) {
  for (const row of rows) {
    const noteUrl = String(row.notaUrl || '');
    if (!noteUrl) {
      continue;
    }

    if (/nfse\.salvador\.ba\.gov\.br/i.test(noteUrl) && row.emissao instanceof Date) {
      row.dataFatoGerador = row.emissao;
    }
  }
}

function buildGeneralTable(finalRows, selectedTables) {
  if (finalRows.length) {
    const headers = CANONICAL_COLUMNS.map((column) => column.title);
    const rows = finalRows.map((row) =>
      CANONICAL_COLUMNS.map((column) => formatPreviewCell(column.key, row[column.key], row)),
    );

    return {
      id: 'general-table',
      title: 'Base Padronizada',
      headers,
      rows,
      rowCount: rows.length,
      columnCount: headers.length
    };
  }

  const combined = buildCombinedTable(selectedTables);
  if (!combined) {
    return null;
  }

  return {
    id: 'general-table',
    title: 'Base Padronizada',
    headers: combined.headers,
    rows: combined.rows,
    rowCount: combined.rowCount,
    columnCount: combined.columnCount
  };
}

function buildRowsPreviewTable(rows, title) {
  if (!Array.isArray(rows) || !rows.length) {
    return null;
  }

  const headers = CANONICAL_COLUMNS.map((column) => column.title);
  const tableRows = rows.map((row) =>
    CANONICAL_COLUMNS.map((column) => formatPreviewCell(column.key, row[column.key], row)),
  );

  return {
    id: `${normalizeKey(title || 'planilha') || 'planilha'}-preview`,
    title: title || 'Planilha',
    headers,
    rows: tableRows,
    rowCount: tableRows.length,
    columnCount: headers.length,
  };
}

function documentNumbersMatch(left, right) {
  if (!left || !right) {
    return false;
  }

  if (normalizeKey(left) === normalizeKey(right)) {
    return true;
  }

  const leftDigits = trimLeadingZeros(digitsOnly(left));
  const rightDigits = trimLeadingZeros(digitsOnly(right));

  if (!leftDigits || !rightDigits) {
    return false;
  }

  if (leftDigits === rightDigits) {
    return true;
  }

  const shortestLength = Math.min(leftDigits.length, rightDigits.length);
  if (shortestLength < 4) {
    return false;
  }

  return leftDigits.endsWith(rightDigits) || rightDigits.endsWith(leftDigits);
}

function trimLeadingZeros(value) {
  return String(value || '').replace(/^0+/, '') || '0';
}

function formatPreviewCell(key, value, row = {}) {
  if (value == null || value === '') {
    if (key === 'dps') {
      return buildDerivedRpsNumber(row);
    }

    return '';
  }

  if (value instanceof Date) {
    const withTime = key === 'emissao';
    return formatDate(value, withTime);
  }

  if (typeof value === 'number') {
    return value.toLocaleString('pt-BR', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  }

  if (key === 'dps') {
    return buildDerivedRpsNumber(row);
  }

  if (key === 'arquivoPdf') {
    return path.basename(String(value));
  }

  return String(value);
}

function buildDerivedRpsNumber(row) {
  const source = String(row.rps || row.dps || '').trim();
  if (!source) {
    return '';
  }

  const match = source.match(/(\d{4,})$/);
  if (match?.[1]) {
    return match[1];
  }

  return source;
}

function filterAttachmentsBySelection(attachments, selectedAttachmentIds, selectionApplied = false) {
  const normalizedIds = normalizeSelectedIds(selectedAttachmentIds);
  if (selectionApplied && !normalizedIds.length) {
    return [];
  }

  if (!normalizedIds.length) {
    return attachments;
  }

  const selected = attachments.filter((attachment) => normalizedIds.includes(Number(attachment.id)));
  if (!selected.length) {
    throw new Error('Nenhum dos anexos selecionados foi encontrado neste e-mail.');
  }

  return selected;
}

function filterTablesBySelection(tables, selectedTableIds, selectionApplied = false, options = {}) {
  const normalizedIds = normalizeSelectedIds(selectedTableIds, false);
  if (selectionApplied && !normalizedIds.length && !options.allowEmptyWhenSelectionApplied) {
    return [];
  }

  if (!normalizedIds.length) {
    return tables;
  }

  const selected = tables.filter((table) => normalizedIds.includes(String(table.id)));
  if (!selected.length && !options.allowEmptyWhenSelectionApplied) {
    throw new Error('Nenhuma das tabelas selecionadas foi encontrada neste e-mail.');
  }

  return selected;
}

function normalizeSelectedIds(value, numeric = true) {
  const list = Array.isArray(value)
    ? value
    : value == null || value === ''
      ? []
      : [value];

  if (!numeric) {
    return list
      .map((item) => String(item || '').trim())
      .filter(Boolean);
  }

  return list
    .map((item) => Number(item))
    .filter((item) => Number.isInteger(item) && item > 0);
}

function alphanumericRichness(value) {
  return normalizeKey(value || '').replace(/\s+/g, '').length;
}

function prefixTableTitle(prefix, title) {
  const label = String(title || '').trim() || 'Tabela';
  return `${prefix} - ${label}`;
}

function buildComparisonEmailSummary(requestEmail, responseEmail) {
  return {
    subject: `Pedido: ${requestEmail.subject || '(sem assunto)'} | Retorno: ${responseEmail.subject || '(sem assunto)'}`,
    from: `Pedido: ${requestEmail.from || '-'} | Retorno: ${responseEmail.from || '-'}`,
    to: `Pedido: ${requestEmail.to || '-'} | Retorno: ${responseEmail.to || '-'}`,
    date: responseEmail.date || requestEmail.date || null,
    source: 'eml',
    sourcePath: `${requestEmail.sourcePath || 'pedido.eml'} | ${responseEmail.sourcePath || 'retorno.eml'}`,
    mailbox: 'comparativo',
    text: [
      'Pedido',
      requestEmail.text?.trim() || '(sem corpo em texto)',
      '',
      'Retorno',
      responseEmail.text?.trim() || '(sem corpo em texto)',
    ].join('\n\n'),
  };
}

module.exports = {
  previewComparisonSource,
  previewEmailSource,
  processEmailAutomation,
  processEmailComparison,
};

const { CANONICAL_COLUMNS } = require('../config/canonical-columns');
const {
  digitsOnly,
  normalizeKey,
  normalizeTextPreservingCase,
  safeNumberCompare,
} = require('../utils/normalizers');

function mergeSpreadsheetRowsWithInvoices(spreadsheetRows, invoices, warnings) {
  const options = arguments[3] || {};
  const appendUnmatchedInvoices = options.appendUnmatchedInvoices !== false;
  const warnUnmatchedRows = options.warnUnmatchedRows !== false;

  if (!Array.isArray(invoices) || !invoices.length) {
    return collapseDuplicateRows(spreadsheetRows)
      .map((row) => finalizeRow(row))
      .sort(compareRows);
  }

  const availableInvoices = invoices.map((invoice, index) => ({ ...invoice, __index: index }));
  const usedInvoices = new Set();

  const mergedRows = collapseDuplicateRows(spreadsheetRows).map((row) => {
    const match = findBestInvoiceMatch(row, availableInvoices, usedInvoices);
    if (!match) {
      if (warnUnmatchedRows) {
        warnings.push(buildUnmatchedSpreadsheetWarning(row));
      }

      return finalizeRow(row);
    }

    usedInvoices.add(match.__index);
    return finalizeRow(mergeRowWithInvoice(row, match));
  });

  for (const invoice of availableInvoices) {
    if (usedInvoices.has(invoice.__index)) {
      continue;
    }

    warnings.push(`PDF "${invoice.sourceFile}" nao encontrou linha correspondente na base.`);
    if (appendUnmatchedInvoices) {
      mergedRows.push(finalizeRow(invoiceToRow(invoice)));
    }
  }

  return collapseDuplicateRows(mergedRows).sort(compareRows);
}

function collapseDuplicateRows(rows) {
  const mergedRows = [];
  const signatureToIndex = new Map();

  for (const row of rows || []) {
    const signatures = buildDuplicateSignatures(row);
    const existingIndex = signatures
      .map((signature) => signatureToIndex.get(signature))
      .find((index) => Number.isInteger(index));

    if (!Number.isInteger(existingIndex)) {
      const nextIndex = mergedRows.length;
      mergedRows.push({ ...row });
      signatures.forEach((signature) => signatureToIndex.set(signature, nextIndex));
      continue;
    }

    const mergedRow = mergeDuplicatePair(mergedRows[existingIndex], row);
    mergedRows[existingIndex] = mergedRow;
    buildDuplicateSignatures(mergedRow).forEach((signature) => signatureToIndex.set(signature, existingIndex));
  }

  return mergedRows;
}

function buildDuplicateSignatures(row) {
  const signatures = [];

  for (const key of ['nfse', 'rps', 'dps']) {
    const normalized = normalizeDocumentIdentity(row[key]);
    if (normalized) {
      signatures.push(`${key}:${normalized}`);
    }
  }

  if (!signatures.length) {
    const fallback = [
      `cnpj:${normalizeDocumentIdentity(row.cnpj)}`,
      `tomador:${normalizeKey(row.tomador || '')}`,
      `emissao:${serializeDateIdentity(row.emissao)}`,
      `valor:${row.valorServico ?? ''}`,
    ].join('|');

    signatures.push(fallback);
  }

  return Array.from(new Set(signatures));
}

function mergeDuplicatePair(leftRow, rightRow) {
  const keepRight = rowRichnessScore(rightRow) > rowRichnessScore(leftRow);
  const primaryRow = keepRight ? { ...rightRow } : { ...leftRow };
  const secondaryRow = keepRight ? leftRow : rightRow;

  for (const column of CANONICAL_COLUMNS) {
    const key = column.key;
    const primaryValue = primaryRow[key];
    const secondaryValue = secondaryRow[key];

    if ((primaryValue == null || primaryValue === '') && secondaryValue != null && secondaryValue !== '') {
      primaryRow[key] = secondaryValue;
      continue;
    }

    if (shouldPreferSecondaryValue(key, primaryValue, secondaryValue)) {
      primaryRow[key] = secondaryValue;
    }
  }

  if (!primaryRow.arquivoPdf && secondaryRow.arquivoPdf) {
    primaryRow.arquivoPdf = secondaryRow.arquivoPdf;
  }

  if (!primaryRow.notaUrl && secondaryRow.notaUrl) {
    primaryRow.notaUrl = secondaryRow.notaUrl;
  }

  return primaryRow;
}

function rowRichnessScore(row) {
  let score = 0;

  for (const column of CANONICAL_COLUMNS) {
    if (row[column.key] != null && row[column.key] !== '') {
      score += 10;
    }
  }

  if (row.arquivoPdf) {
    score += 8;
  }

  if (row.notaUrl) {
    score += 6;
  }

  if (isNormalSituation(row.situacao)) {
    score += 5;
  }

  if (row.cnpj) {
    score += 4;
  }

  if (row.tomador) {
    score += 4;
  }

  return score;
}

function shouldPreferSecondaryValue(key, primaryValue, secondaryValue) {
  if (secondaryValue == null || secondaryValue === '') {
    return false;
  }

  if (key === 'situacao') {
    return !isNormalSituation(primaryValue) && isNormalSituation(secondaryValue);
  }

  if (key === 'rps') {
    return String(secondaryValue).length > String(primaryValue || '').length;
  }

  if (key === 'dps' || key === 'nfse' || key === 'numeroObra') {
    return String(secondaryValue).length > String(primaryValue || '').length;
  }

  if (key === 'tomador' || key === 'intermediario' || key === 'cartaDe') {
    return alphanumericRichness(secondaryValue) > alphanumericRichness(primaryValue);
  }

  if (key === 'cnpj') {
    return !primaryValue && Boolean(secondaryValue);
  }

  if (key === 'valorServico' || key === 'issDevido' || key === 'issPagar') {
    return (primaryValue == null || primaryValue === '') && secondaryValue != null && secondaryValue !== '';
  }

  return false;
}

function isNormalSituation(value) {
  return normalizeKey(value) === 'normal';
}

function findBestInvoiceMatch(row, invoices, usedInvoices) {
  let best = null;
  let bestScore = 0;

  for (const invoice of invoices) {
    if (usedInvoices.has(invoice.__index)) {
      continue;
    }

    const score = scoreMatch(row, invoice);
    if (score > bestScore) {
      bestScore = score;
      best = invoice;
    }
  }

  return bestScore >= 20 ? best : null;
}

function scoreMatch(row, invoice) {
  let score = 0;
  let documentMatches = 0;
  let corroboration = 0;

  if (documentNumbersMatch(row.nfse, invoice.nfse)) {
    score += 100;
    documentMatches += 1;
  }

  if (documentNumbersMatch(row.dps, invoice.dps)) {
    score += 60;
    documentMatches += 1;
  }

  if (documentNumbersMatch(row.rps, invoice.rps)) {
    score += 60;
    documentMatches += 1;
  }

  if (documentNumbersMatch(row.rps, invoice.dps)) {
    score += 55;
    documentMatches += 1;
  }

  if (documentNumbersMatch(row.dps, invoice.rps)) {
    score += 55;
    documentMatches += 1;
  }

  if (row.cnpj && invoice.cnpj && row.cnpj === invoice.cnpj) {
    score += 30;
    corroboration += 1;
  }

  if (row.tomador && invoice.tomador && normalizeKey(row.tomador) === normalizeKey(invoice.tomador)) {
    score += 25;
    corroboration += 1;
  }

  if (row.valorServico != null && invoice.valorServico != null && safeNumberCompare(row.valorServico, invoice.valorServico)) {
    score += 20;
    corroboration += 1;
  }

  if (row.emissao && invoice.emissao && sameCalendarDay(row.emissao, invoice.emissao)) {
    score += 10;
    corroboration += 1;
  }

  if (documentMatches > 0) {
    return score;
  }

  if (corroboration >= 3) {
    return score;
  }

  return 0;
}

function sameCalendarDay(left, right) {
  const leftDate = left instanceof Date ? left : new Date(left);
  const rightDate = right instanceof Date ? right : new Date(right);
  if (Number.isNaN(leftDate.getTime()) || Number.isNaN(rightDate.getTime())) {
    return false;
  }

  return leftDate.getFullYear() === rightDate.getFullYear()
    && leftDate.getMonth() === rightDate.getMonth()
    && leftDate.getDate() === rightDate.getDate();
}

function mergeRowWithInvoice(row, invoice) {
  const merged = { ...row };

  for (const column of CANONICAL_COLUMNS) {
    const key = column.key;
    if (invoice[key] != null && invoice[key] !== '') {
      merged[key] = invoice[key];
    }
  }

  if (!merged.arquivoPdf && invoice.arquivoPdf) {
    merged.arquivoPdf = invoice.arquivoPdf;
  }

  if (!merged.notaUrl && invoice.notaUrl) {
    merged.notaUrl = invoice.notaUrl;
  }

  return merged;
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

function normalizeDocumentIdentity(value) {
  if (!value) {
    return '';
  }

  const digits = trimLeadingZeros(digitsOnly(value));
  if (digits) {
    return digits;
  }

  return normalizeKey(value);
}

function serializeDateIdentity(value) {
  if (!(value instanceof Date) || Number.isNaN(value.getTime())) {
    return '';
  }

  return [
    value.getFullYear(),
    String(value.getMonth() + 1).padStart(2, '0'),
    String(value.getDate()).padStart(2, '0'),
  ].join('-');
}

function trimLeadingZeros(value) {
  return String(value || '').replace(/^0+/, '') || '0';
}

function invoiceToRow(invoice) {
  const row = {};

  for (const column of CANONICAL_COLUMNS) {
    row[column.key] = invoice[column.key] ?? '';
  }

  row.arquivoPdf = invoice.arquivoPdf ?? '';
  row.notaUrl = invoice.notaUrl ?? '';
  return row;
}

function finalizeRow(row) {
  const finalized = {};
  for (const column of CANONICAL_COLUMNS) {
    finalized[column.key] = row[column.key] ?? '';
  }

  if (!finalized.situacao) {
    finalized.situacao = 'Normal';
  }

  finalized.arquivoPdf = row.arquivoPdf ?? '';

  if (row.notaUrl) {
    finalized.notaUrl = row.notaUrl;
  }

  return finalized;
}

function compareRows(left, right) {
  const leftKey = Number(left.nfse || left.dps || left.rps || 0);
  const rightKey = Number(right.nfse || right.dps || right.rps || 0);
  return leftKey - rightKey;
}

function buildUnmatchedSpreadsheetWarning(row) {
  const identity = row.nfse || row.dps || row.cnpj || row.tomador || 'sem identificador';
  return `Linha base sem PDF correspondente: ${normalizeTextPreservingCase(identity)}.`;
}

function alphanumericRichness(value) {
  return normalizeKey(value || '').replace(/\s+/g, '').length;
}

module.exports = {
  collapseDuplicateRows,
  mergeSpreadsheetRowsWithInvoices,
};

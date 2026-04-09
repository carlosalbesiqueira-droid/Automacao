const { CANONICAL_COLUMNS } = require('../config/canonical-columns');
const { digitsOnly, formatDate, normalizeKey } = require('../utils/normalizers');

function pickRequestTable(tables) {
  const candidates = Array.isArray(tables) ? tables : [];
  let bestTable = null;
  let bestScore = 0;

  for (const table of candidates) {
    const score = scoreRequestTable(table);
    if (score > bestScore) {
      bestScore = score;
      bestTable = table;
    }
  }

  return bestScore > 0 ? bestTable : null;
}

function scoreRequestTable(table) {
  const headers = (table?.headers || []).map((header) => normalizeKey(header));
  if (!headers.length) {
    return 0;
  }

  let score = 0;
  if (headers.some((header) => header.includes('fatura'))) {
    score += 25;
  }
  if (headers.some((header) => header.includes('numero rps') || header === 'rps')) {
    score += 35;
  }
  if (headers.some((header) => header.includes('cliente') || header.includes('tomador'))) {
    score += 15;
  }
  if (headers.some((header) => header.includes('cnpj'))) {
    score += 10;
  }
  if (headers.some((header) => header.includes('valor'))) {
    score += 5;
  }
  if ((table?.rowCount || 0) > 0) {
    score += 5;
  }

  return score;
}

function buildPendingReport(requestedRows, proofRows) {
  const normalizedRequestedRows = Array.isArray(requestedRows) ? requestedRows : [];
  const normalizedProofRows = (Array.isArray(proofRows) ? proofRows : []).filter(hasGovernmentProof);

  const pendingRows = normalizedRequestedRows.filter((requestedRow) =>
    !normalizedProofRows.some((proofRow) => requestedRowHasProof(requestedRow, proofRow)),
  );

  return {
    pendingRows,
    pendingTable: buildPendingTable(pendingRows),
  };
}

function hasGovernmentProof(row) {
  return Boolean(row?.arquivoPdf || row?.notaUrl);
}

function requestedRowHasProof(requestedRow, proofRow) {
  if (documentNumbersMatch(requestedRow.nfse, proofRow.nfse)) {
    return true;
  }

  if (documentNumbersMatch(requestedRow.rps, proofRow.rps)) {
    return true;
  }

  if (documentNumbersMatch(requestedRow.rps, proofRow.dps)) {
    return true;
  }

  if (documentNumbersMatch(requestedRow.dps, proofRow.rps)) {
    return true;
  }

  if (documentNumbersMatch(requestedRow.dps, proofRow.dps)) {
    return true;
  }

  const sameClient =
    requestedRow.cnpj
    && proofRow.cnpj
    && normalizeKey(requestedRow.cnpj) === normalizeKey(proofRow.cnpj);
  const sameTomador =
    requestedRow.tomador
    && proofRow.tomador
    && normalizeKey(requestedRow.tomador) === normalizeKey(proofRow.tomador);

  if ((sameClient || sameTomador) && documentNumbersMatch(requestedRow.rps, proofRow.rps || proofRow.dps)) {
    return true;
  }

  return false;
}

function buildPendingTable(rows) {
  const headers = CANONICAL_COLUMNS.map((column) => column.title);
  const normalizedRows = (rows || []).map((row) =>
    CANONICAL_COLUMNS.map((column) => formatPendingCell(column.key, row[column.key], row)),
  );

  return {
    id: 'pending-table',
    title: 'Diferenciacao das pendencias',
    headers,
    rows: normalizedRows,
    rowCount: normalizedRows.length,
    columnCount: headers.length,
  };
}

function formatPendingCell(key, value, row) {
  if (value == null || value === '') {
    if (key === 'dps') {
      return buildDerivedRpsNumber(row);
    }

    return '';
  }

  if (value instanceof Date) {
    return formatDate(value, key === 'emissao');
  }

  if (typeof value === 'number') {
    return value.toLocaleString('pt-BR', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  }

  if (key === 'dps') {
    return buildDerivedRpsNumber(row);
  }

  return String(value);
}

function buildDerivedRpsNumber(row) {
  const source = String(row?.rps || row?.dps || '').trim();
  if (!source) {
    return '';
  }

  const match = source.match(/(\d{4,})$/);
  return match?.[1] || source;
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

module.exports = {
  buildPendingReport,
  pickRequestTable,
  requestedRowHasProof,
};

const path = require('node:path');
const ExcelJS = require('exceljs');
const { parse } = require('csv-parse/sync');
const { CANONICAL_COLUMNS, COLUMN_ALIASES } = require('../config/canonical-columns');
const {
  digitsOnly,
  formatCnpj,
  normalizeBooleanLike,
  normalizeDocumentNumber,
  normalizeKey,
  normalizeTextPreservingCase,
  normalizeUppercaseText,
  parseBrazilianDate,
  parseBrazilianMoney
} = require('../utils/normalizers');
const { resolveSupportedNoteUrl } = require('./note-link');

const CANONICAL_KEY_BY_ALIAS = buildCanonicalAliasMap();

async function normalizeSpreadsheetAttachment(attachment, warnings) {
  const extension = path.extname(attachment.filename || '').toLowerCase();

  if (extension === '.xls') {
    throw new Error(
      `O arquivo "${attachment.filename}" esta em formato .xls. Converta para .xlsx ou .csv para manter a automacao em uma trilha segura.`,
    );
  }

  const { matrix, worksheetName, rowMetadataByIndex } = await readSpreadsheetMatrix(attachment, extension);
  const rows = normalizeMatrixToCanonicalRows(matrix, warnings, {
    sourceLabel: attachment.filename,
    worksheetName,
    rowMetadataByIndex
  });

  warnings.push(
    `Planilha "${attachment.filename}" analisada com ${rows.length} linha(s) uteis na aba "${worksheetName}".`,
  );

  return rows;
}

function normalizeEmailTable(table, warnings) {
  const matrix = [];
  if (Array.isArray(table.headers) && table.headers.length) {
    matrix.push(table.headers);
  }

  for (const row of table.rows || []) {
    matrix.push(row);
  }

  const rows = normalizeMatrixToCanonicalRows(matrix, warnings, {
    sourceLabel: table.title || 'Tabela do e-mail',
    worksheetName: table.title || 'Tabela'
  });

  warnings.push(
    `Tabela "${table.title || 'Tabela'}" analisada com ${rows.length} linha(s) uteis no corpo do e-mail.`,
  );

  return rows;
}

function normalizeMatrixToCanonicalRows(matrix, warnings, options = {}) {
  const headerRowIndex = findHeaderRowIndex(matrix);
  const headerRow = matrix[headerRowIndex] ?? [];
  const headerMap = buildHeaderMap(headerRow);

  if (!Object.keys(headerMap).length) {
    warnings.push(
      `Nenhum cabecalho conhecido foi encontrado em ${options.sourceLabel || 'origem informada'}. A leitura foi feita por inferencia.`,
    );
  }

  const rows = [];
  for (let index = headerRowIndex + 1; index < matrix.length; index += 1) {
    const row = matrix[index];
    if (isEmptyRow(row) || isHeaderLikeRow(row)) {
      continue;
    }

    const normalizedRow = {};
    for (const column of CANONICAL_COLUMNS) {
      const cellIndex = headerMap[column.key];
      const rawValue = cellIndex == null ? '' : row[cellIndex];
      normalizedRow[column.key] = normalizeCellValue(column.key, rawValue);
    }

    const rowMetadata = options.rowMetadataByIndex?.[index];
    if (rowMetadata?.notaUrl) {
      normalizedRow.notaUrl = rowMetadata.notaUrl;
    }
    if (rowMetadata?.arquivoPdf) {
      normalizedRow.arquivoPdf = rowMetadata.arquivoPdf;
    }

    applyInlineRepAndNfHints(normalizedRow, row);
    if (isInlineHintOnlyRow(normalizedRow) && rows.length) {
      const previousRow = rows[rows.length - 1];
      if (!previousRow.nfse && normalizedRow.nfse) {
        previousRow.nfse = normalizedRow.nfse;
      }

      if (!previousRow.dps && normalizedRow.dps) {
        previousRow.dps = normalizedRow.dps;
      }

      if (!previousRow.rps && normalizedRow.rps) {
        previousRow.rps = normalizedRow.rps;
      }

      continue;
    }

    if (isComplementaryCompanyRow(normalizedRow, rows[rows.length - 1])) {
      mergeComplementaryCompanyRow(rows[rows.length - 1], normalizedRow);
      continue;
    }

    if (isMeaningfulRow(normalizedRow)) {
      rows.push(normalizedRow);
    }
  }

  return rows;
}

async function readSpreadsheetMatrix(attachment, extension) {
  if (extension === '.csv') {
    return {
      matrix: readCsvMatrix(attachment.content),
      worksheetName: 'CSV'
    };
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(attachment.content);

  const bestWorksheet = pickWorksheet(workbook);
  if (!bestWorksheet) {
    throw new Error(`Nao foi possivel localizar uma aba com dados em ${attachment.filename}.`);
  }

  return {
    ...worksheetToMatrix(bestWorksheet),
    worksheetName: bestWorksheet.name
  };
}

function readCsvMatrix(buffer) {
  return parse(buffer, {
    bom: true,
    columns: false,
    relax_column_count: true,
    skip_empty_lines: true
  });
}

function worksheetToMatrix(worksheet) {
  const rows = [];
  const rowMetadataByIndex = {};
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    const values = row.values.slice(1).map(normalizeExcelValue);
    while (values.length && !normalizeTextPreservingCase(values[values.length - 1])) {
      values.pop();
    }

    const matrixIndex = rows.length;
    rows.push(values);

    const hyperlinks = [];
    row.eachCell({ includeEmpty: false }, (cell) => {
      const hyperlink = cell.hyperlink || (typeof cell.value === 'object' ? cell.value?.hyperlink : '');
      if (hyperlink) {
        hyperlinks.push(normalizeTextPreservingCase(hyperlink));
      }
    });

    const notaUrl = pickOfficialNoteUrl(hyperlinks);
    const arquivoPdf = pickPdfLink(hyperlinks);
    if (notaUrl || hyperlinks.length) {
      rowMetadataByIndex[matrixIndex] = {
        notaUrl,
        arquivoPdf,
        links: hyperlinks
      };
    }
  });

  return {
    matrix: rows,
    rowMetadataByIndex
  };
}

function normalizeExcelValue(value) {
  if (value == null) {
    return '';
  }

  if (value instanceof Date) {
    return value;
  }

  if (typeof value === 'object') {
    if (value.text) {
      return value.text;
    }

    if (Array.isArray(value.richText)) {
      return value.richText.map((item) => item.text).join('');
    }

    if (value.result != null) {
      return value.result;
    }
  }

  return value;
}

function buildCanonicalAliasMap() {
  const map = new Map();
  for (const [canonicalKey, aliases] of Object.entries(COLUMN_ALIASES)) {
    for (const alias of aliases) {
      map.set(normalizeKey(alias), canonicalKey);
    }
  }

  return map;
}

function pickWorksheet(workbook) {
  let bestSheet = null;
  let bestCount = -1;

  workbook.eachSheet((worksheet) => {
    const { matrix: rows } = worksheetToMatrix(worksheet);
    const count = rows.reduce((total, row) => total + (isEmptyRow(row) ? 0 : 1), 0);
    if (count > bestCount) {
      bestCount = count;
      bestSheet = worksheet;
    }
  });

  return bestSheet;
}

function findHeaderRowIndex(rows) {
  let bestIndex = 0;
  let bestScore = -1;

  for (let index = 0; index < Math.min(rows.length, 30); index += 1) {
    const score = countHeaderMatches(rows[index]);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  }

  return bestIndex;
}

function countHeaderMatches(row) {
  return (row ?? []).reduce((total, cell) => total + (resolveCanonicalKey(cell) ? 1 : 0), 0);
}

function buildHeaderMap(row) {
  const map = {};
  (row ?? []).forEach((cell, index) => {
    const canonicalKey = resolveCanonicalKey(cell);
    if (canonicalKey && map[canonicalKey] == null) {
      map[canonicalKey] = index;
    }
  });
  return map;
}

function resolveCanonicalKey(value) {
  const normalized = normalizeKey(value);
  return CANONICAL_KEY_BY_ALIAS.get(normalized) ?? null;
}

function isHeaderLikeRow(row) {
  return countHeaderMatches(row) >= 3;
}

function isEmptyRow(row) {
  return !(row ?? []).some((cell) => normalizeTextPreservingCase(cell));
}

function normalizeCellValue(key, rawValue) {
  switch (key) {
    case 'nfse':
    case 'dps':
      return digitsOnly(rawValue) ? normalizeDocumentNumber(rawValue) : '';
    case 'rps':
      return digitsOnly(rawValue) ? normalizeDocumentNumber(rawValue) : normalizeTextPreservingCase(rawValue);
    case 'intermediario':
      return normalizeIntermediarioValue(rawValue);
    case 'emissao':
    case 'dataFatoGerador':
      return parseBrazilianDate(rawValue);
    case 'tomador':
      return normalizeUppercaseText(rawValue);
    case 'cnpj':
    case 'prestadorCnpj':
      return formatCnpj(rawValue);
    case 'valorServico':
    case 'valorDeducao':
    case 'issDevido':
    case 'issPagar':
    case 'valorCredito':
      return parseBrazilianMoney(rawValue);
    case 'issRetido':
    case 'issPagoGuia':
      return normalizeBooleanLike(rawValue);
    case 'situacao':
      return normalizeTextPreservingCase(rawValue) || 'Normal';
    case 'cartaDe':
      return normalizeLooseTextValue(rawValue, { zeroMeansBlank: true });
    case 'numeroObra':
      return normalizeLooseTextValue(rawValue, { zeroMeansBlank: true });
    default:
      return normalizeTextPreservingCase(rawValue);
  }
}

function applyInlineRepAndNfHints(normalizedRow, sourceRow) {
  const joinedText = (sourceRow ?? []).map((cell) => normalizeTextPreservingCase(cell)).join(' | ');

  const repMatch = joinedText.match(/(?:^|\b)rep\s*[:#-]?\s*([a-z0-9.\/-]+)/i);
  if (!normalizedRow.nfse && repMatch?.[1]) {
    normalizedRow.nfse = normalizeDocumentNumber(repMatch[1]);
  }

  const nfMatch = joinedText.match(/(?:^|\b)nf\s*[:#-]?\s*([a-z0-9.\/-]+)/i);
  if (!normalizedRow.dps && nfMatch?.[1]) {
    normalizedRow.dps = normalizeDocumentNumber(nfMatch[1]);
  }
}

function isMeaningfulRow(row) {
  return Boolean(
    row.nfse ||
      row.rps ||
      row.dps ||
      row.tomador ||
      row.cnpj ||
      row.valorServico != null,
  );
}

function isInlineHintOnlyRow(row) {
  const hasHintLikeTomador = /^(nf|rep)\b/i.test(row.tomador || '');
  const relevantKeys = ['cnpj', 'prestadorCnpj', 'valorServico', 'emissao', 'dataFatoGerador', 'intermediario'];
  const hasRichData = relevantKeys.some((key) => row[key] != null && row[key] !== '');
  const identityFields = ['nfse', 'rps', 'dps'].filter((key) => row[key]);

  if (!hasRichData && identityFields.length > 0 && (!row.tomador || hasHintLikeTomador)) {
    return true;
  }

  return false;
}

function isComplementaryCompanyRow(row, previousRow) {
  if (!previousRow) {
    return false;
  }

  if (normalizeKey(row.situacao) !== 'cancelar') {
    return false;
  }

  if (!sharesDocumentIdentity(row, previousRow)) {
    return false;
  }

  if (looksLikeCompanyIdentifier(row.tomador)) {
    return true;
  }

  if (row.cnpj && !previousRow.cnpj) {
    return true;
  }

  return false;
}

function mergeComplementaryCompanyRow(previousRow, row) {
  if (!previousRow.cnpj) {
    const cnpjFromTomador = extractCnpjFromIdentifier(row.tomador);
    if (cnpjFromTomador) {
      previousRow.cnpj = cnpjFromTomador;
    } else if (row.cnpj) {
      previousRow.cnpj = row.cnpj;
    }
  }

  for (const key of ['nfse', 'rps', 'dps', 'emissao', 'dataFatoGerador', 'prestadorCnpj', 'intermediario', 'valorServico', 'valorDeducao', 'issDevido', 'issPagar', 'valorCredito', 'issRetido', 'issPagoGuia', 'cartaDe', 'numeroObra']) {
    if ((previousRow[key] == null || previousRow[key] === '') && row[key] != null && row[key] !== '') {
      previousRow[key] = row[key];
    }
  }

  if (!previousRow.notaUrl && row.notaUrl) {
    previousRow.notaUrl = row.notaUrl;
  }
}

function sharesDocumentIdentity(row, previousRow) {
  const documentKeys = ['nfse', 'rps', 'dps'];
  for (const key of documentKeys) {
    if (documentIdentityEquals(row[key], previousRow[key])) {
      return true;
    }
  }

  return false;
}

function documentIdentityEquals(left, right) {
  const leftText = normalizeTextPreservingCase(left);
  const rightText = normalizeTextPreservingCase(right);

  if (!leftText || !rightText) {
    return false;
  }

  if (leftText === rightText) {
    return true;
  }

  const leftDigits = trimIdentityDigits(digitsOnly(leftText));
  const rightDigits = trimIdentityDigits(digitsOnly(rightText));
  if (!leftDigits || !rightDigits) {
    return false;
  }

  return leftDigits === rightDigits;
}

function trimIdentityDigits(value) {
  return String(value || '').replace(/^0+/, '') || '0';
}

function looksLikeCompanyIdentifier(value) {
  const text = normalizeTextPreservingCase(value);
  if (!text) {
    return false;
  }

  if (extractCnpjFromIdentifier(text)) {
    return true;
  }

  return /^inscri[cç][aã]o\s*:/i.test(text);
}

function extractCnpjFromIdentifier(value) {
  const text = normalizeTextPreservingCase(value);
  const digits = digitsOnly(text);
  if (digits.length !== 14) {
    return '';
  }

  return formatCnpj(digits);
}

function normalizeIntermediarioValue(value) {
  const text = normalizeTextPreservingCase(value);
  if (!text || /^-+$/.test(text) || /^nao identificado$/i.test(normalizeKey(text))) {
    return '';
  }

  const digits = digitsOnly(text);
  if (digits.length === 14) {
    return formatCnpj(digits);
  }

  if (digits && digits === text.replace(/\s+/g, '')) {
    if (digits === '0' || digits.length <= 2) {
      return '';
    }

    return digits;
  }

  if (normalizeKey(text) === 'atividade') {
    return '';
  }

  return normalizeUppercaseText(text);
}

function normalizeLooseTextValue(value, options = {}) {
  const text = normalizeTextPreservingCase(value);
  if (!text || /^-+$/.test(text)) {
    return '';
  }

  if (options.zeroMeansBlank && text === '0') {
    return '';
  }

  return text;
}

function pickOfficialNoteUrl(links) {
  const candidates = Array.isArray(links) ? links : [];
  for (const link of candidates) {
    const resolved = resolveSupportedNoteUrl(link);
    if (resolved) {
      return resolved;
    }
  }

  return '';
}

function pickPdfLink(links) {
  const candidates = Array.isArray(links) ? links : [];
  return candidates.find((link) => /\.pdf(?:$|[?#])/i.test(link) || (/^[a-z]:\\/i.test(link) && /\.pdf$/i.test(link))) || '';
}

module.exports = {
  normalizeEmailTable,
  normalizeMatrixToCanonicalRows,
  normalizeSpreadsheetAttachment
};

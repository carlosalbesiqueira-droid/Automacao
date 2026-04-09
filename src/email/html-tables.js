const { normalizeTextPreservingCase } = require('../utils/normalizers');

function extractHtmlTables(html) {
  const source = String(html || '');
  if (!source.trim()) {
    return [];
  }

  const tables = [];
  const matches = source.matchAll(/<table\b[^>]*>[\s\S]*?<\/table>/gi);

  for (const [index, match] of Array.from(matches).entries()) {
    const parsed = parseTable(match[0], index + 1);
    if (parsed) {
      tables.push(parsed);
    }
  }

  return tables;
}

function parseTable(tableHtml, index) {
  const rowMatches = Array.from(tableHtml.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi));
  if (!rowMatches.length) {
    return null;
  }

  const rows = rowMatches
    .map((match) => parseRow(match[1]))
    .filter((row) => row.cells.length);

  if (!rows.length) {
    return null;
  }

  const caption = cleanCellText(capture(tableHtml, /<caption\b[^>]*>([\s\S]*?)<\/caption>/i));
  const headerRowIndex = rows.findIndex((row) => row.hasHeader);

  let headers = [];
  let dataRows = rows;

  if (headerRowIndex >= 0) {
    headers = rows[headerRowIndex].cells;
    dataRows = rows.filter((_, rowIndex) => rowIndex !== headerRowIndex);
  } else if (rows.length > 1 && rows[0].cells.length > 1) {
    headers = rows[0].cells;
    dataRows = rows.slice(1);
  }

  const columnCount = Math.max(
    headers.length,
    ...dataRows.map((row) => row.cells.length),
    0,
  );

  if (!columnCount) {
    return null;
  }

  const normalizedHeaders = normalizeRow(
    headers.length ? headers : Array.from({ length: columnCount }, (_, itemIndex) => `Coluna ${itemIndex + 1}`),
    columnCount,
  );

  const normalizedRows = dataRows.map((row) => normalizeRow(row.cells, columnCount));

  return {
    id: `table-${index}`,
    title: caption || `Tabela ${index}`,
    headers: normalizedHeaders,
    rows: normalizedRows,
    rowCount: normalizedRows.length,
    columnCount
  };
}

function parseRow(rowHtml) {
  const cells = [];
  const matches = rowHtml.matchAll(/<(td|th)\b[^>]*>([\s\S]*?)<\/\1>/gi);

  for (const match of matches) {
    cells.push(cleanCellText(match[2]));
  }

  return {
    cells,
    hasHeader: /<th\b/i.test(rowHtml)
  };
}

function normalizeRow(cells, columnCount) {
  const row = Array.from({ length: columnCount }, (_, index) => normalizeTextPreservingCase(cells[index] || ''));
  return row;
}

function cleanCellText(value) {
  return normalizeTextPreservingCase(
    decodeHtmlEntities(
      String(value || '')
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/p>/gi, '\n')
        .replace(/<[^>]+>/g, ' ')
        .replace(/\s*\n\s*/g, '\n')
    )
  );
}

function decodeHtmlEntities(value) {
  const namedEntities = {
    nbsp: ' ',
    amp: '&',
    lt: '<',
    gt: '>',
    quot: '"',
    apos: "'",
    ordm: 'ª',
    ordf: 'º',
    aacute: 'á',
    Aacute: 'Á',
    eacute: 'é',
    Eacute: 'É',
    iacute: 'í',
    Iacute: 'Í',
    oacute: 'ó',
    Oacute: 'Ó',
    uacute: 'ú',
    Uacute: 'Ú',
    agrave: 'à',
    Agrave: 'À',
    atilde: 'ã',
    Atilde: 'Ã',
    otilde: 'õ',
    Otilde: 'Õ',
    acirc: 'â',
    Acirc: 'Â',
    ecirc: 'ê',
    Ecirc: 'Ê',
    ocirc: 'ô',
    Ocirc: 'Ô',
    ccedil: 'ç',
    Ccedil: 'Ç'
  };

  return String(value || '')
    .replace(/&#(\d+);/g, (_match, code) => String.fromCharCode(Number(code)))
    .replace(/&#x([0-9a-f]+);/gi, (_match, code) => String.fromCharCode(parseInt(code, 16)))
    .replace(/&([a-zA-Z]+);/g, (match, name) => namedEntities[name] ?? match);
}

function capture(text, pattern) {
  return text.match(pattern)?.[1] ?? '';
}

module.exports = {
  extractHtmlTables
};

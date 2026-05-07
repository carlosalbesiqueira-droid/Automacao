const { normalizeTextPreservingCase } = require('../utils/normalizers');

function buildCombinedTable(tables) {
  const sourceTables = Array.isArray(tables) ? tables : [];
  if (!sourceTables.length) {
    return null;
  }

  const headers = [];
  const seenHeaders = new Set();

  for (const table of sourceTables) {
    for (const header of table.headers || []) {
      const normalizedHeader = normalizeTextPreservingCase(header);
      if (!normalizedHeader || seenHeaders.has(normalizedHeader)) {
        continue;
      }

      seenHeaders.add(normalizedHeader);
      headers.push(normalizedHeader);
    }
  }

  const fallbackHeaders = headers.length
    ? headers
    : Array.from({ length: Math.max(...sourceTables.map((table) => table.columnCount || 0), 0) }, (_item, index) => `Coluna ${index + 1}`);

  const rows = [];

  for (const table of sourceTables) {
    const tableHeaders = normalizeHeaders(table.headers, table.columnCount);
    for (const row of table.rows || []) {
      const outputRow = fallbackHeaders.map(() => '');

      fallbackHeaders.forEach((header, headerIndex) => {
        const sourceIndex = tableHeaders.indexOf(header);
        outputRow[headerIndex] = sourceIndex >= 0 ? normalizeTextPreservingCase(row[sourceIndex] || '') : '';
      });

      rows.push(outputRow);
    }
  }

  return {
    id: 'combined-table',
    title: 'Tabela Geral',
    headers: fallbackHeaders,
    rows,
    rowCount: rows.length,
    columnCount: fallbackHeaders.length
  };
}

function normalizeHeaders(headers = [], columnCount = 0) {
  const baseHeaders = Array.isArray(headers) ? headers : [];
  const width = Math.max(baseHeaders.length, columnCount, 0);

  return Array.from({ length: width }, (_item, index) =>
    normalizeTextPreservingCase(baseHeaders[index] || `Coluna ${index + 1}`),
  );
}

module.exports = {
  buildCombinedTable,
  normalizeHeaders
};

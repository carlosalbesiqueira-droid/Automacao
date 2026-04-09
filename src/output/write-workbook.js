const path = require('node:path');
const ExcelJS = require('exceljs');
const { CANONICAL_COLUMNS } = require('../config/canonical-columns');
const { ensureDirectory, makeTimestamp } = require('../utils/fs');
const { formatDate } = require('../utils/normalizers');

async function writeOutputWorkbook({
  outputDirectory,
  email,
  rows,
  pendingRows = [],
  warnings,
  selectedTables = [],
  generalTable = null,
  pendingTable = null,
}) {
  await ensureDirectory(outputDirectory);

  const timestamp = makeTimestamp();
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Codex';
  workbook.created = new Date();

  if (Array.isArray(rows) && rows.length) {
    writeNormalizedSheet(workbook, rows, 'Planilha Padronizada');
  }

  if (Array.isArray(pendingRows) && pendingRows.length) {
    writeNormalizedSheet(workbook, pendingRows, 'Pendencias Remanescentes');
  }

  selectedTables.forEach((table, index) => {
    writeTableSheet(workbook, table, `Tabela ${index + 1}`);
  });

  writeEmailSummarySheet(workbook, email);
  writeWarningsSheet(workbook, warnings);

  const workbookFileName = `planilha_padronizada_${timestamp}.xlsx`;
  const workbookPath = path.join(outputDirectory, workbookFileName);
  await workbook.xlsx.writeFile(workbookPath);

  let generalFilePath = '';
  if ((Array.isArray(rows) && rows.length) || (generalTable && generalTable.headers.length)) {
    const generalWorkbook = new ExcelJS.Workbook();
    generalWorkbook.creator = 'Codex';
    generalWorkbook.created = new Date();

    if (Array.isArray(rows) && rows.length) {
      writeNormalizedSheet(generalWorkbook, rows, 'Base Padronizada');
    } else if (generalTable && Array.isArray(generalTable.headers) && generalTable.headers.length) {
      writeTableSheet(generalWorkbook, generalTable, 'Base Padronizada');
    }

    const generalFileName = `arquivo_geral_${timestamp}.xlsx`;
    generalFilePath = path.join(outputDirectory, generalFileName);
    await generalWorkbook.xlsx.writeFile(generalFilePath);
  }

  let pendingFilePath = '';
  if ((Array.isArray(pendingRows) && pendingRows.length) || (pendingTable && pendingTable.headers.length)) {
    const pendingWorkbook = new ExcelJS.Workbook();
    pendingWorkbook.creator = 'Codex';
    pendingWorkbook.created = new Date();

    if (Array.isArray(pendingRows) && pendingRows.length) {
      writeNormalizedSheet(pendingWorkbook, pendingRows, 'Pendencias Remanescentes');
    } else if (pendingTable && Array.isArray(pendingTable.headers) && pendingTable.headers.length) {
      writeTableSheet(pendingWorkbook, pendingTable, 'Pendencias Remanescentes');
    }

    const pendingFileName = `pendencias_remanescentes_${timestamp}.xlsx`;
    pendingFilePath = path.join(outputDirectory, pendingFileName);
    await pendingWorkbook.xlsx.writeFile(pendingFilePath);
  }

  return {
    workbookPath,
    generalFilePath,
    pendingFilePath,
  };
}

function writeNormalizedSheet(workbook, rows, sheetName = 'Planilha Padronizada') {
  const worksheet = workbook.addWorksheet(sheetName, {
    views: [{ state: 'frozen', ySplit: 1 }]
  });

  worksheet.columns = CANONICAL_COLUMNS.map((column) => ({
    header: column.title,
    key: column.key,
    width: column.width
  }));

  rows.forEach((row) => {
    const excelRow = worksheet.addRow(buildExcelRow(row));
    applyCanonicalRowFormatting(excelRow, row);
  });

  worksheet.autoFilter = {
    from: 'A1',
    to: `${columnNumberToName(CANONICAL_COLUMNS.length)}1`
  };

  styleHeaderRow(worksheet.getRow(1));
  applyBodyBorder(worksheet);
}

function writeTableSheet(workbook, table, fallbackName) {
  const sheetName = uniqueWorksheetName(workbook, sanitizeSheetName(table.title || fallbackName || 'Tabela'));
  const worksheet = workbook.addWorksheet(sheetName, {
    views: [{ state: 'frozen', ySplit: 1 }]
  });

  const headers = Array.isArray(table.headers) ? table.headers : [];
  if (!headers.length) {
    headers.push('Coluna 1');
  }

  worksheet.columns = headers.map((header) => ({
    header,
    key: header,
    width: Math.min(Math.max(String(header || '').length + 6, 14), 28)
  }));

  (table.rows || []).forEach((row) => {
    worksheet.addRow(headers.map((_header, index) => row[index] ?? ''));
  });

  worksheet.autoFilter = {
    from: 'A1',
    to: `${columnNumberToName(headers.length)}1`
  };

  styleHeaderRow(worksheet.getRow(1));
  applyBodyBorder(worksheet);
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      return;
    }

    row.alignment = { vertical: 'top', wrapText: true };
  });
}

function writeEmailSummarySheet(workbook, email) {
  const worksheet = workbook.addWorksheet('Resumo Email');
  const rows = [
    ['Assunto', email.subject || ''],
    ['Remetente', email.from || ''],
    ['Destinatario', email.to || ''],
    ['Data', email.date ? formatDate(new Date(email.date), true) : ''],
    ['Origem', email.source === 'imap' ? `IMAP/${email.mailbox}` : email.sourcePath || ''],
    ['Corpo do e-mail', email.text?.trim() || '']
  ];

  rows.forEach((values) => worksheet.addRow(values));
  worksheet.columns = [
    { width: 20 },
    { width: 120 }
  ];

  worksheet.getColumn(2).alignment = { wrapText: true, vertical: 'top' };
  worksheet.eachRow((row, rowNumber) => {
    row.getCell(1).font = { bold: true };
    if (rowNumber === rows.length) {
      row.height = 120;
    }
  });
}

function writeWarningsSheet(workbook, warnings) {
  const worksheet = workbook.addWorksheet('Avisos');
  worksheet.columns = [
    { header: 'Aviso', key: 'warning', width: 120 }
  ];
  warnings.forEach((warning) => worksheet.addRow({ warning }));
  worksheet.getRow(1).font = { bold: true };
}

function buildExcelRow(row) {
  const output = {};

  for (const column of CANONICAL_COLUMNS) {
    if (column.key === 'nfse' && row[column.key] && (row.notaUrl || row.arquivoPdf)) {
      output[column.key] = {
        text: String(row[column.key]),
        hyperlink: normalizeHyperlinkTarget(row.notaUrl || row.arquivoPdf)
      };
      continue;
    }

    if (column.key === 'dps') {
      output[column.key] = buildDerivedRpsNumber(row);
      continue;
    }

    if ((column.type === 'date' || column.type === 'dateTime') && row[column.key] instanceof Date) {
      output[column.key] = formatDate(row[column.key], column.type === 'dateTime');
      continue;
    }

    if (column.type === 'link' && row[column.key]) {
      const targetPath = normalizeHyperlinkTarget(row[column.key]);
      output[column.key] = {
        text: path.basename(String(row[column.key])),
        hyperlink: targetPath
      };
      continue;
    }

    output[column.key] = row[column.key] ?? '';
  }

  return output;
}

function applyCanonicalRowFormatting(excelRow, sourceRow) {
  CANONICAL_COLUMNS.forEach((column, index) => {
    const cell = excelRow.getCell(index + 1);
    if (
      column.key === 'rps'
      || column.key === 'dps'
      || column.key === 'cnpj'
      || column.key === 'nfse'
      || column.type === 'date'
      || column.type === 'dateTime'
    ) {
      cell.numFmt = '@';
    }

    if (column.type === 'currency' && typeof sourceRow[column.key] === 'number') {
      cell.numFmt = '#,##0.00';
    }

    if (cell.value && typeof cell.value === 'object' && cell.value.hyperlink) {
      cell.font = {
        color: { argb: 'FF0563C1' },
        underline: true,
      };
    }
  });
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

function styleHeaderRow(headerRow) {
  headerRow.height = 24;
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF365F7C' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = border();
  });
}

function applyBodyBorder(worksheet) {
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      return;
    }

    row.eachCell((cell) => {
      cell.border = border('FFD9E2F3');
      cell.alignment = { vertical: 'middle' };
    });
  });
}

function uniqueWorksheetName(workbook, baseName) {
  const existingNames = new Set(workbook.worksheets.map((worksheet) => worksheet.name));
  if (!existingNames.has(baseName)) {
    return baseName;
  }

  let counter = 2;
  while (existingNames.has(`${baseName} ${counter}`)) {
    counter += 1;
  }

  return `${baseName} ${counter}`;
}

function sanitizeSheetName(value) {
  return String(value || 'Tabela')
    .replace(/[\\/*?:[\]]/g, ' ')
    .trim()
    .slice(0, 31) || 'Tabela';
}

function border(color = 'FFB7C9D6') {
  return {
    top: { style: 'thin', color: { argb: color } },
    left: { style: 'thin', color: { argb: color } },
    bottom: { style: 'thin', color: { argb: color } },
    right: { style: 'thin', color: { argb: color } }
  };
}

function normalizeHyperlinkTarget(value) {
  const text = String(value || '').trim();
  if (!text) {
    return '';
  }

  if (/^(https?:|mailto:|file:)/i.test(text)) {
    return text;
  }

  if (/^[a-z]:\\/i.test(text)) {
    return encodeURI(`file:///${text.replace(/\\/g, '/')}`);
  }

  return text;
}

function columnNumberToName(index) {
  let dividend = index;
  let columnName = '';

  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return columnName;
}

module.exports = {
  writeOutputWorkbook
};

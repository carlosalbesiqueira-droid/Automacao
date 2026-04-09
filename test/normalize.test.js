const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const { normalizeSpreadsheetAttachment } = require('../src/parsers/spreadsheet');

test('normaliza planilha com cabecalhos repetidos e mapeia Rep/NF', async () => {
  const rows = [
    ['Relatorio Mensal de Notas'],
    ['Rep', 'Cliente', 'CNPJ', 'Valor Total'],
    ['17915', 'BELGO BEKAERT ARAMES LTDA', '61074506002698', '424,72'],
    ['Rep', 'Cliente', 'CNPJ', 'Valor Total'],
    ['18000', 'ARCELORMITTAL BRASIL S.A.', '174690701008313', '612,88'],
    ['Observacao', 'NF 17838']
  ];

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Dados');
  rows.forEach((row) => worksheet.addRow(row));
  const content = await workbook.xlsx.writeBuffer();

  const warnings = [];
  const result = await normalizeSpreadsheetAttachment({ filename: 'teste.xlsx', content }, warnings);

  assert.equal(result.length, 2);
  assert.equal(result[0].nfse, '17915');
  assert.equal(result[0].tomador, 'BELGO BEKAERT ARAMES LTDA');
  assert.equal(result[0].cnpj, '61.074.506/0026-98');
  assert.equal(result[0].valorServico, 424.72);
});

test('une linhas complementares com nome e identificador do tomador', async () => {
  const rows = [
    ['NFS-e', 'RPS', 'Emissao', 'Data Fato Gerador', 'Tomador de Servicos', 'Valor Servico', 'Situacao'],
    ['1470608', '00U1.00472929', '23/03/2026 14:49', '18/03/2026', 'ARCELORMITTAL BRASIL S.A.', '586,37', 'Normal'],
    ['1470608', '00U1.00472929', '23/03/2026 14:49', '18/03/2026', 'Inscricao: 3.251.805-6', '586,37', 'CANCELAR'],
    ['1470607', '00U1.00472928', '23/03/2026 14:49', '18/03/2026', 'ARCELORMITTAL BRASIL S.A.', '1176,58', 'Normal'],
    ['1470607', '00U1.00472928', '23/03/2026 14:49', '18/03/2026', '17.469.701/0083-13', '1176,58', 'CANCELAR']
  ];

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Dados');
  rows.forEach((row) => worksheet.addRow(row));
  const content = await workbook.xlsx.writeBuffer();

  const warnings = [];
  const result = await normalizeSpreadsheetAttachment({ filename: 'duplicada.xlsx', content }, warnings);

  assert.equal(result.length, 2);
  assert.equal(result[0].nfse, '1470608');
  assert.equal(result[0].rps, '00U1.00472929');
  assert.equal(result[0].tomador, 'ARCELORMITTAL BRASIL S.A.');
  assert.equal(result[0].situacao, 'Normal');
  assert.equal(result[0].cnpj, '');
  assert.equal(result[1].nfse, '1470607');
  assert.equal(result[1].rps, '00U1.00472928');
  assert.equal(result[1].cnpj, '17.469.701/0083-13');
  assert.equal(result[1].situacao, 'Normal');
});

test('preserva zeros a esquerda em campos de documento da planilha', async () => {
  const rows = [
    ['NFS-e', 'RPS', 'CNPJ', 'Tomador de Servicos', 'Situacao'],
    ['00017915', 'U1.00017864', '01746970100831', 'ARCELORMITTAL BRASIL S.A.', 'Normal']
  ];

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Dados');
  rows.forEach((row) => worksheet.addRow(row));
  const content = await workbook.xlsx.writeBuffer();

  const warnings = [];
  const result = await normalizeSpreadsheetAttachment({ filename: 'zeros.xlsx', content }, warnings);

  assert.equal(result.length, 1);
  assert.equal(result[0].nfse, '00017915');
  assert.equal(result[0].rps, 'U1.00017864');
  assert.equal(result[0].cnpj, '01.746.970/1008-31');
});

test('carrega o hyperlink da NF a partir da planilha anexada', async () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Dados');
  worksheet.addRow(['NFS-e', 'RPS', 'Tomador de Servicos']);
  const row = worksheet.addRow(['1470608', '00U1.00472929', 'ARCELORMITTAL BRASIL S.A.']);
  row.getCell(1).value = {
    text: '1470608',
    hyperlink: 'https://nfe.prefeitura.sp.gov.br/contribuinte/notaprint.aspx?inscricao=27143449&nf=1470608&verificacao=U6WUINR6'
  };

  const content = await workbook.xlsx.writeBuffer();
  const warnings = [];
  const result = await normalizeSpreadsheetAttachment({ filename: 'links.xlsx', content }, warnings);

  assert.equal(result.length, 1);
  assert.equal(result[0].nfse, '1470608');
  assert.equal(result[0].notaUrl, 'https://nfe.prefeitura.sp.gov.br/contribuinte/notaprint.aspx?inscricao=27143449&nf=1470608&verificacao=U6WUINR6');
});

test('detecta link oficial mesmo quando ele esta fora da coluna da NF', async () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Dados');
  worksheet.addRow(['NFS-e', 'RPS', 'Tomador de Servicos', 'Observacao']);
  const row = worksheet.addRow(['1470609', '00U1.00472930', 'ARCELORMITTAL BRASIL S.A.', 'Abrir nota']);
  row.getCell(4).value = {
    text: 'Abrir nota',
    hyperlink: 'https://nfe.prefeitura.sp.gov.br/contribuinte/notaprint.aspx?inscricao=27143449&nf=1470609&verificacao=ABCDEFGH'
  };

  const content = await workbook.xlsx.writeBuffer();
  const warnings = [];
  const result = await normalizeSpreadsheetAttachment({ filename: 'links-em-outra-coluna.xlsx', content }, warnings);

  assert.equal(result.length, 1);
  assert.equal(result[0].nfse, '1470609');
  assert.equal(result[0].notaUrl, 'https://nfe.prefeitura.sp.gov.br/contribuinte/notaprint.aspx?inscricao=27143449&nf=1470609&verificacao=ABCDEFGH');
});

test('preserva o codigo de atividade quando a planilha usa esse campo no lugar do intermediario', async () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Dados');
  worksheet.addRow(['NFS-e', 'RPS', 'Tomador de Servicos', 'Atividade', 'Valor Servicos']);
  worksheet.addRow(['17915', 'U1.00017838', 'BELGO BEKAERT ARAMES LTDA', '3101', '424,72']);

  const content = await workbook.xlsx.writeBuffer();
  const warnings = [];
  const result = await normalizeSpreadsheetAttachment({ filename: 'atividade.xlsx', content }, warnings);

  assert.equal(result.length, 1);
  assert.equal(result[0].nfse, '17915');
  assert.equal(result[0].intermediario, '3101');
});

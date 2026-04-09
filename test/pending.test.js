const test = require('node:test');
const assert = require('node:assert/strict');
const { buildPendingReport, pickRequestTable, requestedRowHasProof } = require('../src/normalize/pending');

test('identifica a tabela principal de pendencias pelo cabecalho', () => {
  const tables = [
    {
      id: 'table-1',
      title: 'Tabela 1',
      headers: ['Coluna 1', 'Coluna 2'],
      rows: [['A', 'B']],
      rowCount: 1,
      columnCount: 2,
    },
    {
      id: 'table-2',
      title: 'Tabela 2',
      headers: ['FATURA', 'VENC ORIGINAL', 'NUMERO RPS', 'CLIENTE', 'CNPJ ARCELOR'],
      rows: [['0001', '04/25/2026', '3947', 'ARCELORMITTAL PECEM SA', '09.509.535/0001-67']],
      rowCount: 1,
      columnCount: 5,
    },
  ];

  const picked = pickRequestTable(tables);

  assert.equal(picked.id, 'table-2');
});

test('considera comprovada a pendencia quando existe PDF oficial ou link oficial com o mesmo RPS', () => {
  const requestedRow = {
    nfse: '',
    rps: '003947',
    dps: '',
    tomador: 'ARCELORMITTAL PECEM SA',
    cnpj: '09.509.535/0001-67',
  };
  const proofRow = {
    nfse: '37426',
    rps: 'U1.00003947',
    dps: '00003947',
    tomador: 'ARCELORMITTAL PECEM SA',
    cnpj: '09.509.535/0001-67',
    notaUrl: 'https://nfe.prefeitura.sp.gov.br/contribuinte/notaprint.aspx?nf=37426',
  };

  assert.equal(requestedRowHasProof(requestedRow, proofRow), true);
});

test('retorna apenas as pendencias realmente nao comprovadas', () => {
  const requestedRows = [
    {
      nfse: '',
      rps: '003947',
      dps: '',
      tomador: 'ARCELORMITTAL PECEM SA',
      cnpj: '09.509.535/0001-67',
    },
    {
      nfse: '',
      rps: '079124',
      dps: '',
      tomador: 'ARCELORMITTAL BRASIL',
      cnpj: '17.469.701/0106-44',
    },
  ];

  const proofRows = [
    {
      nfse: '37426',
      rps: 'U1.00003947',
      dps: '00003947',
      tomador: 'ARCELORMITTAL PECEM SA',
      cnpj: '09.509.535/0001-67',
      arquivoPdf: '41349.pdf',
    },
  ];

  const report = buildPendingReport(requestedRows, proofRows);

  assert.equal(report.pendingRows.length, 1);
  assert.equal(report.pendingRows[0].rps, '079124');
  assert.equal(report.pendingTable.rowCount, 1);
});

const test = require('node:test');
const assert = require('node:assert/strict');
const { collapseDuplicateRows, mergeSpreadsheetRowsWithInvoices } = require('../src/normalize/merge');

test('nao encaixa PDF em linha errada apenas por valor parecido', () => {
  const warnings = [];
  const rows = [
    {
      nfse: '1470608',
      rps: 'U1.00472929',
      dps: '',
      emissao: new Date('2026-03-23T14:49:00'),
      dataFatoGerador: new Date('2026-03-18T00:00:00'),
      tomador: 'ARCELORMITTAL BRASIL S.A.',
      cnpj: '',
      intermediario: '',
      valorServico: 616.96,
      valorDeducao: 0,
      issDevido: 30.85,
      issPagar: 30.85,
      valorCredito: 0,
      issRetido: 'Nao',
      situacao: 'Normal',
      issPagoGuia: '',
      arquivoPdf: ''
    }
  ];

  const invoices = [
    {
      nfse: '3388',
      rps: '1224',
      dps: '1224',
      emissao: new Date('2026-03-24T09:25:20'),
      dataFatoGerador: new Date('2026-03-24T00:00:00'),
      tomador: 'ARCELORMITTAL BRASIL S.A.',
      cnpj: '17.469.701/0106-44',
      intermediario: '001',
      valorServico: 616.96,
      valorDeducao: 0,
      issDevido: 30.85,
      issPagar: 30.85,
      valorCredito: 0,
      issRetido: 'Nao Retido',
      situacao: 'Normal',
      issPagoGuia: '',
      arquivoPdf: '1224.pdf',
      sourceFile: '1224.pdf'
    }
  ];

  const result = mergeSpreadsheetRowsWithInvoices(rows, invoices, warnings, {
    appendUnmatchedInvoices: false,
    warnUnmatchedRows: false
  });

  assert.equal(result.length, 1);
  assert.equal(result[0].nfse, '1470608');
  assert.equal(result[0].arquivoPdf, '');
  assert.match(warnings[0], /1224\.pdf/);
});

test('mescla PDF quando o documento bate por RPS ou DPS', () => {
  const warnings = [];
  const rows = [
    {
      nfse: '',
      rps: '078502',
      dps: '',
      emissao: null,
      dataFatoGerador: null,
      tomador: 'ARCELORMITTAL BRASIL',
      cnpj: '',
      intermediario: '',
      valorServico: 2217.66,
      valorDeducao: '',
      issDevido: '',
      issPagar: '',
      valorCredito: '',
      issRetido: '',
      situacao: 'Normal',
      issPagoGuia: '',
      arquivoPdf: ''
    }
  ];

  const invoices = [
    {
      nfse: '2431',
      rps: '78502',
      dps: '78502',
      emissao: new Date('2026-03-11T16:55:34'),
      dataFatoGerador: new Date('2026-03-11T00:00:00'),
      tomador: 'ARCELORMITTAL BRASIL S A',
      cnpj: '17.469.701/0110-20',
      intermediario: '001',
      valorServico: 2217.66,
      valorDeducao: 0,
      issDevido: 110.88,
      issPagar: 110.88,
      valorCredito: 0,
      issRetido: 'Nao Retido',
      situacao: 'Normal',
      issPagoGuia: '',
      arquivoPdf: '78502.pdf',
      sourceFile: '78502.pdf'
    }
  ];

  const result = mergeSpreadsheetRowsWithInvoices(rows, invoices, warnings, {
    appendUnmatchedInvoices: false,
    warnUnmatchedRows: false
  });

  assert.equal(result.length, 1);
  assert.equal(result[0].nfse, '2431');
  assert.equal(result[0].rps, '78502');
  assert.equal(result[0].arquivoPdf, '78502.pdf');
  assert.equal(warnings.length, 0);
});

test('colapsa linhas duplicadas da mesma nota sem perder o CNPJ', () => {
  const rows = [
    {
      nfse: '17937',
      rps: 'U1.00017864',
      dps: '00017864',
      emissao: new Date('2026-03-12T10:39:00'),
      dataFatoGerador: new Date('2026-03-10T00:00:00'),
      tomador: 'ARCELORMITTAL BRASIL S.A.',
      cnpj: '',
      intermediario: '',
      valorServico: 612.66,
      valorDeducao: 0,
      issDevido: 30.63,
      issPagar: 30.63,
      valorCredito: 0,
      issRetido: 'Não',
      situacao: 'Normal',
      issPagoGuia: 'Sim',
      cartaDe: '',
      numeroObra: '',
      arquivoPdf: ''
    },
    {
      nfse: '17937',
      rps: 'U1.00017864',
      dps: '00017864',
      emissao: new Date('2026-03-12T10:39:00'),
      dataFatoGerador: new Date('2026-03-10T00:00:00'),
      tomador: '17.469.701/0083-13',
      cnpj: '17.469.701/0083-13',
      intermediario: '',
      valorServico: 612.66,
      valorDeducao: 0,
      issDevido: 30.63,
      issPagar: 30.63,
      valorCredito: 0,
      issRetido: 'Não',
      situacao: 'CANCELAR',
      issPagoGuia: 'Sim',
      cartaDe: '',
      numeroObra: '',
      arquivoPdf: ''
    }
  ];

  const result = collapseDuplicateRows(rows);

  assert.equal(result.length, 1);
  assert.equal(result[0].nfse, '17937');
  assert.equal(result[0].rps, 'U1.00017864');
  assert.equal(result[0].cnpj, '17.469.701/0083-13');
  assert.equal(result[0].tomador, 'ARCELORMITTAL BRASIL S.A.');
  assert.equal(result[0].situacao, 'Normal');
});

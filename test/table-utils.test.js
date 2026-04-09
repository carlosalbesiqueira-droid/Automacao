const test = require('node:test');
const assert = require('node:assert/strict');
const { buildCombinedTable } = require('../src/email/table-utils');

test('une tabelas diferentes em uma tabela geral usando o conjunto de cabecalhos', () => {
  const combined = buildCombinedTable([
    {
      id: 'table-1',
      title: 'Tabela 1',
      headers: ['FATURA', 'CLIENTE'],
      rows: [['0001', 'ARCELOR']],
      rowCount: 1,
      columnCount: 2
    },
    {
      id: 'table-2',
      title: 'Tabela 2',
      headers: ['CLIENTE', 'VALOR'],
      rows: [['ARMAR', '616,96']],
      rowCount: 1,
      columnCount: 2
    }
  ]);

  assert.equal(combined.title, 'Tabela Geral');
  assert.deepEqual(combined.headers, ['FATURA', 'CLIENTE', 'VALOR']);
  assert.deepEqual(combined.rows[0], ['0001', 'ARCELOR', '']);
  assert.deepEqual(combined.rows[1], ['', 'ARMAR', '616,96']);
});

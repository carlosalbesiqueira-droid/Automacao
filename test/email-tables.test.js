const test = require('node:test');
const assert = require('node:assert/strict');
const { extractHtmlTables } = require('../src/email/html-tables');

test('extrai tabelas HTML do corpo do e-mail', () => {
  const html = `
    <p>Conforme falamos ontem, segue em anexo as pendencias:</p>
    <table border="1">
      <tr>
        <th>FATURA</th>
        <th>VENC ORIGINAL</th>
        <th>NUMERO RPS</th>
        <th>CLIENTE</th>
      </tr>
      <tr>
        <td>00022118481-0003-DADOS</td>
        <td>25/04/2026</td>
        <td>3947</td>
        <td>ARCELORMITTAL PECEM SA</td>
      </tr>
      <tr>
        <td>00007624776-0045</td>
        <td>25/04/2026</td>
        <td>79124</td>
        <td>ARCELORMITTAL BRASIL</td>
      </tr>
    </table>
  `;

  const tables = extractHtmlTables(html);

  assert.equal(tables.length, 1);
  assert.equal(tables[0].headers[0], 'FATURA');
  assert.equal(tables[0].headers[1], 'VENC ORIGINAL');
  assert.equal(tables[0].rows.length, 2);
  assert.equal(tables[0].rows[0][2], '3947');
  assert.equal(tables[0].rows[1][3], 'ARCELORMITTAL BRASIL');
});

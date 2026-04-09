const test = require('node:test');
const assert = require('node:assert/strict');
const { buildFollowUpDraft, composeFollowUpPayload } = require('../src/email/follow-up');

test('gera rascunho de follow-up com destinatario, anexo e tabela html', () => {
  const draft = buildFollowUpDraft({
    pendingRows: [
      {
        nfse: '17915',
        rps: 'U1.00017838',
        emissao: new Date(2026, 2, 10, 8, 52),
        tomador: 'BELGO BEKAERT ARAMES LTDA',
        cnpj: '61.074.506/0026-98',
        valorServico: 424.72,
      },
    ],
    requestEmail: {
      subject: 'ENC: NFS - CONSUMO - EMISSAO MARCO 2026',
      to: '"Cliente" <cliente@exemplo.com>',
      from: '"Y3" <operacao@y3.com.br>',
    },
    responseEmail: {
      from: '"Cliente retorno" <retorno@exemplo.com>',
    },
    pendingFilePath: 'C:\\saida\\pendencias.xlsx',
  });

  assert.equal(draft.to, 'retorno@exemplo.com; cliente@exemplo.com');
  assert.equal(draft.cc, 'operacao@y3.com.br');
  assert.match(draft.subject, /Reforco de tratativa/);
  assert.equal(draft.attachmentNames[0], 'pendencias.xlsx');
  assert.match(draft.htmlBody, /BELGO BEKAERT ARAMES LTDA/);
  assert.match(draft.htmlBody, /424,72/);
});

test('recompõe payload editado mantendo anexos e pendencias', () => {
  const payload = composeFollowUpPayload({
    to: 'cliente@exemplo.com',
    cc: 'gestao@exemplo.com',
    subject: 'Cobranca',
    message: 'Prezados,\n\nFavor tratar as NFs abaixo.',
    pendingRows: [
      {
        nfse: '',
        rps: '003947',
        tomador: 'ARCELORMITTAL PECEM SA',
        cnpj: '09.509.535/0001-67',
        valorServico: 758.01,
      },
    ],
    attachmentPaths: ['C:\\saida\\pendencias.xlsx'],
  });

  assert.equal(payload.to, 'cliente@exemplo.com');
  assert.equal(payload.cc, 'gestao@exemplo.com');
  assert.equal(payload.pendingCount, 1);
  assert.equal(payload.attachmentNames[0], 'pendencias.xlsx');
  assert.match(payload.textBody, /Pendencias remanescentes/);
  assert.match(payload.htmlBody, /003947/);
});

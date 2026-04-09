const test = require('node:test');
const assert = require('node:assert/strict');
const { parseInvoicePdfText } = require('../src/parsers/pdf');

test('extrai dados principais da NFS-e a partir do texto do PDF', () => {
  const sampleText = `
PREFEITURA MUNICIPAL DO SALVADOR
Numero da Nota:
00017915
Data e Hora de Emissao:
10/03/2026 08:52:56
Codigo de Verificacao:
6K62-QFAU

PRESTADOR DE SERVICOS
Nome/Razao Social:
TELMEX DO BRASIL SA
CPF/CNPJ:
02.667.694/0046-42

TOMADOR DE SERVICOS
Nome/Razao Social:
BELGO BEKAERT ARAMES LTDA
CPF/CNPJ:
61.074.506/0026-98

VALOR TOTAL DA NOTA = R$424,72
Valor Deducoes (R$)
0,00
Valor do ISS (R$)
21,24
Credito Nota Salvador (R$)
0,00

OUTRAS INFORMACOES
- Esta Nota Salvador substitui o RPS no 17838 Serie U1, emitido em 09/03/2026.
- O ISS referente a esta Nota Salvador foi recolhido em 02/04/2026.
- Codigo de Tributacao do Municipio: 3101-001 - Servicos tecnicos em edificacoes
  `;

  const parsed = parseInvoicePdfText(sampleText, 'nfse.pdf');

  assert.equal(parsed.nfse, '00017915');
  assert.equal(parsed.rps, 'U1.00017838');
  assert.equal(parsed.dps, '17838');
  assert.equal(parsed.tomador, 'BELGO BEKAERT ARAMES LTDA');
  assert.equal(parsed.cnpj, '61.074.506/0026-98');
  assert.equal(parsed.valorServico, 424.72);
  assert.equal(parsed.issDevido, 21.24);
  assert.equal(parsed.issPagoGuia, 'Sim');
  assert.equal(parsed.intermediario, '3101');
  assert.equal(parsed.cartaDe, '');
  assert.equal(parsed.numeroObra, '');
});

test('preserva zeros a esquerda em numeros extraidos do PDF', () => {
  const sampleText = `
PREFEITURA MUNICIPAL DO SALVADOR
Numero da Nota:
00000045

OUTRAS INFORMACOES
- Esta Nota Salvador substitui o RPS no 00001234 Serie U1, emitido em 09/03/2026.
  `;

  const parsed = parseInvoicePdfText(sampleText, '00001234.pdf');

  assert.equal(parsed.nfse, '00000045');
  assert.equal(parsed.rps, 'U1.00001234');
  assert.equal(parsed.dps, '00001234');
});

test('extrai dados principais da DANFSe nacional do Rio de Janeiro', () => {
  const sampleText = `
DANFSe v1.0
Numero da NFS-e
2431
Competencia da NFS-e
11/03/2026
Data e Hora da emissao da NFS-e
11/03/2026 16:55:34
Numero da DPS
78502
Serie da DPS
1

EMITENTE DA NFS-e
Nome / Nome Empresarial
TELMEX DO BRASIL S/A
CNPJ / CPF / NIF
02.667.694/0002-21

TOMADOR DO SERVICO CNPJ / CPF / NIF
17.469.701/0110-20
Nome / Nome Empresarial
ARCELORMITTAL BRASIL S A

SERVICO PRESTADO
Codigo de Tributacao Nacional
31.01.04 - Servicos tecnicos em telecomunicacoes e congeneres.
Codigo de Tributacao Municipal
001 - Servicos tecnicos em telecomunicacoes.

Valor do Servico
R$ 2.217,66
Retencao do ISSQN
Nao Retido
ISSQN Apurado
R$ 110,88
  `;

  const parsed = parseInvoicePdfText(sampleText, '78502.pdf');

  assert.equal(parsed.nfse, '2431');
  assert.equal(parsed.rps, '78502');
  assert.equal(parsed.dps, '78502');
  assert.equal(parsed.tomador, 'ARCELORMITTAL BRASIL S A');
  assert.equal(parsed.cnpj, '17.469.701/0110-20');
  assert.equal(parsed.valorServico, 2217.66);
  assert.equal(parsed.issDevido, 110.88);
  assert.equal(parsed.issRetido, 'Nao Retido');
  assert.equal(parsed.intermediario, '3101');
});

test('extrai dados do layout de Sao Paulo sem misturar rotulos com tomador', () => {
  const sampleText = `
PREFEITURA DO MUNICIPIO DE SAO PAULO
SECRETARIA MUNICIPAL DA FAZENDA
NOTA FISCAL ELETRONICA DE SERVICOS - NFS-e
Numero da Nota
Data e Hora de Emissao
Codigo de Verificacao
20260408u02667694005029 RPS No 467601 Serie 00U1, emitido em 13/03/2026
01464835
18/03/2026 11:06:14
ZABS-VVIE
PRESTADOR DE SERVICOS
CPF/CNPJ: Inscricao Municipal:
Nome/Razao Social:
Endereco:
02.667.694/0001-40 2.714.344-9
TELMEX DO BRASIL S/A
R DOS INGLESES 600, 12 ANDAR - PARTE - MORRO DOS INGLESES - CEP: 01329-904
Municipio: Sao Paulo UF: SP
TOMADOR DE SERVICOS
Nome/Razao Social:
CPF/CNPJ: Inscricao Municipal:
Endereco:
Municipio: UF: E-mail:
ARCELORMITTAL BRASIL S.A.
17.469.701/0043-26 3.251.805-6
R ARLINDO BETTIO S/N, GALPAO 1 - JARDIM VERONICA - CEP: 03828-000
Sao Paulo SP nfe@arcelormittal.com.br
INTERMEDIARIO DE SERVICOS
CPF/CNPJ: Nome/Razao Social: ---- ----
DISCRIMINACAO DE SERVICOS
SERVICO VALOR ADICIONADO PABX VIRTUAL
VALOR TOTAL DO SERVICO = R$ 1.059,63
Valor Total das Deducoes (R$) Base de Calculo (R$) Aliquota (%) Valor do ISS (R$) Credito Programa da NFP (R$)
0,00 1.059,63 2,90% 30,72 0,00
  `;

  const parsed = parseInvoicePdfText(sampleText, 'nota_1464835.pdf');

  assert.equal(parsed.rps, '00U1.00467601');
  assert.equal(parsed.tomador, 'ARCELORMITTAL BRASIL S.A.');
  assert.equal(parsed.cnpj, '17.469.701/0043-26');
  assert.equal(parsed.intermediario, '');
  assert.equal(parsed.valorServico, 1059.63);
  assert.equal(parsed.issDevido, 30.72);
  assert.equal(parsed.numeroObra, '');
});

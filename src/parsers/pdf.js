const pdfParseModule = require('pdf-parse');
const {
  digitsOnly,
  formatCnpj,
  normalizeBooleanLike,
  normalizeDocumentNumber,
  normalizeTextPreservingCase,
  normalizeUppercaseText,
  parseBrazilianDate,
  parseBrazilianMoney
} = require('../utils/normalizers');

async function parsePdfAttachment(attachment, warnings) {
  const result = await extractPdfText(attachment.content);
  const parsed = parseInvoicePdfText(result.text, attachment.filename);

  if (!parsed.nfse && !parsed.dps) {
    warnings.push(`Nao foi possivel localizar o numero da nota ou do DPS no PDF "${attachment.filename}".`);
  }

  return {
    ...parsed,
    arquivoPdf: attachment.savedPath,
    arquivoPdfNome: attachment.filename
  };
}

async function extractPdfText(buffer) {
  if (typeof pdfParseModule === 'function') {
    return pdfParseModule(buffer);
  }

  if (typeof pdfParseModule.PDFParse === 'function') {
    const parser = new pdfParseModule.PDFParse({
      data: buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer)
    });

    try {
      return await parser.getText();
    } finally {
      if (typeof parser.destroy === 'function') {
        await parser.destroy();
      }
    }
  }

  throw new Error('Biblioteca pdf-parse indisponivel para leitura dos anexos em PDF.');
}

function parseInvoicePdfText(text, fileName = '') {
  const searchText = normalizeForSearch(text);
  const compactText = normalizeTextPreservingCase(searchText);
  const fileNumber = extractFileNumber(fileName);

  const prestadorSection =
    extractSection(searchText, 'PRESTADOR DE SERVICOS', 'TOMADOR DE SERVICOS') ||
    extractSection(searchText, 'EMITENTE DA NFS-E', 'TOMADOR DO SERVICO') ||
    extractSection(searchText, 'EMITENTE PRESTADOR DO SERVICO', 'TOMADOR DO SERVICO');
  const tomadorSection =
    extractSection(searchText, 'TOMADOR DE SERVICOS', 'DISCRIMINACAO DOS SERVICOS') ||
    extractSection(searchText, 'TOMADOR DO SERVICO', 'INTERMEDIARIO DO SERVICO') ||
    extractSection(searchText, 'TOMADOR DO SERVICO', 'SERVICO PRESTADO') ||
    extractSection(searchText, 'TOMADOR DO SERVICO', 'TRIBUTACAO MUNICIPAL');
  const infoSection =
    extractSection(searchText, 'OUTRAS INFORMACOES', '') ||
    extractSection(searchText, 'INFORMACOES COMPLEMENTARES', '');

  const nfse = normalizeDocumentNumber(
    capture(searchText, /NUMERO DA NOTA:\s*([0-9]+)/i) ||
      capture(searchText, /NUMERO DA NFS-E\s*([0-9]+)/i) ||
      capture(searchText, /COMPETENCIA\s+([0-9]+)\s*\/\s*[A-Z0-9]+\s+[0-9]{2}\/[0-9]{4}/i),
  );

  const emissao = parseBrazilianDate(
    capture(searchText, /DATA E HORA DE EMISSAO:\s*([0-9/: ]+)/i) ||
      capture(searchText, /DATA E HORA DA EMISSAO DA NFS-E\s*([0-9/: ]+)/i) ||
      capture(searchText, /([0-9]{2}\/[0-9]{2}\/[0-9]{4}\s+[0-9]{2}:[0-9]{2}:[0-9]{2})\s*DATA E HORA DE EMISSAO/i),
  );

  const rpsLegacyMatch =
    searchText.match(/RPS\s*N\S*\s*([0-9]+)\s*SERIE\s*([A-Z0-9]+)/i) ||
    searchText.match(/SUBSTITUI O RPS\s*N\S*\s*([0-9]+)\s*SERIE\s*([A-Z0-9]+)/i);
  const rpsCampinasMatch =
    searchText.match(/([0-9]+)\s*\/\s*([A-Z0-9]+)\s*NUMERO\s*\/\s*SERIE DO RPS/i);
  const dps = normalizeDocumentNumber(
    capture(searchText, /NUMERO DA DPS\s*([0-9]+)/i) ||
      rpsLegacyMatch?.[1] ||
      rpsCampinasMatch?.[1] ||
      fileNumber,
  );
  const rps = rpsLegacyMatch?.[1]
    ? buildRpsValue(rpsLegacyMatch[2], rpsLegacyMatch[1])
    : rpsCampinasMatch?.[1]
      ? buildRpsValue(rpsCampinasMatch[2], rpsCampinasMatch[1])
      : dps;

  const factDate = parseBrazilianDate(
    capture(searchText, /EMITIDO EM\s*([0-9]{2}\/[0-9]{2}\/[0-9]{4})/i) ||
      capture(searchText, /COMPETENCIA DA NFS-E\s*([0-9]{2}\/[0-9]{2}\/[0-9]{4})/i),
  );

  let tomador = extractNameFromSection(tomadorSection);
  let tomadorCnpj = formatCnpj(extractCnpj(tomadorSection));
  let prestador = extractNameFromSection(prestadorSection);
  let prestadorCnpj = formatCnpj(extractCnpj(prestadorSection));

  const campinasParties = extractCampinasParties(searchText);
  if ((!tomador || looksLikeServiceCodeLine(tomador)) && campinasParties.tomadorName) {
    tomador = campinasParties.tomadorName;
  }
  if (!tomadorCnpj && campinasParties.tomadorCnpj) {
    tomadorCnpj = campinasParties.tomadorCnpj;
  }
  if (!prestador && campinasParties.prestadorName) {
    prestador = campinasParties.prestadorName;
  }
  if (!prestadorCnpj && campinasParties.prestadorCnpj) {
    prestadorCnpj = campinasParties.prestadorCnpj;
  }

  const valorServico =
    parseBrazilianMoney(capture(searchText, /VALOR TOTAL DA NOTA\s*=\s*R\$\s*([0-9.,]+)/i)) ??
    parseBrazilianMoney(capture(searchText, /VALOR TOTAL DA NFSE CAMPINAS \(R\$\)\s*([0-9.,]+)/i)) ??
    parseBrazilianMoney(capture(searchText, /VALOR TOTAL DO SERVICO\s*=\s*R\$\s*([0-9.,]+)/i)) ??
    parseBrazilianMoney(capture(searchText, /VALOR TOTAL DA NFS-E[\s\S]*?VALOR DO SERVICO\s*R\$\s*([0-9.,]+)/i)) ??
    parseBrazilianMoney(capture(searchText, /VALOR DO SERVICO\s*R\$\s*([0-9.,]+)/i)) ??
    parseBrazilianMoney(capture(searchText, /VALOR LIQUIDO DA NFS(?:E|-E)[^\n]*\(R\$\)\s*([0-9.,]+)/i));

  const valorDeducao =
    parseBrazilianMoney(capture(searchText, /VALOR DEDUCOES \(R\$\)\s*([0-9.,]+)/i)) ??
    parseBrazilianMoney(capture(searchText, /TOTAL DEDUCOES\/REDUCOES\s*R\$\s*([0-9.,]+)/i)) ??
    0;
  const issDevido =
    parseBrazilianMoney(capture(searchText, /VALOR DO ISS \(R\$\)\s*([0-9.,]+)/i)) ??
    parseBrazilianMoney(capture(searchText, /ISSQN APURADO\s*R\$\s*([0-9.,]+)/i)) ??
    parseBrazilianMoney(captureSaoPauloIssValue(searchText)) ??
    parseBrazilianMoney(captureCampinasIssValue(searchText));
  const valorCredito =
    parseBrazilianMoney(capture(searchText, /CREDITO NOTA SALVADOR \(R\$\)\s*([0-9.,]+)/i)) ??
    0;
  const codigoServico =
    normalizeDocumentNumber(capture(infoSection, /CODIGO DE TRIBUTACAO DO MUNICIPIO:\s*([0-9-]+)/i)) ||
    normalizeDocumentNumber(capture(searchText, /CODIGO DE TRIBUTACAO MUNICIPAL\s*([0-9-]+)/i)) ||
    normalizeDocumentNumber(capture(searchText, /ITEM DA LISTA DE SERVICOS:\s*([0-9]+)/i)) ||
    normalizeDocumentNumber(capture(searchText, /\b([0-9]{2}\.[0-9]{2})\s*-\s*[A-Z]/i));
  const intermediario = extractIntermediario(searchText) || extractServiceActivityCode(searchText);
  const cartaDe = extractCartaDe(searchText);
  const numeroObra = extractNumeroObra(searchText);

  const issPagoGuia = /FOI RECOLHIDO EM/i.test(infoSection) ? 'Sim' : '';
  const situacao = /CANCELAD/i.test(compactText) ? 'Cancelada' : 'Normal';
  const issRetido = normalizeBooleanLike(
    capture(searchText, /ISS RETIDO\s*[:\-]?\s*(SIM|NAO)/i) ||
      capture(searchText, /RETENCAO DO ISSQN\s*(NAO RETIDO|RETIDO|NAO)/i) ||
      'Nao',
  );

  return {
    nfse,
    rps,
    dps,
    emissao,
    dataFatoGerador: factDate || emissao,
    tomador: normalizeUppercaseText(tomador),
    cnpj: tomadorCnpj,
    intermediario,
    valorServico,
    valorDeducao,
    issDevido,
    issPagar: issDevido,
    valorCredito,
    issRetido,
    situacao,
    issPagoGuia,
    cartaDe,
    numeroObra,
    prestador,
    prestadorCnpj,
    codigoServico,
    codigoVerificacao: normalizeTextPreservingCase(
      capture(searchText, /CODIGO DE VERIFICACAO:\s*([A-Z0-9-]+)/i),
    ),
    rawText: text,
    sourceFile: fileName
  };
}

function normalizeForSearch(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function buildRpsValue(series, number) {
  const normalizedSeries = normalizeTextPreservingCase(series);
  const digits = digitsOnly(number);
  if (!normalizedSeries) {
    return normalizeDocumentNumber(number);
  }

  return `${normalizedSeries}.${String(digits || '').padStart(8, '0')}`;
}

function extractFileNumber(fileName) {
  const match = String(fileName || '').match(/(\d+)/);
  return match?.[1] ? normalizeDocumentNumber(match[1]) : '';
}

function capture(text, pattern) {
  return text.match(pattern)?.[1]?.trim() ?? '';
}

function extractSection(text, startLabel, endLabel) {
  const startIndex = text.search(new RegExp(startLabel, 'i'));
  if (startIndex === -1) {
    return '';
  }

  const sliced = text.slice(startIndex);
  if (!endLabel) {
    return sliced;
  }

  const endIndex = sliced.search(new RegExp(endLabel, 'i'));
  return endIndex === -1 ? sliced : sliced.slice(0, endIndex);
}

function extractNameFromSection(sectionText) {
  const lines = sectionText
    .split(/\r?\n/)
    .map((line) => normalizeTextPreservingCase(line))
    .filter(Boolean);

  for (let index = 0; index < lines.length; index += 1) {
    if (/^nome\/razao social:?$/i.test(lines[index]) || /^nome\s*\/\s*nome empresarial:?$/i.test(lines[index])) {
      for (let nextIndex = index + 1; nextIndex <= Math.min(lines.length - 1, index + 6); nextIndex += 1) {
        const candidate = lines[nextIndex];
        if (looksLikeCompanyNameLine(candidate)) {
          return candidate;
        }
      }
    }
  }

  const fallbackLines = lines
    .filter((line) => !/prestador|tomador|cpf\/cnpj|cpf \/ cnpj|inscricao|endereco|e-mail|telefone|municipio|cep/i.test(line))
    .filter((line) => !/^[0-9./-]+$/.test(line))
    .filter((line) => !/^[0-9]+ \//.test(line));

  return fallbackLines[0] ?? '';
}

function extractCnpj(sectionText) {
  return (
    capture(sectionText, /CPF\/CNPJ:\s*([0-9./-]+)/i) ||
    capture(sectionText, /CNPJ\s*\/\s*CPF\s*\/\s*NIF\s*([0-9./-]+)/i) ||
    sectionText.match(/\b\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}\b/)?.[0] ||
    ''
  );
}

function extractCampinasParties(text) {
  const lines = String(text || '')
    .split(/\r?\n/)
    .map((line) => normalizeTextPreservingCase(line))
    .filter(Boolean);

  const cnpjs = [];
  lines.forEach((line, index) => {
    const match = line.match(/\b\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}\b/);
    if (match) {
      cnpjs.push({ index, value: formatCnpj(match[0]) });
    }
  });

  if (cnpjs.length < 2) {
    return {};
  }

  return {
    prestadorName: findNearbyCompanyName(lines, cnpjs[0].index),
    prestadorCnpj: cnpjs[0].value,
    tomadorName: findNearbyCompanyName(lines, cnpjs[1].index),
    tomadorCnpj: cnpjs[1].value
  };
}

function findNearbyCompanyName(lines, cnpjIndex) {
  for (let index = cnpjIndex - 1; index >= Math.max(0, cnpjIndex - 6); index -= 1) {
    const candidate = lines[index];
    if (looksLikeCompanyNameLine(candidate)) {
      return candidate;
    }
  }

  for (let index = cnpjIndex + 1; index <= Math.min(lines.length - 1, cnpjIndex + 6); index += 1) {
    const candidate = lines[index];
    if (looksLikeCompanyNameLine(candidate)) {
      return candidate;
    }
  }

  return '';
}

function looksLikeCompanyNameLine(value) {
  const text = normalizeTextPreservingCase(value);
  if (!text) {
    return false;
  }

  if (looksLikeServiceCodeLine(text)) {
    return false;
  }

  if (looksLikeLabelOnlyLine(text)) {
    return false;
  }

  if (/nome|cpf|cnpj|inscri[cç][aã]o|municipio|endere[cç]o|cep|telefone|e-mail|email|competencia|prefeitura|codigo|chave|tributa[cç][aã]o|reten[cç][aã]o|servi[cç]o|descricao/i.test(text)) {
    return false;
  }

  if (/^\(?\d{2}\)?/.test(text) || /\b\d{5}-\d{3}\b/.test(text) || /\/\s*[A-Z]{2}\b/.test(text)) {
    return false;
  }

  if (/\d/.test(text)) {
    return false;
  }

  if (!/[A-Z]/.test(text) || !/[A-Z]{2,}/.test(text)) {
    return false;
  }

  return true;
}

function looksLikeServiceCodeLine(value) {
  const text = normalizeTextPreservingCase(value);
  return /^\d{2,5}(?:[-./]\d{1,4})+/.test(text) || /^\d{2}\.\d{2}\s*-/.test(text);
}

function looksLikeLabelOnlyLine(value) {
  const text = normalizeTextPreservingCase(value);
  return /:$/.test(text) || /^(cpf\/cnpj|nome\/razao social|nome \/ nome empresarial|endereco|municipio|uf|e-mail|email|inscricao municipal)$/i.test(text);
}

function captureCampinasIssValue(text) {
  const match = String(text || '').match(
    /C[ÁA]LCULO DO ISSQN[\s\S]*?VALOR DO ISSQN \(R\$\)\s*([0-9.,]+)\s+([0-9.,]+)\s+([0-9.,]+)\s+([0-9.,]+)\s+([0-9.,]+)\s+([0-9.,]+)/i,
  );

  return match?.[2] || '';
}

function captureSaoPauloIssValue(text) {
  const match = String(text || '').match(
    /VALOR TOTAL DAS DEDUCOES \(R\$\)\s+BASE DE CALCULO \(R\$\)\s+ALIQUOTA \(%\)\s+VALOR DO ISS \(R\$\)\s+[^\n]*\n([0-9.,-]+)\s+([0-9.,-]+)\s+([0-9.,%]+)\s+([0-9.,-]+)/i,
  );

  return match?.[4] || '';
}

function extractIntermediario(text) {
  const section =
    extractSection(text, 'INTERMEDIARIO DO SERVICO', 'SERVICO PRESTADO') ||
    extractSection(text, 'INTERMEDIARIO DE SERVICOS', 'DISCRIMINACAO DOS SERVICOS') ||
    '';

  const normalizedSection = normalizeTextPreservingCase(section);
  if (!normalizedSection) {
    return '';
  }

  if (/nao identificado|não identificado/i.test(normalizedSection)) {
    return '';
  }

  const cnpj = extractCnpj(section);
  if (cnpj) {
    return formatCnpj(cnpj);
  }

  const name = extractNameFromSection(section);
  if (
    name &&
    !looksLikeServiceCodeLine(name) &&
    !/^intermediario de servicos?$/i.test(normalizeTextPreservingCase(name)) &&
    !/^----(?:\s+----)?$/.test(normalizeTextPreservingCase(name))
  ) {
    return normalizeUppercaseText(name);
  }

  return '';
}

function extractServiceActivityCode(text) {
  const candidates = [
    capture(text, /ITEM DA LISTA DE SERVICOS:\s*([0-9./-]+)/i),
    capture(text, /CODIGO DE TRIBUTACAO DO MUNICIPIO:\s*([0-9./-]+)/i),
    capture(text, /CODIGO DE TRIBUTACAO MUNICIPAL\s*([0-9./-]+)/i),
    capture(text, /CODIGO DE TRIBUTACAO NACIONAL\s*([0-9./-]+)/i),
    capture(text, /CNAE\/CBO\s*([0-9./-]+)/i),
  ];

  for (const candidate of candidates) {
    const normalized = normalizeServiceActivityValue(candidate);
    if (normalized) {
      return normalized;
    }
  }

  return '';
}

function normalizeServiceActivityValue(value) {
  const text = normalizeTextPreservingCase(value);
  if (!text) {
    return '';
  }

  const digits = String(digitsOnly(text)).replace(/^0+/, '');
  if (!digits) {
    return '';
  }

  if (digits.length >= 4) {
    return digits.slice(0, 4);
  }

  if (digits.length >= 3) {
    return digits.padStart(4, '0');
  }

  return '';
}

function extractCartaDe(text) {
  const inlineValue =
    capture(text, /CARTA DE(?: CORRECAO| CORREÇÃO)?\s*[:\-]?\s*([^\n]+)/i) ||
    capture(text, /CARTA DE\s*([^\n]+)/i);
  const normalized = normalizeTextPreservingCase(inlineValue);
  if (!normalized || /^-+$/.test(normalized) || /^carta de$/i.test(normalized)) {
    return '';
  }

  return normalized;
}

function extractNumeroObra(text) {
  const explicitValue =
    capture(text, /N(?:º|O|UMERO)?\s*DA OBRA\s*[:\-]\s*([A-Z0-9./-]+)/i) ||
    capture(text, /INSCRICAO DA OBRA\s*[:\-]\s*([A-Z0-9./-]+)/i) ||
    capture(text, /INSCRIÇÃO DA OBRA\s*[:\-]\s*([A-Z0-9./-]+)/i);
  const explicitNormalized = normalizeObraValue(explicitValue);
  if (explicitNormalized) {
    return explicitNormalized;
  }

  const headerBlock = String(text || '').match(/NUMERO INSCRICAO DA OBRA[^\n]*\n([^\n]+)/i);
  if (!headerBlock?.[1]) {
    return '';
  }

  const normalizedLine = normalizeTextPreservingCase(headerBlock[1]);
  if (!normalizedLine || /^-\s+-\s+-$/.test(normalizedLine)) {
    return '';
  }

  const parts = normalizedLine.split(/\s{2,}|\t/).map((part) => normalizeObraValue(part)).filter(Boolean);
  if (parts.length >= 2) {
    return parts[1];
  }

  return '';
}

function normalizeObraValue(value) {
  const normalized = normalizeTextPreservingCase(value);
  if (!normalized || /^-+$/.test(normalized)) {
    return '';
  }

  return normalized;
}

module.exports = {
  parseInvoicePdfText,
  parsePdfAttachment
};

const path = require('node:path');
const { formatDate, normalizeWhitespace } = require('../utils/normalizers');

const EMAIL_TABLE_COLUMNS = [
  { key: 'nfse', title: 'NFS-e' },
  { key: 'rpsDisplay', title: 'RPS' },
  { key: 'emissao', title: 'Emissao' },
  { key: 'tomador', title: 'Tomador de Servicos' },
  { key: 'cnpj', title: 'CNPJ' },
  { key: 'valorServico', title: 'Valor Servicos' },
];

function buildFollowUpDraft({ pendingRows, requestEmail, responseEmail, pendingFilePath }) {
  const normalizedRows = Array.isArray(pendingRows) ? pendingRows : [];
  const to = joinEmailAddresses(
    extractEmailAddresses(responseEmail?.from),
    extractEmailAddresses(requestEmail?.to),
  );
  const cc = joinEmailAddresses(
    extractEmailAddresses(requestEmail?.from).filter((address) => !includesEmail(to, address)),
  );
  const referenceSubject = normalizeWhitespace(
    requestEmail?.subject || responseEmail?.subject || 'pendencias de NFS-e',
  );
  const subject = `Reforco de tratativa das NFs pendentes | ${referenceSubject}`;
  const message = buildDefaultMessage(normalizedRows.length);
  const attachmentPaths = pendingFilePath ? [pendingFilePath] : [];

  return composeFollowUpPayload({
    to,
    cc,
    subject,
    message,
    pendingRows: normalizedRows,
    attachmentPaths,
  });
}

function composeFollowUpPayload({ to, cc, subject, message, pendingRows, attachmentPaths = [] }) {
  const normalizedRows = Array.isArray(pendingRows) ? pendingRows : [];
  const normalizedMessage = normalizeWhitespaceBlocks(message);
  const normalizedSubject = normalizeWhitespace(subject);
  const normalizedTo = normalizeEmailList(to);
  const normalizedCc = normalizeEmailList(cc);
  const normalizedAttachments = attachmentPaths
    .map((item) => String(item || '').trim())
    .filter(Boolean);

  return {
    to: normalizedTo,
    cc: normalizedCc,
    subject: normalizedSubject,
    message: normalizedMessage,
    pendingCount: normalizedRows.length,
    attachmentPaths: normalizedAttachments,
    attachmentNames: normalizedAttachments.map((item) => path.basename(item)),
    htmlBody: buildFollowUpHtml(normalizedMessage, normalizedRows),
    textBody: buildFollowUpText(normalizedMessage, normalizedRows),
  };
}

function buildDefaultMessage(pendingCount) {
  const countLabel = pendingCount === 1
    ? '1 NF que continua sem comprovacao oficial'
    : `${pendingCount} NFs que continuam sem comprovacao oficial`;

  return [
    'Prezados,',
    '',
    `Apos a conferencia entre o relatorio enviado e o retorno recebido, identificamos ${countLabel}.`,
    '',
    'Solicitamos, por gentileza, a tratativa dessas pendencias e o reenvio da documentacao correspondente.',
    '',
    'Segue anexa a planilha com o detalhamento das pendencias remanescentes.',
    '',
    'Ficamos no aguardo.',
    '',
    'Atenciosamente,',
    'Y3 Gestao Telecom',
  ].join('\n');
}

function buildFollowUpHtml(message, pendingRows) {
  const paragraphs = String(message || '')
    .split(/\n\s*\n/)
    .map((item) => normalizeWhitespace(item))
    .filter(Boolean);

  const introHtml = paragraphs
    .map((paragraph) => `<p style="margin:0 0 14px 0;line-height:1.6;color:#10263d;">${escapeHtml(paragraph).replace(/\n/g, '<br>')}</p>`)
    .join('');

  const tableHtml = buildPendingTableHtml(pendingRows);
  const countLabel = pendingRows.length === 1 ? '1 pendencia remanescente' : `${pendingRows.length} pendencias remanescentes`;

  return [
    '<div style="font-family:Segoe UI,Arial,sans-serif;color:#10263d;background:#ffffff;">',
    '  <div style="max-width:980px;margin:0 auto;padding:24px;">',
    introHtml,
    '    <div style="margin:20px 0 12px 0;padding:14px 16px;border:1px solid #d6e6f4;border-radius:14px;background:#f5fbff;">',
    `      <strong style="display:block;color:#0b4f75;font-size:14px;letter-spacing:0.06em;text-transform:uppercase;">Pendencias remanescentes</strong>`,
    `      <span style="display:block;margin-top:8px;font-size:16px;color:#10263d;">${escapeHtml(countLabel)}</span>`,
    '    </div>',
    tableHtml,
    '    <p style="margin:18px 0 0 0;line-height:1.6;color:#425d76;">A planilha com o detalhamento das pendencias remanescentes segue anexa a este e-mail.</p>',
    '  </div>',
    '</div>',
  ].join('');
}

function buildPendingTableHtml(rows) {
  if (!rows.length) {
    return [
      '<div style="padding:16px;border:1px dashed #d6e6f4;border-radius:14px;background:#fbfdff;color:#425d76;">',
      'Nenhuma pendencia remanescente foi identificada nesta execucao.',
      '</div>',
    ].join('');
  }

  const headerHtml = EMAIL_TABLE_COLUMNS
    .map((column) => `<th style="padding:10px 12px;border:1px solid #d6e6f4;background:#eef7ff;text-align:left;font-size:13px;">${escapeHtml(column.title)}</th>`)
    .join('');

  const rowHtml = rows.map((row) => {
    const cellHtml = EMAIL_TABLE_COLUMNS
      .map((column) => `<td style="padding:10px 12px;border:1px solid #d6e6f4;vertical-align:top;font-size:13px;">${escapeHtml(formatEmailCell(column.key, row))}</td>`)
      .join('');
    return `<tr>${cellHtml}</tr>`;
  }).join('');

  return [
    '<table style="width:100%;border-collapse:collapse;border-spacing:0;background:#ffffff;">',
    `  <thead><tr>${headerHtml}</tr></thead>`,
    `  <tbody>${rowHtml}</tbody>`,
    '</table>',
  ].join('');
}

function buildFollowUpText(message, pendingRows) {
  const baseMessage = normalizeWhitespaceBlocks(message);
  const lines = [baseMessage, '', 'Pendencias remanescentes:'];

  if (!pendingRows.length) {
    lines.push('- Nenhuma pendencia remanescente identificada.');
    return lines.join('\n');
  }

  pendingRows.forEach((row) => {
    lines.push([
      `- NFS-e: ${formatEmailCell('nfse', row) || '-'}`,
      `RPS: ${formatEmailCell('rpsDisplay', row) || '-'}`,
      `Tomador: ${formatEmailCell('tomador', row) || '-'}`,
      `CNPJ: ${formatEmailCell('cnpj', row) || '-'}`,
      `Valor: ${formatEmailCell('valorServico', row) || '-'}`,
    ].join(' | '));
  });

  return lines.join('\n');
}

function formatEmailCell(key, row) {
  if (!row || typeof row !== 'object') {
    return '';
  }

  if (key === 'rpsDisplay') {
    return buildRpsDisplay(row);
  }

  const value = row[key];
  if (value == null || value === '') {
    return '';
  }

  if (value instanceof Date) {
    return formatDate(value, key === 'emissao');
  }

  if (typeof value === 'number') {
    return value.toLocaleString('pt-BR', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  }

  return String(value);
}

function buildRpsDisplay(row) {
  return String(row?.rps || row?.dps || '').trim();
}

function extractEmailAddresses(value) {
  const matches = String(value || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/ig) || [];
  return dedupe(matches.map((item) => item.trim()));
}

function includesEmail(emailList, address) {
  return normalizeEmailList(emailList)
    .split(';')
    .map((item) => item.trim().toLowerCase())
    .filter(Boolean)
    .includes(String(address || '').trim().toLowerCase());
}

function joinEmailAddresses(...groups) {
  return normalizeEmailList(groups.flat());
}

function normalizeEmailList(value) {
  const list = Array.isArray(value)
    ? value
    : String(value || '')
      .split(/[;,]/)
      .map((item) => item.trim())
      .filter(Boolean);

  return dedupe(list).join('; ');
}

function dedupe(values) {
  const seen = new Set();
  const result = [];

  values.forEach((item) => {
    const normalized = String(item || '').trim();
    const key = normalized.toLowerCase();
    if (!normalized || seen.has(key)) {
      return;
    }

    seen.add(key);
    result.push(normalized);
  });

  return result;
}

function normalizeWhitespaceBlocks(value) {
  return String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .map((line) => line.replace(/\s+/g, ' ').trim())
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

module.exports = {
  buildFollowUpDraft,
  composeFollowUpPayload,
};

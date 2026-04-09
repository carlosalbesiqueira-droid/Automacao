const fs = require('node:fs/promises');
const path = require('node:path');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const { extractHtmlTables } = require('./html-tables');
const { ensureDirectory, safeFileName } = require('../utils/fs');
const { formatDate, normalizeTextPreservingCase } = require('../utils/normalizers');

async function loadEmailFromEml(emlPath) {
  const absolutePath = path.resolve(emlPath);
  const raw = await fs.readFile(absolutePath);
  return loadEmailFromBuffer(raw, {
    sourcePath: absolutePath,
    mailbox: 'arquivo .eml'
  });
}

async function loadEmailFromBuffer(buffer, metadata = {}) {
  const parsed = await simpleParser(buffer);

  return buildEmailPayload(parsed, {
    source: 'eml',
    mailbox: metadata.mailbox || 'arquivo .eml',
    sourcePath: metadata.sourcePath ?? '',
    uid: metadata.uid ?? null,
    provider: metadata.provider ?? ''
  });
}

async function fetchEmailBySubject({
  subject,
  mailbox = 'INBOX',
  unseenOnly = false,
  connection = {}
}) {
  const resolvedConnection = resolveMailConnection(connection);

  const client = new ImapFlow({
    host: resolvedConnection.host,
    port: resolvedConnection.port,
    secure: resolvedConnection.secure,
    connectionTimeout: 15000,
    greetingTimeout: 15000,
    socketTimeout: 30000,
    auth: {
      user: resolvedConnection.user,
      pass: resolvedConnection.pass
    }
  });

  let lock = null;
  try {
    await client.connect();
    lock = await client.getMailboxLock(mailbox);

    const query = { subject };
    if (unseenOnly) {
      query.seen = false;
    }

    const uids = await client.search(query, { uid: true });
    if (!uids.length) {
      throw new Error(`Nenhum e-mail foi encontrado com o assunto contendo "${subject}" na pasta "${mailbox}".`);
    }

    const messages = [];
    for (const uid of uids) {
      const message = await client.fetchOne(uid, { envelope: true, internalDate: true }, { uid: true });
      if (message) {
        messages.push(message);
      }
    }

    messages.sort((left, right) => {
      const leftTime = left.internalDate ? new Date(left.internalDate).getTime() : 0;
      const rightTime = right.internalDate ? new Date(right.internalDate).getTime() : 0;
      return rightTime - leftTime;
    });

    const selected = messages[0];
    const rawMessage = await client.fetchOne(selected.uid, { source: true }, { uid: true });
    const parsed = await simpleParser(rawMessage.source);

    return buildEmailPayload(parsed, {
      source: 'imap',
      mailbox,
      uid: selected.uid,
      provider: resolvedConnection.provider
    });
  } catch (error) {
    throw normalizeImapError(error, resolvedConnection);
  } finally {
    if (lock) {
      lock.release();
    }

    try {
      if (!client.usable || client.disabled) {
        client.close();
      } else {
        await client.logout();
      }
    } catch {
      client.close();
    }
  }
}

function buildEmailPayload(parsed, metadata) {
  return {
    source: metadata.source,
    sourcePath: metadata.sourcePath ?? '',
    mailbox: metadata.mailbox,
    uid: metadata.uid ?? null,
    provider: metadata.provider ?? '',
    subject: normalizeTextPreservingCase(parsed.subject),
    from: parsed.from?.text ?? '',
    to: parsed.to?.text ?? '',
    date: parsed.date ?? null,
    text: parsed.text ?? '',
    html: parsed.html ?? '',
    tables: extractHtmlTables(parsed.html ?? ''),
    attachments: (parsed.attachments ?? []).map((attachment, index) => ({
      id: index + 1,
      filename: attachment.filename || `anexo_${index + 1}`,
      contentType: attachment.contentType || '',
      contentDisposition: attachment.contentDisposition || '',
      size: attachment.size || attachment.content?.length || 0,
      kind: detectAttachmentKind(attachment.filename, attachment.contentType),
      content: attachment.content
    }))
  };
}

async function saveAttachments(attachments, targetDirectory) {
  await ensureDirectory(targetDirectory);

  const saved = [];
  for (const attachment of attachments) {
    const safeName = safeFileName(attachment.filename || `anexo_${attachment.id}`);
    const fullPath = path.join(targetDirectory, safeName);
    await fs.writeFile(fullPath, attachment.content);
    saved.push({
      ...attachment,
      savedPath: fullPath
    });
  }

  return saved;
}

function summarizeEmail(email) {
  return [
    `Assunto: ${email.subject || '(sem assunto)'}`,
    `Remetente: ${email.from || '(sem remetente)'}`,
    `Destinatario: ${email.to || '(sem destinatario)'}`,
    `Data: ${email.date ? formatDate(new Date(email.date), true) : '(sem data)'}`,
    `Origem: ${email.source === 'imap' ? `IMAP/${email.mailbox}` : email.sourcePath}`,
    '',
    'Corpo do e-mail:',
    email.text?.trim() || '(sem corpo em texto)'
  ].join('\n');
}

function resolveMailConnection(connection = {}) {
  const provider = String(connection.provider || process.env.MAIL_PROVIDER || 'custom').toLowerCase();
  const presets = {
    google: {
      host: 'imap.gmail.com',
      port: 993,
      secure: true,
      label: 'Google'
    },
    outlook: {
      host: 'outlook.office365.com',
      port: 993,
      secure: true,
      label: 'Outlook'
    }
  };

  const preset = presets[provider] || null;
  const host = connection.host || preset?.host || process.env.MAIL_HOST || '';
  const port = Number(connection.port || preset?.port || process.env.MAIL_PORT || 993);
  const secureValue = connection.secure ?? preset?.secure ?? process.env.MAIL_SECURE ?? true;
  const secure = typeof secureValue === 'boolean'
    ? secureValue
    : String(secureValue).toLowerCase() !== 'false';
  const user = connection.user || process.env.MAIL_USER || '';
  const pass = connection.pass || process.env.MAIL_PASS || '';

  const missing = [];
  if (!host) {
    missing.push('host IMAP');
  }
  if (!user) {
    missing.push('e-mail de acesso');
  }
  if (!pass) {
    missing.push(provider === 'google' ? 'senha de app/senha' : 'senha');
  }

  if (missing.length) {
    const providerLabel = preset?.label || 'provedor personalizado';
    throw new Error(
      `Configuracao incompleta para ${providerLabel}. Informe ${missing.join(', ')} para visualizar ou processar o e-mail.`,
    );
  }

  return {
    provider,
    host,
    port,
    secure,
    user,
    pass
  };
}

function normalizeImapError(error, connection) {
  if (!error) {
    return new Error('Falha desconhecida ao conectar no provedor de e-mail.');
  }

  const providerLabel = connection.provider === 'google'
    ? 'Google'
    : connection.provider === 'outlook'
      ? 'Outlook'
      : 'provedor de e-mail';

  const responseText = normalizeTextPreservingCase(error.responseText || '');

  if (error.authenticationFailed || /invalid credentials/i.test(responseText)) {
    if (connection.provider === 'google') {
      return new Error(
        'Falha na autenticacao do Google. O Gmail recusou este acesso com senha comum. Para esta automacao, use uma senha de app com a verificacao em duas etapas ativada.',
      );
    }

    if (connection.provider === 'outlook') {
      return new Error(
        'Falha na autenticacao do Outlook. Revise o e-mail, a senha informada e se a caixa permite acesso IMAP.',
      );
    }

    return new Error(
      `Falha na autenticacao do ${providerLabel}. Revise o e-mail, a senha e as configuracoes IMAP informadas.`,
    );
  }

  if (error.code === 'ETIMEDOUT' || /timed out/i.test(error.message || '')) {
    return new Error(
      `Tempo esgotado ao conectar no ${providerLabel}. Verifique a internet, o host IMAP e tente novamente.`,
    );
  }

  if (responseText) {
    return new Error(`Falha no ${providerLabel}: ${responseText}`);
  }

  return new Error(error.message || `Falha ao consultar o ${providerLabel}.`);
}

function detectAttachmentKind(filename, contentType) {
  const name = String(filename || '').toLowerCase();
  const type = String(contentType || '').toLowerCase();

  if (/\.(xlsx|xls|xlsm|csv)$/i.test(name) || /spreadsheet|excel|csv/.test(type)) {
    return 'spreadsheet';
  }

  if (/\.pdf$/i.test(name) || type.includes('pdf')) {
    return 'pdf';
  }

  return 'other';
}

module.exports = {
  fetchEmailBySubject,
  loadEmailFromBuffer,
  loadEmailFromEml,
  saveAttachments,
  summarizeEmail
};

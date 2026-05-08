const { spawn } = require('node:child_process');

async function sendEmailWithOutlook({
  from,
  to,
  cc,
  subject,
  htmlBody,
  textBody = '',
  attachmentPaths = [],
  inlineAttachments = [],
  sendNow = false,
}) {
  if (!to) {
    throw new Error('Informe ao menos um destinatario para abrir ou enviar o e-mail pelo Outlook.');
  }

  if (!subject) {
    throw new Error('Informe o assunto do e-mail antes de continuar.');
  }

  if (!htmlBody) {
    throw new Error('Nao foi possivel montar o corpo do e-mail para o Outlook.');
  }

  const payload = JSON.stringify({
    from: from || '',
    to,
    cc,
    subject,
    htmlBody,
    attachments: attachmentPaths,
    inlineAttachments,
    sendNow: Boolean(sendNow),
  });

  try {
    const output = await runPowerShell(buildOutlookScript(), payload);
    try {
      return JSON.parse(output);
    } catch {
      return {
        ok: true,
        action: sendNow ? 'send' : 'draft',
        raw: output,
      };
    }
  } catch (error) {
    if (!isDesktopOutlookUnavailable(error)) {
      throw error;
    }

    return openOutlookFallback({
      to,
      cc,
      subject,
      body: textBody || stripHtml(htmlBody),
      sendNow,
      hasAttachments: attachmentPaths.length > 0 || inlineAttachments.length > 0,
    });
  }
}

function runPowerShell(script, stdinPayload) {
  return new Promise((resolve, reject) => {
    const child = spawn(
      'powershell.exe',
      ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-Command', script],
      {
        windowsHide: true,
        stdio: ['pipe', 'pipe', 'pipe'],
      },
    );

    let stdout = '';
    let stderr = '';

    child.stdout.on('data', (chunk) => {
      stdout += String(chunk);
    });

    child.stderr.on('data', (chunk) => {
      stderr += String(chunk);
    });

    child.on('error', (error) => {
      reject(new Error(`Nao foi possivel iniciar o PowerShell para acionar o Outlook: ${error.message}`));
    });

    child.on('close', (code) => {
      if (code !== 0) {
        reject(new Error(
          normalizePowerShellError(stderr)
          || 'O Outlook nao respondeu ao comando de envio. Verifique se ele esta instalado e configurado nesta maquina.',
        ));
        return;
      }

      resolve(stdout.trim());
    });

    child.stdin.write(stdinPayload);
    child.stdin.end();
  });
}

function buildOutlookScript() {
  return `
    $ErrorActionPreference = 'Stop'
    $payloadJson = [Console]::In.ReadToEnd()
    if ([string]::IsNullOrWhiteSpace($payloadJson)) {
      throw 'Payload vazio para o Outlook.'
    }

    $payload = $payloadJson | ConvertFrom-Json
    $outlook = $null
    $mail = $null

    try {
      $outlook = New-Object -ComObject Outlook.Application
      $mail = $outlook.CreateItem(0)

      if ($payload.from) {
        $matchedAccount = $null
        foreach ($account in $outlook.Session.Accounts) {
          if ($account.SmtpAddress -and $account.SmtpAddress.ToLower() -eq ([string]$payload.from).ToLower()) {
            $matchedAccount = $account
            break
          }
        }

        if (-not $matchedAccount) {
          throw "Nao encontrei a conta Outlook do remetente informado: $($payload.from)"
        }

        $mail.SendUsingAccount = $matchedAccount
      }

      $mail.To = [string]$payload.to

      if ($payload.cc) {
        $mail.CC = [string]$payload.cc
      }

      $mail.Subject = [string]$payload.subject
      $mail.HTMLBody = [string]$payload.htmlBody

      foreach ($attachmentPath in @($payload.attachments)) {
        if ($attachmentPath -and (Test-Path -LiteralPath $attachmentPath)) {
          [void]$mail.Attachments.Add($attachmentPath)
        }
      }

      foreach ($inlineAttachment in @($payload.inlineAttachments)) {
        if ($inlineAttachment.path -and (Test-Path -LiteralPath $inlineAttachment.path)) {
          $mailAttachment = $mail.Attachments.Add($inlineAttachment.path)

          if ($inlineAttachment.cid) {
            $propertyAccessor = $mailAttachment.PropertyAccessor
            $propertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F', [string]$inlineAttachment.cid)
            $propertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x7FFE000B', $true)
          }
        }
      }

      if ($payload.sendNow -eq $true) {
        $mail.Send()
        $action = 'send'
      }
      else {
        $mail.Display()
        $action = 'draft'
      }

      [pscustomobject]@{
        ok = $true
        action = $action
        to = [string]$payload.to
        subject = [string]$payload.subject
      } | ConvertTo-Json -Compress
    }
    catch {
      Write-Error $_.Exception.Message
      exit 1
    }
  `;
}

async function openOutlookFallback({ to, cc, subject, body, sendNow, hasAttachments }) {
  const composeUri = buildMsOutlookComposeUri({ to, cc, subject, body });
  const webUrl = buildOutlookWebComposeUrl({ to, cc, subject, body });
  const fallbackPayload = JSON.stringify({ composeUri, webUrl });

  try {
    const output = await runPowerShell(buildOutlookFallbackScript(), fallbackPayload);
    const parsed = JSON.parse(output);

    return {
      ok: true,
      action: 'draft',
      fallback: parsed.channel || 'new-outlook',
      message: sendNow
        ? buildFallbackMessage('O envio automatico nao esta disponivel neste tipo de Outlook. Abrimos um rascunho preenchido para voce concluir o envio.', hasAttachments)
        : buildFallbackMessage('Abrimos um rascunho preenchido no Outlook disponivel nesta maquina.', hasAttachments),
    };
  } catch {
    throw new Error('Nao foi possivel abrir automaticamente o Outlook desta maquina. Use os botoes de copiar assunto e mensagem e envie pelo seu cliente de e-mail.');
  }
}

function buildOutlookFallbackScript() {
  return `
    $ErrorActionPreference = 'Stop'
    $payloadJson = [Console]::In.ReadToEnd()
    if ([string]::IsNullOrWhiteSpace($payloadJson)) {
      throw 'Payload vazio para o fallback do Outlook.'
    }

    $payload = $payloadJson | ConvertFrom-Json

    try {
      Start-Process $payload.composeUri
      [pscustomobject]@{
        ok = $true
        channel = 'new-outlook'
      } | ConvertTo-Json -Compress
    }
    catch {
      Start-Process $payload.webUrl
      [pscustomobject]@{
        ok = $true
        channel = 'outlook-web'
      } | ConvertTo-Json -Compress
    }
  `;
}

function buildMsOutlookComposeUri({ to, cc, subject, body }) {
  return `ms-outlook://compose?${buildComposeQueryString({ to, cc, subject, body })}`;
}

function buildOutlookWebComposeUrl({ to, cc, subject, body }) {
  return `https://outlook.office.com/mail/deeplink/compose?${buildComposeQueryString({ to, cc, subject, body })}`;
}

function buildComposeQueryString(fields) {
  return Object.entries(fields || {})
    .map(([key, value]) => [key, String(value || '').trim()])
    .filter(([, value]) => value)
    .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
    .join('&');
}

function buildFallbackMessage(baseMessage, hasAttachments) {
  if (!hasAttachments) {
    return baseMessage;
  }

  return `${baseMessage} Como este modo de abertura nao permite anexar arquivos automaticamente, confira e anexe a planilha manualmente antes de enviar.`;
}

function isDesktopOutlookUnavailable(error) {
  return /n[aã]o foi poss[ií]vel abrir o outlook nesta m[aá]quina|nao foi possivel abrir o outlook nesta maquina|class not registered|80040154|activex/i.test(String(error?.message || ''));
}

function stripHtml(value) {
  return String(value || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

function normalizePowerShellError(stderr) {
  const message = String(stderr || '')
    .replace(/\r/g, '\n')
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
    .filter((line) => !/^at line:/i.test(line))
    .join(' ');

  if (!message) {
    return '';
  }

  if (/active x component can.t create object|cannot create activex component|class not registered|classe n[aã]o registrada|80040154/i.test(message)) {
    return 'Nao foi possivel abrir o Outlook nesta maquina. Se o Outlook nao estiver disponivel, copie o assunto e a mensagem e envie pelo seu cliente de e-mail.';
  }

  return message;
}

module.exports = {
  buildOutlookWebComposeUrl,
  sendEmailWithOutlook,
};

const { spawn } = require('node:child_process');

async function sendEmailWithOutlook({ to, cc, subject, htmlBody, attachmentPaths = [], sendNow = false }) {
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
    to,
    cc,
    subject,
    htmlBody,
    attachments: attachmentPaths,
    sendNow: Boolean(sendNow),
  });

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
        reject(new Error(normalizePowerShellError(stderr) || 'O Outlook nao respondeu ao comando de envio. Verifique se ele esta instalado e configurado nesta maquina.'));
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
    return 'Nao foi possivel abrir o Outlook nesta maquina. Verifique se o Outlook desktop esta instalado e configurado.';
  }

  return message;
}

module.exports = {
  sendEmailWithOutlook,
};

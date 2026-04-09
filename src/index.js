#!/usr/bin/env node
const dotenv = require('dotenv');
const { Command } = require('commander');
const { processEmailAutomation } = require('./process-email');

dotenv.config();

const program = new Command();

program
  .name('processar-email')
  .description('Busca um e-mail com anexos, normaliza a planilha e cruza os dados com as NFS-e em PDF.')
  .option('--subject <texto>', 'Assunto ou trecho do assunto do e-mail')
  .option('--eml <caminho>', 'Arquivo .eml salvo localmente para processamento manual')
  .option('--mailbox <nome>', 'Caixa IMAP a ser consultada', process.env.MAILBOX || 'INBOX')
  .option('--output <pasta>', 'Pasta de saida', 'output')
  .option('--unseen-only', 'Busca apenas e-mails nao lidos', false)
  .parse(process.argv);

const options = program.opts();

async function main() {
  validateOptions(options);
  const result = await processEmailAutomation({
    subject: options.subject,
    emlPath: options.eml,
    mailbox: options.mailbox,
    output: options.output,
    unseenOnly: options.unseenOnly
  });

  console.log(`Assunto processado: ${result.email.subject}`);
  console.log(`Planilha original: ${result.spreadsheetAttachment.filename}`);
  console.log(`PDFs encontrados: ${result.pdfAttachments.length}`);
  console.log(`Linhas finais: ${result.finalRows.length}`);
  console.log(`Saida principal: ${result.workbookPath}`);
  console.log(`Resumo do e-mail: ${result.emailSummaryPath}`);
}

function validateOptions(currentOptions) {
  if (!currentOptions.subject && !currentOptions.eml) {
    throw new Error('Informe --subject para busca IMAP ou --eml para processar um arquivo salvo localmente.');
  }
}

main().catch((error) => {
  console.error(`Erro: ${error.message}`);
  process.exitCode = 1;
});

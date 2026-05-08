const TIFLUX_API_BASE = 'https://SEU-ENDERECO-PUBLICO/bot-faturas-api';
const TIFLUX_SHEET_NAME = 'TIFLUX_PREENCHIMENTO';
const TIFLUX_BUTTON_CELL = 'S2';
const TIFLUX_STATUS_CELL = 'T2';
const TIFLUX_JOB_CELL = 'U2';

const TIFLUX_HEADERS = [
  'NUMERO_TICKET',
  'Histórico da Fatura',
  'Impedimento',
  'Tratativas/observações',
  'Fatura Assumida (Data)',
  'BO+DT (Data)',
  'RPS+NF (Data)',
  'NF Prefeitura',
  'AE (Data)',
  'Importação (Data)',
  'Envio (Data)',
  'Concluído (Data)',
  'Estágio',
  'STATUS_EXECUCAO',
  'MENSAGEM_EXECUCAO',
  'PROCESSADO_EM',
  'CAMPOS_APLICADOS',
  'EVIDENCIA',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TiFlux Y3')
    .addItem('Preparar aba do TiFlux', 'criarAbaModeloTiflux')
    .addItem('Executar agora', 'processarAbaTiflux')
    .addItem('Instalar botao por checkbox', 'instalarAcionadorBotaoTiflux')
    .addItem('Consultar ultimo job', 'consultarUltimoJobTiflux')
    .addToUi();
}

function criarAbaModeloTiflux() {
  const spreadsheet = SpreadsheetApp.getActive();
  const existing = spreadsheet.getSheetByName(TIFLUX_SHEET_NAME);
  const sheet = existing || spreadsheet.insertSheet(TIFLUX_SHEET_NAME);

  sheet.clear();
  sheet.getRange(1, 1, 1, TIFLUX_HEADERS.length).setValues([TIFLUX_HEADERS]);
  sheet.getRange('S1:U3').setValues([
    ['ACAO', 'STATUS_BOTAO', 'ULTIMO_JOB'],
    [false, 'Aguardando', ''],
    ['Marque a caixa em S2', '', ''],
  ]);
  sheet.getRange('S2').insertCheckboxes();
  sheet.setFrozenRows(1);
  sheet.getRange('A1:Q1').setFontWeight('bold').setBackground('#d9ead3');
  sheet.getRange('S1:U3').setFontWeight('bold').setBackground('#fce5cd');
  sheet.autoResizeColumns(1, 21);

  SpreadsheetApp.getUi().alert(
    'Aba preparada. Preencha os tickets e os campos desejados. Depois marque a caixa S2 para executar.'
  );
}

function instalarAcionadorBotaoTiflux() {
  const spreadsheet = SpreadsheetApp.getActive();
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'aoEditarBotaoTiflux') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('aoEditarBotaoTiflux')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    'Acionador instalado. Agora, sempre que a caixa S2 for marcada na aba TIFLUX_PREENCHIMENTO, a automacao sera disparada.'
  );
}

function aoEditarBotaoTiflux(e) {
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== TIFLUX_SHEET_NAME) {
    return;
  }
  if (e.range.getA1Notation() !== TIFLUX_BUTTON_CELL) {
    return;
  }
  if (String(e.value).toUpperCase() !== 'TRUE') {
    return;
  }

  executarAutomacaoTiflux_({ showAlert: false, resetCheckbox: true });
}

function processarAbaTiflux() {
  executarAutomacaoTiflux_({ showAlert: true, resetCheckbox: false });
}

function executarAutomacaoTiflux_(options) {
  const settings = Object.assign({ showAlert: true, resetCheckbox: false }, options || {});
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName(TIFLUX_SHEET_NAME);

  if (!sheet) {
    if (settings.showAlert) {
      SpreadsheetApp.getUi().alert('A aba TIFLUX_PREENCHIMENTO nao foi encontrada.');
    }
    return;
  }

  sheet.getRange(TIFLUX_STATUS_CELL).setValue('Processando...');

  const payload = {
    spreadsheet_id: spreadsheet.getId(),
    worksheet_name: sheet.getName(),
  };

  const response = UrlFetchApp.fetch(`${TIFLUX_API_BASE}/v1/tiflux/google-sheet/run`, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const body = JSON.parse(response.getContentText() || '{}');
  if (response.getResponseCode() >= 300 || !body.ok) {
    sheet.getRange(TIFLUX_STATUS_CELL).setValue(`Erro: ${body.detail || body.message || body.error || 'Falha no disparo'}`);
    if (settings.resetCheckbox) {
      sheet.getRange(TIFLUX_BUTTON_CELL).setValue(false);
    }
    if (settings.showAlert) {
      SpreadsheetApp.getUi().alert(`Falha ao iniciar o processamento: ${body.detail || body.message || body.error || response.getContentText()}`);
    }
    return;
  }

  PropertiesService.getDocumentProperties().setProperty('TIFLUX_LAST_JOB_ID', body.job_id);
  sheet.getRange(TIFLUX_STATUS_CELL).setValue('Processamento iniciado');
  sheet.getRange(TIFLUX_JOB_CELL).setValue(body.job_id);

  if (settings.resetCheckbox) {
    sheet.getRange(TIFLUX_BUTTON_CELL).setValue(false);
  }

  if (settings.showAlert) {
    SpreadsheetApp.getUi().alert(
      `Processamento iniciado com sucesso.\nJob: ${body.job_id}\nA aba sera atualizada nas colunas de retorno conforme os tickets forem processados.`
    );
  }
}

function consultarUltimoJobTiflux() {
  const jobId = PropertiesService.getDocumentProperties().getProperty('TIFLUX_LAST_JOB_ID');
  if (!jobId) {
    SpreadsheetApp.getUi().alert('Nenhum job do TiFlux foi executado ainda nesta planilha.');
    return;
  }

  const response = UrlFetchApp.fetch(`${TIFLUX_API_BASE}/v1/tiflux/google-sheet/jobs/${jobId}`, {
    method: 'get',
    muteHttpExceptions: true,
  });

  const body = JSON.parse(response.getContentText() || '{}');
  if (response.getResponseCode() >= 300) {
    SpreadsheetApp.getUi().alert(`Nao foi possivel consultar o job ${jobId}: ${body.detail || response.getContentText()}`);
    return;
  }

  const summary = body.processed
    ? `\nLinhas processadas: ${body.processed}\nOK: ${body.updated || 0}\nFalhas: ${body.failed || 0}`
    : '';

  SpreadsheetApp.getUi().alert(
    `Job: ${jobId}\nStatus: ${body.status || 'desconhecido'}\nMensagem: ${body.message || '-'}${summary}`
  );
}

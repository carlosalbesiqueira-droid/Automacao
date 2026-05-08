let apiBase = '';
let currentRunId = '';
let currentFiles = null;
const embedMode = new URLSearchParams(window.location.search).get('embed') === '1';
let lastEmbeddedHeight = 0;
let embeddedHeightTimer = null;

const form = document.getElementById('compareForm');
const healthBadge = document.getElementById('healthBadge');
const requestInput = document.getElementById('requestEmlFile');
const responseInput = document.getElementById('responseEmlFile');
const requestUploadText = document.getElementById('requestUploadText');
const responseUploadText = document.getElementById('responseUploadText');
const previewButton = document.getElementById('comparePreviewButton');
const submitButton = document.getElementById('compareSubmitButton');

const requestPreviewState = document.getElementById('requestPreviewState');
const requestPreviewContent = document.getElementById('requestPreviewContent');
const requestPreviewSubject = document.getElementById('requestPreviewSubject');
const requestPreviewMeta = document.getElementById('requestPreviewMeta');
const requestPreviewAttachments = document.getElementById('requestPreviewAttachments');
const requestPreviewAttachmentHint = document.getElementById('requestPreviewAttachmentHint');
const requestPreviewTables = document.getElementById('requestPreviewTables');
const requestPreviewTableCount = document.getElementById('requestPreviewTableCount');

const responsePreviewState = document.getElementById('responsePreviewState');
const responsePreviewContent = document.getElementById('responsePreviewContent');
const responsePreviewSubject = document.getElementById('responsePreviewSubject');
const responsePreviewMeta = document.getElementById('responsePreviewMeta');
const responsePreviewAttachments = document.getElementById('responsePreviewAttachments');
const responsePreviewAttachmentHint = document.getElementById('responsePreviewAttachmentHint');
const responsePreviewTables = document.getElementById('responsePreviewTables');
const responsePreviewTableCount = document.getElementById('responsePreviewTableCount');

const resultPlaceholder = document.getElementById('resultPlaceholder');
const resultContent = document.getElementById('resultContent');
const resultTitle = document.getElementById('resultTitle');
const resultBadge = document.getElementById('resultBadge');
const resultInsightTitle = document.getElementById('resultInsightTitle');
const resultInsightText = document.getElementById('resultInsightText');
const resultMetrics = document.getElementById('resultMetrics');
const resultFiles = document.getElementById('resultFiles');
const resultTabDifference = document.getElementById('resultTabDifference');
const resultTabPending = document.getElementById('resultTabPending');
const resultTabGeneral = document.getElementById('resultTabGeneral');
const resultDifferenceTabCount = document.getElementById('resultDifferenceTabCount');
const resultPendingTabCount = document.getElementById('resultPendingTabCount');
const resultCombinedTabCount = document.getElementById('resultCombinedTabCount');
const resultDifferenceCount = document.getElementById('resultDifferenceCount');
const resultDifferenceTable = document.getElementById('resultDifferenceTable');
const resultPendingCount = document.getElementById('resultPendingCount');
const resultCombinedCount = document.getElementById('resultCombinedCount');
const resultPendingTable = document.getElementById('resultPendingTable');
const resultCombinedTable = document.getElementById('resultCombinedTable');
const warningsCount = document.getElementById('warningsCount');
const warningsList = document.getElementById('warningsList');
const followUpPanel = document.getElementById('followUpPanel');
const followUpHint = document.getElementById('followUpHint');
const followUpFrom = document.getElementById('followUpFrom');
const followUpTo = document.getElementById('followUpTo');
const followUpCc = document.getElementById('followUpCc');
const followUpSubject = document.getElementById('followUpSubject');
const followUpMessage = document.getElementById('followUpMessage');
const followUpSignatureImage = document.getElementById('followUpSignatureImage');
const followUpSignatureHint = document.getElementById('followUpSignatureHint');
const followUpMeta = document.getElementById('followUpMeta');
let followUpWebButton = document.getElementById('followUpWebButton');
const followUpDraftButton = document.getElementById('followUpDraftButton');
const followUpSendButton = document.getElementById('followUpSendButton');
const copyFollowUpSubjectButton = document.getElementById('copyFollowUpSubjectButton');
const copyFollowUpMessageButton = document.getElementById('copyFollowUpMessageButton');
const followUpFeedback = document.getElementById('followUpFeedback');

boot();

async function boot() {
  ensureFollowUpWebButton();
  normalizeFollowUpLabels();
  applyFollowUpTextOverrides();
  attachPreviewHandlers();
  attachFormHandler();
  attachResultHandlers();
  attachEmbedReporter();
  setPreviewIdle('request', 'Envie o e-mail ou a planilha do pedido para visualizar o que foi solicitado.');
  setPreviewIdle('response', 'Envie o e-mail ou a planilha do retorno para visualizar o que o cliente mandou.');
  setFollowUpPlaceholder(
    'O e-mail de cobrança será montado aqui ao final do comparativo.',
    'Quando houver pendências remanescentes ou diferenças relevantes, o sistema preenche remetente, destinatários, assunto e mensagem para você editar antes de abrir ou enviar.'
  );
  await checkHealth();
}

function ensureFollowUpWebButton() {
  if (followUpWebButton) {
    return;
  }

  const actionRow = followUpDraftButton?.closest('.followup-actions');
  if (!actionRow) {
    return;
  }

  followUpWebButton = document.createElement('button');
  followUpWebButton.type = 'button';
  followUpWebButton.id = 'followUpWebButton';
  followUpWebButton.className = 'submit-button';
  followUpWebButton.innerHTML = [
    '<span>Abrir no Outlook Web</span>',
    '<small>Abre o rascunho no navegador e baixa a planilha para anexar</small>',
  ].join('');

  actionRow.insertBefore(followUpWebButton, actionRow.firstChild);
}

function applyFollowUpTextOverrides() {
  const title = followUpPanel?.querySelector('.preview-pane-header strong');
  if (title) {
    title.textContent = 'E-mail de cobranca pronto para editar e enviar';
  }

  if (followUpHint) {
    followUpHint.textContent = 'Depois do comparativo, os campos abaixo serao preenchidos automaticamente';
  }

  setFieldCopy(followUpFrom, 'Remetente do e-mail', 'Use o e-mail de referencia. No Outlook Web, o envio sai pela conta que estiver logada no navegador.');
  setFieldCopy(followUpTo, 'Para', 'Voce pode ajustar os destinatarios antes de abrir ou enviar.');
  setFieldCopy(followUpCc, 'CC', 'Use ponto e virgula para mais de um destinatario.');
  setFieldCopy(followUpMessage, 'Mensagem editavel', 'Ao abrir no Outlook Web, a mensagem sera preenchida neste aparelho. A planilha de pendencias sera baixada para voce anexar antes do envio.');
  setFieldCopy(followUpSignatureImage, 'Assinatura em imagem (opcional)', 'Se necessario, selecione uma imagem para baixar junto e inserir manualmente no Outlook Web.');

  if (followUpDraftButton) {
    followUpDraftButton.hidden = true;
    followUpDraftButton.setAttribute('aria-hidden', 'true');
  }

  if (followUpSendButton) {
    followUpSendButton.hidden = true;
    followUpSendButton.setAttribute('aria-hidden', 'true');
  }

  updateButtonCopy(followUpWebButton, 'Abrir no Outlook Web', 'Abre o rascunho no navegador e baixa a planilha para anexar');
  updateButtonCopy(copyFollowUpSubjectButton, 'Copiar assunto', 'Util quando o Outlook nao estiver disponivel');
  updateButtonCopy(copyFollowUpMessageButton, 'Copiar mensagem', 'Copie o texto e envie pelo cliente de e-mail de sua preferencia');
}

function updateButtonCopy(button, label, helper) {
  if (!button) {
    return;
  }

  const span = button.querySelector('span');
  const small = button.querySelector('small');

  if (span && label) {
    span.textContent = label;
  }

  if (small && helper) {
    small.textContent = helper;
  }
}

function normalizeFollowUpLabels() {
  setFieldCopy(followUpFrom, 'Remetente do e-mail', 'Use o e-mail que deve aparecer como remetente. Se o Outlook estiver disponível, a automação tenta usar essa conta.');
  setFieldCopy(followUpTo, 'Para', 'Você pode ajustar os destinatários antes de abrir ou enviar.');
  setFieldCopy(followUpCc, 'CC', 'Use ponto e vírgula para mais de um destinatário.');
  setFieldCopy(followUpSubject, 'Assunto', '');
  setFieldCopy(followUpMessage, 'Mensagem editável', 'A tabela com as pendências e a planilha anexa serão adicionadas automaticamente. Se o Outlook não estiver disponível, copie a mensagem e envie pelo seu cliente de e-mail.');
  setFieldCopy(followUpSignatureImage, 'Assinatura em imagem (opcional)', 'Se necessário, anexe uma imagem de assinatura para entrar no e-mail.');

  if (followUpHint) {
    followUpHint.textContent = 'Depois do comparativo, os campos abaixo serão preenchidos automaticamente';
  }

  const followUpTitle = followUpPanel?.querySelector('.preview-pane-header strong');
  if (followUpTitle) {
    followUpTitle.textContent = 'E-mail de cobrança pronto para editar e enviar';
  }

  if (followUpDraftButton) {
    const small = followUpDraftButton.querySelector('small');
    if (small) {
      small.textContent = 'Gera o e-mail para revisar antes do envio, quando o Outlook estiver disponível';
    }
  }

  if (followUpSendButton) {
    const small = followUpSendButton.querySelector('small');
    if (small) {
      small.textContent = 'Tenta enviar automaticamente usando o Outlook, quando disponível';
    }
  }

  if (copyFollowUpSubjectButton) {
    const small = copyFollowUpSubjectButton.querySelector('small');
    if (small) {
      small.textContent = 'Útil quando o Outlook não estiver disponível';
    }
  }

  if (copyFollowUpMessageButton) {
    const small = copyFollowUpMessageButton.querySelector('small');
    if (small) {
      small.textContent = 'Copie o texto e envie pelo cliente de e-mail de sua preferência';
    }
  }
}

function setFieldCopy(field, labelText, helperText) {
  if (!field) {
    return;
  }

  const wrapper = field.closest('.field');
  if (!wrapper) {
    return;
  }

  const label = wrapper.querySelector('span');
  if (label && labelText) {
    label.textContent = labelText;
  }

  const helper = wrapper.querySelector('small');
  if (helper && helperText) {
    helper.textContent = helperText;
  }
}

function attachEmbedReporter() {
  if (!embedMode || window.parent === window) {
    return;
  }

  const postHeight = () => {
    const height = Math.ceil(Math.max(
      document.body?.scrollHeight || 0,
      document.documentElement?.scrollHeight || 0,
      document.body?.offsetHeight || 0,
      document.documentElement?.offsetHeight || 0,
      document.body?.getBoundingClientRect()?.height || 0,
      document.documentElement?.getBoundingClientRect()?.height || 0,
    ));

    if (!Number.isFinite(height) || height <= 0) {
      return;
    }

    if (Math.abs(height - lastEmbeddedHeight) < 4) {
      return;
    }

    lastEmbeddedHeight = height;

    window.parent.postMessage({
      type: 'y3-compare-height',
      height,
    }, '*');
  };

  const schedulePostHeight = () => {
    window.clearTimeout(embeddedHeightTimer);
    embeddedHeightTimer = window.setTimeout(postHeight, 60);
  };

  window.addEventListener('load', postHeight);
  window.addEventListener('resize', schedulePostHeight);

  if (typeof ResizeObserver === 'function') {
    const observer = new ResizeObserver(() => schedulePostHeight());
    if (document.body) {
      observer.observe(document.body);
    }
    if (document.documentElement) {
      observer.observe(document.documentElement);
    }
  }

  window.setTimeout(postHeight, 60);
  window.setTimeout(postHeight, 600);
}

function attachPreviewHandlers() {
  requestInput.addEventListener('change', async () => {
    requestUploadText.innerHTML = requestInput.files?.[0]
      ? `Arquivo selecionado: <strong title="${escapeHtml(requestInput.files[0].name)}">${escapeHtml(compactFileName(requestInput.files[0].name))}</strong>`
      : 'Clique para escolher o <code>.eml</code> ou a planilha do que foi solicitado.';

    if (requestInput.files?.[0]) {
      await requestSidePreview('request');
    } else {
      setPreviewIdle('request', 'Envie o e-mail ou a planilha do pedido para visualizar o que foi solicitado.');
    }
  });

  responseInput.addEventListener('change', async () => {
    responseUploadText.innerHTML = responseInput.files?.[0]
      ? `Arquivo selecionado: <strong title="${escapeHtml(responseInput.files[0].name)}">${escapeHtml(compactFileName(responseInput.files[0].name))}</strong>`
      : 'Clique para escolher o <code>.eml</code> ou a planilha do que foi entregue pelo cliente.';

    if (responseInput.files?.[0]) {
      await requestSidePreview('response');
    } else {
      setPreviewIdle('response', 'Envie o e-mail ou a planilha do retorno para visualizar o que o cliente mandou.');
    }
  });

  previewButton.addEventListener('click', async () => {
    await Promise.all([
      requestSidePreview('request'),
      requestSidePreview('response'),
    ]);
  });

  document.addEventListener('change', (event) => {
    if (event.target.classList.contains('attachment-checkbox')) {
      updateAttachmentSelectionLabel(event.target.dataset.side);
    }

    if (event.target.classList.contains('table-checkbox')) {
      updateTableSelectionLabel(event.target.dataset.side);
    }
  });
}

function attachFormHandler() {
  form.addEventListener('submit', async (event) => {
    event.preventDefault();

    if (!requestInput.files?.[0]) {
      renderError('Envie um arquivo .eml ou uma planilha no portal "o que pedimos".');
      return;
    }

    if (!responseInput.files?.[0]) {
      renderError('Envie um arquivo .eml ou uma planilha no portal "o que o cliente mandou".');
      return;
    }

    if (!await ensureApiReady()) {
      renderError('A API da automação não respondeu. Rode `npm.cmd run web` para processar o comparativo.');
      return;
    }

    submitButton.disabled = true;
    submitButton.querySelector('span').textContent = 'Processando comparativo...';
    submitButton.querySelector('small').textContent = 'Comparando o pedido com o retorno do cliente';
    setResultLoading();

    try {
      const response = await fetch(apiUrl('/api/process-compare'), {
        method: 'POST',
        body: buildCompareFormData()
      });

      const data = await response.json();
      if (!response.ok || !data.ok) {
        throw new Error(data.error || 'Falha ao processar o comparativo.');
      }

      renderSuccess(data);
    } catch (error) {
      renderError(error.message);
    } finally {
      submitButton.disabled = false;
      submitButton.querySelector('span').textContent = 'Processar comparativo';
      submitButton.querySelector('small').textContent = 'Retornar apenas as pendencias que continuam sem comprovacao oficial';
    }
  });
}

function attachResultHandlers() {
  resultFiles.addEventListener('click', async (event) => {
    const trigger = event.target.closest('[data-download-url]');
    if (!trigger) {
      return;
    }

    event.preventDefault();
    await downloadToBrowser(trigger.dataset.downloadUrl, trigger.dataset.downloadName, trigger);
  });

  resultContent.addEventListener('click', (event) => {
    const trigger = event.target.closest('[data-result-view]');
    if (!trigger) {
      return;
    }

    setResultView(trigger.dataset.resultView);
  });

  if (followUpWebButton) {
    followUpWebButton.addEventListener('click', async () => {
      await openFollowUpInWeb();
    });
  }

  followUpDraftButton.addEventListener('click', async () => {
    await openFollowUpInMailApp();
  });

  followUpSendButton.addEventListener('click', async () => {
    if (!window.confirm('Isso vai enviar o follow-up agora pelo Outlook. Deseja continuar?')) {
      return;
    }

    await sendFollowUp('send');
  });

  if (followUpSignatureImage) {
    followUpSignatureImage.addEventListener('change', () => {
      updateSignatureHint();
    });
  }

  if (copyFollowUpSubjectButton) {
    copyFollowUpSubjectButton.addEventListener('click', async () => {
      const subject = followUpSubject.value.trim();
      if (!subject) {
        showFollowUpFeedback('error', 'Preencha o assunto antes de copiar.');
        return;
      }

      await copyToClipboard(subject, 'Assunto copiado para a área de transferência.');
    });
  }

  if (copyFollowUpMessageButton) {
    copyFollowUpMessageButton.addEventListener('click', async () => {
      const message = followUpMessage.value.trim();
      if (!message) {
        showFollowUpFeedback('error', 'Preencha a mensagem antes de copiar.');
        return;
      }

      await copyToClipboard(message, 'Mensagem copiada para a área de transferência.');
    });
  }
}

async function checkHealth() {
  const candidates = buildApiCandidates();

  for (const candidate of candidates) {
    try {
      const response = await fetch(`${candidate}/api/health`);
      const data = await response.json();

      if (!response.ok || !data.ok) {
        continue;
      }

      apiBase = candidate;
      healthBadge.textContent = 'Sistema online';
      healthBadge.classList.remove('is-error');
      healthBadge.classList.add('is-ok');
      return true;
    } catch {
      // tenta a proxima origem
    }
  }

  apiBase = '';
  healthBadge.textContent = 'API offline';
  healthBadge.classList.remove('is-ok');
  healthBadge.classList.add('is-error');
  return false;
}

async function ensureApiReady() {
  return apiBase ? true : checkHealth();
}

async function requestSidePreview(side) {
  const input = side === 'request' ? requestInput : responseInput;
  if (!input.files?.[0]) {
    setPreviewIdle(side, side === 'request'
      ? 'Envie o e-mail ou a planilha do pedido para visualizar o que foi solicitado.'
      : 'Envie o e-mail ou a planilha do retorno para visualizar o que o cliente mandou.');
    return;
  }

  if (!await ensureApiReady()) {
    renderSidePreviewError(side, 'A API da automacao nao respondeu. Rode `npm.cmd run web` e tente novamente.');
    return;
  }

  setPreviewLoading(side);

  try {
    const payload = new FormData();
    payload.set('side', side);
    payload.set(side === 'request' ? 'requestEmlFile' : 'responseEmlFile', input.files[0]);

    const response = await fetch(apiUrl('/api/preview-compare-side'), {
      method: 'POST',
      body: payload
    });

    const data = await response.json();
    if (!response.ok || !data.ok) {
      throw new Error(data.error || 'Não foi possível carregar a prévia deste portal.');
    }

    renderSidePreview(side, data.preview);
  } catch (error) {
    renderSidePreviewError(side, error.message);
  }
}

function buildCompareFormData() {
  const payload = new FormData();
  payload.set('output', form.elements.output.value || 'output');

  if (requestInput.files?.[0]) {
    payload.set('requestEmlFile', requestInput.files[0]);
  }

  if (responseInput.files?.[0]) {
    payload.set('responseEmlFile', responseInput.files[0]);
  }

  const requestAttachmentIds = getSelectedIds('request', 'attachment');
  if (document.querySelector('.attachment-checkbox[data-side="request"]')) {
    payload.set('requestAttachmentSelectionApplied', 'true');
  }
  requestAttachmentIds.forEach((id) => payload.append('requestSelectedAttachmentIds', id));

  const responseAttachmentIds = getSelectedIds('response', 'attachment');
  if (document.querySelector('.attachment-checkbox[data-side="response"]')) {
    payload.set('responseAttachmentSelectionApplied', 'true');
  }
  responseAttachmentIds.forEach((id) => payload.append('responseSelectedAttachmentIds', id));

  const requestTableIds = getSelectedIds('request', 'table');
  if (document.querySelector('.table-checkbox[data-side="request"]')) {
    payload.set('requestTableSelectionApplied', 'true');
  }
  requestTableIds.forEach((id) => payload.append('requestSelectedTableIds', id));

  const responseTableIds = getSelectedIds('response', 'table');
  if (document.querySelector('.table-checkbox[data-side="response"]')) {
    payload.set('responseTableSelectionApplied', 'true');
  }
  responseTableIds.forEach((id) => payload.append('responseSelectedTableIds', id));

  return payload;
}

function getSelectedIds(side, kind) {
  return Array.from(document.querySelectorAll(`.${kind}-checkbox[data-side="${side}"]:checked`))
    .map((input) => input.value)
    .filter(Boolean);
}

function setPreviewIdle(side, message) {
  const state = side === 'request' ? requestPreviewState : responsePreviewState;
  const content = side === 'request' ? requestPreviewContent : responsePreviewContent;
  state.classList.remove('is-hidden');
  content.classList.add('is-hidden');
  state.innerHTML = `
    <strong>Aguardando arquivo</strong>
    <p>${escapeHtml(message)}</p>
  `;
}

function setPreviewLoading(side) {
  const state = side === 'request' ? requestPreviewState : responsePreviewState;
  const content = side === 'request' ? requestPreviewContent : responsePreviewContent;
  state.classList.remove('is-hidden');
  content.classList.add('is-hidden');
  state.innerHTML = `
    <strong>Lendo este portal</strong>
    <p>Estamos extraindo tabelas e anexos do arquivo enviado.</p>
  `;
}

function renderSidePreview(side, preview) {
  const state = side === 'request' ? requestPreviewState : responsePreviewState;
  const content = side === 'request' ? requestPreviewContent : responsePreviewContent;
  const subject = side === 'request' ? requestPreviewSubject : responsePreviewSubject;
  const meta = side === 'request' ? requestPreviewMeta : responsePreviewMeta;
  const attachments = side === 'request' ? requestPreviewAttachments : responsePreviewAttachments;
  const tables = side === 'request' ? requestPreviewTables : responsePreviewTables;
  const tableCount = side === 'request' ? requestPreviewTableCount : responsePreviewTableCount;

  state.classList.add('is-hidden');
  content.classList.remove('is-hidden');
  subject.textContent = preview.subject || '(sem assunto)';
  meta.innerHTML = [
    metaCard('Remetente', preview.from || '-'),
    metaCard('Destinatario', preview.to || '-'),
    metaCard('Data', formatDateTime(preview.date)),
    metaCard('Origem', preview.mailbox || 'upload .eml')
  ].join('');

  attachments.innerHTML = preview.attachments.length
    ? preview.attachments.map((attachment) => attachmentOption(attachment, side)).join('')
    : '<div class="empty-inline">Nenhum anexo foi identificado neste portal.</div>';

  tables.innerHTML = preview.tables.length
    ? preview.tables.map((table, index) => tablePreview(table, index, side)).join('')
    : '<div class="empty-inline">Nenhuma tabela estruturada foi identificada neste portal.</div>';

  tableCount.textContent = `${preview.tableCount || 0} tabela(s)`;
  updateAttachmentSelectionLabel(side);
  updateTableSelectionLabel(side);
}

function renderSidePreviewError(side, message) {
  const state = side === 'request' ? requestPreviewState : responsePreviewState;
  const content = side === 'request' ? requestPreviewContent : responsePreviewContent;
  state.classList.remove('is-hidden');
  content.classList.add('is-hidden');
  state.innerHTML = `
    <strong>Não foi possível carregar a prévia</strong>
    <p>${escapeHtml(message)}</p>
  `;
}

function updateAttachmentSelectionLabel(side) {
  const hint = side === 'request' ? requestPreviewAttachmentHint : responsePreviewAttachmentHint;
  const total = document.querySelectorAll(`.attachment-checkbox[data-side="${side}"]`).length;
  const selected = document.querySelectorAll(`.attachment-checkbox[data-side="${side}"]:checked`).length;
  hint.textContent = total
    ? `${selected} de ${total} anexo(s) selecionado(s)`
    : 'Nenhum anexo encontrado';
}

function updateTableSelectionLabel(side) {
  const target = side === 'request' ? requestPreviewTableCount : responsePreviewTableCount;
  const total = document.querySelectorAll(`.table-checkbox[data-side="${side}"]`).length;
  const selected = document.querySelectorAll(`.table-checkbox[data-side="${side}"]:checked`).length;
  target.textContent = total
    ? `${selected} de ${total} tabela(s) selecionada(s)`
    : '0 tabela(s)';
}

function setResultLoading() {
  currentRunId = '';
  currentFiles = null;
  resultPlaceholder.classList.add('is-hidden');
  resultContent.classList.remove('is-hidden');
  resultTitle.textContent = 'Processando comparativo';
  resultBadge.textContent = 'Em execução';
  resultBadge.classList.remove('is-error');
  resultInsightTitle.textContent = 'Comparando pedido e retorno';
  resultInsightText.textContent = 'Estamos conferindo o que foi pedido contra a planilha do cliente, os links oficiais e os PDFs do governo.';
  resultMetrics.innerHTML = '';
  resultFiles.innerHTML = '<div class="help-callout"><strong>Execucao em andamento</strong><p>A automacao esta comparando o pedido original com o retorno do cliente.</p></div>';
  resultDifferenceTabCount.textContent = '...';
  resultPendingTabCount.textContent = '...';
  resultCombinedTabCount.textContent = '...';
  resultDifferenceCount.textContent = '...';
  resultPendingCount.textContent = '...';
  resultCombinedCount.textContent = '...';
  resultDifferenceTable.innerHTML = '<div class="empty-inline">As diferenças entre as bases serão exibidas aqui ao final da execução.</div>';
  resultPendingTable.innerHTML = '<div class="empty-inline">As pendências remanescentes serão exibidas aqui ao final da execução.</div>';
  resultCombinedTable.innerHTML = '<div class="empty-inline">A base padronizada do retorno será exibida aqui ao final da execução.</div>';
  warningsList.innerHTML = '';
  warningsCount.textContent = '...';
  setFollowUpPlaceholder(
    'Estamos preparando o e-mail de cobrança.',
    'Assim que o comparativo terminar, os campos serão preenchidos automaticamente com base nas pendências remanescentes ou nas diferenças encontradas.'
  );
  setResultView('pending');
}

function renderSuccess(data) {
  const { summary, files, generalTable, pendingTable, differenceTable, followUp } = data;
  currentRunId = data.runId || '';
  currentFiles = files || null;
  resultPlaceholder.classList.add('is-hidden');
  resultContent.classList.remove('is-hidden');
  resultTitle.textContent = 'Comparativo concluido';
  resultBadge.textContent = 'Concluido';
  resultBadge.classList.remove('is-error');

  resultMetrics.innerHTML = [
    metricCard('Pedido', summary.requestedRowCount ?? 0),
    metricCard('PDFs oficiais', summary.pdfCount),
    metricCard('Diferenças', summary.differenceRowCount ?? 0),
    metricCard('Pendências abertas', summary.pendingRowCount ?? 0),
    metricCard('Base de apoio', summary.finalRowCount),
    metricCard('Tabelas lidas', summary.tableCount ?? 0),
    metricCard('Avisos', summary.warnings.length)
  ].join('');

  if ((summary.differenceRowCount ?? 0) > 0) {
    resultInsightTitle.textContent = `${summary.differenceRowCount} diferença(s) foram identificadas`;
    resultInsightText.textContent = 'A aba "Diferenças entre as bases" mostra tudo o que existe só em um dos lados. Só o que continuar sem prova oficial entra em "Pendências remanescentes".';
  } else if ((summary.pendingRowCount ?? 0) > 0) {
    resultInsightTitle.textContent = `${summary.pendingRowCount} pendência(s) continuam em aberto`;
    resultInsightText.textContent = 'Use a aba "Pendências remanescentes" para cobrar apenas o que estava no pedido e não apareceu com comprovação oficial no retorno do cliente.';
  } else {
    resultInsightTitle.textContent = 'Tudo o que foi pedido foi comprovado';
    resultInsightText.textContent = 'Nenhuma pendência remanescente foi encontrada. O retorno do cliente apresentou comprovação oficial para todas as notas solicitadas.';
  }

  resultFiles.innerHTML = [
    files.differenceFileUrl
      ? fileCard('Diferenças entre as bases', files.differenceFileName, files.differenceFileUrl)
      : '',
    files.pendingFileUrl
      ? fileCard('Pendências remanescentes', files.pendingFileName, files.pendingFileUrl)
      : '',
    files.generalFileUrl
      ? fileCard('Base padronizada do retorno', files.generalFileName, files.generalFileUrl)
      : '',
    fileCard('Pacote completo com abas', files.workbookName, files.workbookUrl),
    fileCard('Resumo do comparativo', files.summaryName, files.summaryUrl),
    `<div class="help-callout"><strong>Leitura recomendada</strong><p>Primeiro baixe as pendências remanescentes. Depois, use a base padronizada do retorno apenas como apoio para conferência.</p></div>`,
    `<div class="help-callout"><strong>Follow-up pelo Outlook Web</strong><p>Se ainda existirem pendências, use o bloco "E-mail de cobranca pronto para editar e enviar" para abrir o Outlook Web com a mensagem preenchida e baixar a planilha para anexo manual.</p></div>`,
    `<div class="help-callout"><strong>Pasta de saída da automação</strong><p><code>${escapeHtml(files.outputDirectory)}</code></p></div>`
  ].filter(Boolean).join('');

  resultDifferenceCount.textContent = `${differenceTable?.rowCount || 0} linha(s)`;
  resultDifferenceTabCount.textContent = resultDifferenceCount.textContent;
  resultDifferenceTable.innerHTML = differenceTable?.rowCount
    ? tableCardMarkup(differenceTable, { showSelector: false, limitRows: 60, note: differenceTable.truncated ? 'Mostrando as primeiras 60 linhas das diferenças entre as bases.' : '' })
    : '<div class="empty-inline">Nenhuma diferença estrutural foi identificada entre pedido e retorno.</div>';

  resultPendingCount.textContent = `${pendingTable?.rowCount || 0} linha(s)`;
  resultPendingTabCount.textContent = resultPendingCount.textContent;
  resultPendingTable.innerHTML = pendingTable?.rowCount
    ? tableCardMarkup(pendingTable, { showSelector: false, limitRows: 60, note: pendingTable.truncated ? 'Mostrando as primeiras 60 linhas das pendências remanescentes.' : '' })
    : '<div class="empty-inline">Nenhuma pendência remanescente foi identificada.</div>';

  resultCombinedCount.textContent = `${generalTable?.rowCount || 0} linha(s)`;
  resultCombinedTabCount.textContent = resultCombinedCount.textContent;
  resultCombinedTable.innerHTML = generalTable
    ? tableCardMarkup(generalTable, { showSelector: false, limitRows: 60, note: generalTable.truncated ? 'Mostrando as primeiras 60 linhas da base padronizada do retorno.' : '' })
    : '<div class="empty-inline">Nenhuma base padronizada foi gerada nesta execução.</div>';

  warningsCount.textContent = String(summary.warnings.length);
  warningsList.innerHTML = summary.warnings.length
    ? summary.warnings.map((warning) => `<li>${escapeHtml(warning)}</li>`).join('')
    : '<li>Nenhum aviso relevante. O comparativo foi concluido sem alertas automaticos.</li>';

  if (followUp?.available) {
    setFollowUpState(followUp);
  } else {
    setFollowUpManualState(summary, files);
  }
  setResultView((differenceTable?.rowCount || 0) > 0 ? 'difference' : ((pendingTable?.rowCount || 0) > 0 ? 'pending' : 'general'));
  resultContent.scrollIntoView({ behavior: 'smooth', block: 'start' });

  if (files.pendingFileUrl || files.generalFileUrl) {
    window.setTimeout(() => {
      downloadToBrowser(files.pendingFileUrl || files.generalFileUrl, files.pendingFileName || files.generalFileName).catch(() => {
        // fallback via cards
      });
    }, 200);
  }
}

function renderError(message) {
  currentRunId = '';
  currentFiles = null;
  resultPlaceholder.classList.add('is-hidden');
  resultContent.classList.remove('is-hidden');
  resultTitle.textContent = 'Falha no comparativo';
  resultBadge.textContent = 'Erro';
  resultBadge.classList.add('is-error');
  resultInsightTitle.textContent = 'Não foi possível concluir o comparativo';
  resultInsightText.textContent = 'Revise os dois portais de anexos e tente novamente.';
  resultMetrics.innerHTML = '';
  resultFiles.innerHTML = '<div class="help-callout"><strong>Detalhe</strong><p>Revise os arquivos enviados nos dois portais e tente novamente.</p></div>';
  resultDifferenceTabCount.textContent = '0 linha(s)';
  resultPendingTabCount.textContent = '0 linha(s)';
  resultCombinedTabCount.textContent = '0 linha(s)';
  resultDifferenceCount.textContent = '0 linha(s)';
  resultPendingCount.textContent = '0 linha(s)';
  resultCombinedCount.textContent = '0 linha(s)';
  resultDifferenceTable.innerHTML = '<div class="empty-inline">As diferenças entre as bases não ficaram disponíveis por causa da falha na execução.</div>';
  resultPendingTable.innerHTML = '<div class="empty-inline">As pendências remanescentes não ficaram disponíveis por causa da falha na execução.</div>';
  resultCombinedTable.innerHTML = '<div class="empty-inline">A base padronizada não ficou disponível por causa da falha na execução.</div>';
  warningsCount.textContent = '1';
  warningsList.innerHTML = `<li>${escapeHtml(message)}</li>`;
  setFollowUpPlaceholder(
    'O e-mail de cobrança não ficou disponível.',
    'Revise os dois arquivos do comparativo e tente novamente. Quando a execução terminar com sucesso, a mensagem de follow-up será montada aqui.'
  );
  setResultView('difference');
}

function setResultView(view) {
  const normalizedView = ['difference', 'pending', 'general'].includes(view) ? view : 'difference';
  [resultTabDifference, resultTabPending, resultTabGeneral].forEach((button) => {
    const isActive = button.dataset.resultView === normalizedView;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-selected', isActive ? 'true' : 'false');
  });

  Array.from(document.querySelectorAll('.result-tab-panel')).forEach((panel) => {
    panel.classList.toggle('is-active', panel.dataset.resultPanel === normalizedView);
  });
}

function attachmentOption(attachment, side) {
  const isRelevant = attachment.kind === 'spreadsheet' || attachment.kind === 'pdf';
  return `
    <label class="attachment-option${isRelevant ? '' : ' attachment-option-muted'}">
      <input class="attachment-checkbox" data-side="${escapeHtml(side)}" type="checkbox" value="${escapeHtml(String(attachment.id))}" ${isRelevant ? 'checked' : ''}>
      <div class="attachment-copy">
        <strong>${escapeHtml(attachment.filename || 'anexo sem nome')}</strong>
        <span>${escapeHtml(formatAttachmentMeta(attachment))}</span>
      </div>
      <span class="attachment-kind attachment-kind-${escapeHtml(attachment.kind || 'other')}">
        ${escapeHtml(formatAttachmentKind(attachment.kind))}
      </span>
    </label>
  `;
}

function tablePreview(table, index, side) {
  return tableCardMarkup(table, {
    showSelector: true,
    checkboxId: `${side}-table-${index + 1}`,
    limitRows: 8,
    side,
  });
}

function tableCardMarkup(table, options = {}) {
  const {
    showSelector = false,
    checkboxId = '',
    limitRows = 30,
    note = '',
    side = '',
  } = options;

  const limitedRows = (table.rows || []).slice(0, limitRows);
  const infoNote = note || (table.truncated || (table.rows || []).length > limitRows
    ? `Mostrando apenas as primeiras ${limitRows} linhas desta tabela.`
    : '');

  const selector = showSelector
    ? `
        <label class="table-selector" for="${checkboxId}">
          <input class="table-checkbox" data-side="${escapeHtml(side)}" id="${checkboxId}" type="checkbox" value="${escapeHtml(String(table.id))}" checked>
          <span>Selecionar tabela</span>
        </label>
      `
    : '';

  const headerHtml = (table.headers || []).length
    ? `<tr>${table.headers.map((header) => `<th>${escapeHtml(header || '-')}</th>`).join('')}</tr>`
    : '';

  const rowsHtml = limitedRows.length
    ? limitedRows.map((row) => `<tr>${row.map((cell, index) => `<td>${escapeHtml(formatTableCell(cell, table.headers?.[index]))}</td>`).join('')}</tr>`).join('')
    : `<tr><td colspan="${Math.max(table.columnCount || 1, 1)}">A tabela nao possui linhas de dados.</td></tr>`;

  return `
    <article class="table-card">
      <div class="table-card-header">
        <div>
          <strong>${escapeHtml(table.title || 'Tabela')}</strong>
          <span>${escapeHtml(`${table.rowCount || limitedRows.length} linha(s) x ${table.columnCount || (table.headers || []).length} coluna(s)`)}</span>
        </div>
        ${selector}
      </div>
      <div class="table-scroll">
        <table class="email-table-preview">
          <thead>${headerHtml}</thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
      ${infoNote ? `<p class="table-note">${escapeHtml(infoNote)}</p>` : ''}
    </article>
  `;
}

function metricCard(label, value) {
  return `
    <article class="metric-card">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(String(value))}</strong>
    </article>
  `;
}

function fileCard(label, fileName, href) {
  return `
    <button class="result-file" type="button" data-download-url="${escapeHtml(href)}" data-download-name="${escapeHtml(fileName)}">
      <div>
        <strong>${escapeHtml(label)}</strong>
        <small>${escapeHtml(fileName)}</small>
      </div>
      <code>Baixar</code>
    </button>
  `;
}

function setFollowUpState(followUp) {
  hideFollowUpFeedback();

  if (!followUp || !followUp.available) {
    setFollowUpManualState();
    return;
  }

  setFollowUpEnabled(true);
  followUpFrom.value = followUp.from || '';
  followUpTo.value = followUp.to || '';
  followUpCc.value = followUp.cc || '';
  followUpSubject.value = followUp.subject || '';
  followUpMessage.value = followUp.message || '';
  const label = followUp.itemLabel || 'item(ns)';
  const modeLabel = followUp.mode === 'difference' ? 'diferença(s)' : label;
  followUpHint.textContent = `${followUp.pendingCount || 0} ${modeLabel} entrarão no e-mail com a planilha anexa.`;
  followUpMeta.innerHTML = [
    metaCard('Remetente', followUp.from || 'Conta padrão do Outlook'),
    metaCard(followUp.mode === 'difference' ? 'Diferenças no follow-up' : 'Pendências no follow-up', `${followUp.pendingCount || 0} linha(s)`),
    metaCard('Planilha anexada', (followUp.attachmentNames || []).join(', ') || 'Nenhum anexo definido'),
    metaCard('Assinatura em imagem', getSignatureLabel()),
  ].join('');
  updateSignatureHint();
}

function setFollowUpPlaceholder(title, description) {
  hideFollowUpFeedback();
  setFollowUpEnabled(false);
  followUpFrom.value = '';
  followUpTo.value = '';
  followUpCc.value = '';
  followUpSubject.value = '';
  followUpMessage.value = '';
  followUpHint.textContent = title;
  followUpMeta.innerHTML = `
    <div class="followup-placeholder">
      <strong>${escapeHtml(title)}</strong>
      <p>${escapeHtml(description)}</p>
    </div>
  `;
  updateSignatureHint();
}

function setFollowUpManualState(summary = null, files = null) {
  hideFollowUpFeedback();
  setFollowUpEnabled(true);
  if (!followUpFrom.value) {
    followUpFrom.value = 'carlos.siqueira@y3gestao.com.br';
  }
  if (!followUpSubject.value) {
    followUpSubject.value = buildManualFollowUpSubject(summary);
  }
  if (!followUpMessage.value) {
    followUpMessage.value = buildManualFollowUpMessage(summary);
  }
  followUpHint.textContent = 'Você pode editar livremente os campos abaixo. Se o Outlook não estiver disponível, use os botões de copiar logo abaixo.';
  followUpMeta.innerHTML = `
    <div class="followup-placeholder">
      <strong>Modo manual habilitado</strong>
      <p>Não houve uma cobrança automática pronta nesta execução, mas você ainda pode editar os campos, anexar uma assinatura em imagem e abrir um rascunho manual.</p>
    </div>
    ${metaCard('Planilha sugerida', resolveFollowUpAttachmentName(files))}
    ${metaCard('Assinatura em imagem', getSignatureLabel())}
  `;
  updateSignatureHint();
}

function setFollowUpEnabled(isEnabled) {
  [followUpFrom, followUpTo, followUpCc, followUpSubject, followUpMessage, followUpSignatureImage].forEach((field) => {
    if (!field) {
      return;
    }
    field.disabled = !isEnabled;
  });
  if (followUpWebButton) {
    followUpWebButton.disabled = !isEnabled;
  }
  followUpDraftButton.disabled = !isEnabled;
  followUpSendButton.disabled = !isEnabled;
  if (copyFollowUpSubjectButton) {
    copyFollowUpSubjectButton.disabled = !isEnabled;
  }
  if (copyFollowUpMessageButton) {
    copyFollowUpMessageButton.disabled = !isEnabled;
  }
}

async function openFollowUpInWeb() {
  if (!currentRunId) {
    showFollowUpFeedback('error', 'Processe o comparativo antes de tentar abrir o Outlook Web.');
    return;
  }

  const from = followUpFrom.value.trim();
  const to = followUpTo.value.trim();
  const cc = followUpCc.value.trim();
  const subject = followUpSubject.value.trim();
  const message = followUpMessage.value;

  if (!to) {
    showFollowUpFeedback('error', 'Informe ao menos um destinatario no campo "Para".');
    followUpTo.focus();
    return;
  }

  if (!subject) {
    showFollowUpFeedback('error', 'Informe o assunto do e-mail antes de continuar.');
    followUpSubject.focus();
    return;
  }

  setFollowUpBusy(true, 'web');

  try {
    const attachmentInfo = resolveFollowUpAttachmentInfo(currentFiles);
    const signatureFile = followUpSignatureImage?.files?.[0] || null;
    const composeMessage = buildWebComposeMessage(message, {
      attachmentName: attachmentInfo.name,
      attachmentUrl: attachmentInfo.absoluteUrl,
      signatureName: signatureFile?.name || '',
    });
    const composeUrl = buildOutlookWebComposeUrlClient({ to, cc, subject, body: composeMessage });

    openUrlInNewTab(composeUrl);

    if (attachmentInfo.relativeUrl) {
      downloadToBrowser(attachmentInfo.relativeUrl, attachmentInfo.name || resolveFollowUpAttachmentName(currentFiles)).catch(() => {
        // fallback silencioso
      });
    }

    if (signatureFile) {
      downloadLocalFile(signatureFile);
    }

    showFollowUpFeedback('success', attachmentInfo.relativeUrl
      ? 'Outlook Web aberto com a mensagem preenchida. A planilha foi baixada neste aparelho para voce anexar antes do envio.'
      : 'Outlook Web aberto com a mensagem preenchida.');
  } catch (error) {
    showFollowUpFeedback('error', error.message);
  } finally {
    setFollowUpBusy(false, 'web');
  }
}

async function sendFollowUp(action) {
  if (!currentRunId) {
    showFollowUpFeedback('error', 'Processe o comparativo antes de tentar abrir ou enviar o follow-up.');
    return;
  }

  const from = followUpFrom.value.trim();
  const to = followUpTo.value.trim();
  const cc = followUpCc.value.trim();
  const subject = followUpSubject.value.trim();
  const message = followUpMessage.value;

  if (!to) {
    showFollowUpFeedback('error', 'Informe ao menos um destinatário no campo "Para".');
    followUpTo.focus();
    return;
  }

  if (!subject) {
    showFollowUpFeedback('error', 'Informe o assunto do e-mail antes de continuar.');
    followUpSubject.focus();
    return;
  }

  let signatureImage = null;
  try {
    signatureImage = await buildSignaturePayload();
  } catch (error) {
    showFollowUpFeedback('error', error.message);
    return;
  }
  setFollowUpBusy(true, action);

  try {
    const response = await fetch(apiUrl(`/api/runs/${currentRunId}/follow-up/outlook`), {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        action,
        from,
        to,
        cc,
        subject,
        message,
        signatureImage,
      }),
    });

    const data = await response.json();
    if (!response.ok || !data.ok) {
      throw new Error(data.error || 'Não foi possível acionar o Outlook para o follow-up.');
    }

    showFollowUpFeedback('success', data.message || (action === 'send'
      ? 'E-mail enviado pelo Outlook.'
      : 'Rascunho aberto no Outlook.'));
  } catch (error) {
    showFollowUpFeedback('error', error.message);
  } finally {
    setFollowUpBusy(false, action);
  }
}

async function openFollowUpInMailApp() {
  if (!currentRunId) {
    showFollowUpFeedback('error', 'Processe o comparativo antes de tentar abrir o aplicativo de e-mail.');
    return;
  }

  const to = followUpTo.value.trim();
  const cc = followUpCc.value.trim();
  const subject = followUpSubject.value.trim();
  const message = followUpMessage.value;

  if (!to) {
    showFollowUpFeedback('error', 'Informe ao menos um destinatario no campo "Para".');
    followUpTo.focus();
    return;
  }

  if (!subject) {
    showFollowUpFeedback('error', 'Informe o assunto do e-mail antes de continuar.');
    followUpSubject.focus();
    return;
  }

  setFollowUpBusy(true, 'draft');

  try {
    const attachmentInfo = resolveFollowUpAttachmentInfo(currentFiles);
    const signatureFile = followUpSignatureImage?.files?.[0] || null;
    const composeMessage = buildMailAppMessage(message, {
      attachmentName: attachmentInfo.name,
      attachmentUrl: attachmentInfo.absoluteUrl,
      signatureName: signatureFile?.name || '',
    });
    const mailtoUrl = buildMailtoUrl({ to, cc, subject, body: composeMessage });

    openMailClient(mailtoUrl);

    if (attachmentInfo.relativeUrl) {
      downloadToBrowser(attachmentInfo.relativeUrl, attachmentInfo.name || resolveFollowUpAttachmentName(currentFiles)).catch(() => {
        // fallback silencioso
      });
    }

    if (signatureFile) {
      downloadLocalFile(signatureFile);
    }

    showFollowUpFeedback('success', attachmentInfo.relativeUrl
      ? 'Aplicativo de e-mail aberto com a mensagem preenchida. A planilha foi baixada neste aparelho para voce anexar antes do envio.'
      : 'Aplicativo de e-mail aberto com a mensagem preenchida.');
  } catch (error) {
    showFollowUpFeedback('error', error.message);
  } finally {
    setFollowUpBusy(false, 'draft');
  }
}

function setFollowUpBusy(isBusy, action) {
  followUpDraftButton.disabled = isBusy;
  followUpSendButton.disabled = isBusy;
  if (copyFollowUpSubjectButton) {
    copyFollowUpSubjectButton.disabled = isBusy;
  }
  if (copyFollowUpMessageButton) {
    copyFollowUpMessageButton.disabled = isBusy;
  }

  if (!isBusy) {
    followUpDraftButton.querySelector('span').textContent = 'Abrir rascunho no Outlook';
    followUpDraftButton.querySelector('small').textContent = 'Gera o e-mail para revisar antes do envio';
    followUpSendButton.querySelector('span').textContent = 'Enviar agora pelo Outlook';
    followUpSendButton.querySelector('small').textContent = 'Tenta enviar automaticamente usando o Outlook, quando disponível';
    return;
  }

  if (action === 'send') {
    followUpSendButton.querySelector('span').textContent = 'Enviando pelo Outlook...';
    followUpSendButton.querySelector('small').textContent = 'Aguarde enquanto o Outlook processa o follow-up';
  } else {
    followUpDraftButton.querySelector('span').textContent = 'Abrindo rascunho...';
    followUpDraftButton.querySelector('small').textContent = 'Aguarde enquanto o Outlook monta o e-mail';
  }
}

function buildManualFollowUpSubject(summary) {
  if ((summary?.pendingRowCount ?? 0) > 0) {
    return 'Reforço de tratativa das NFs pendentes';
  }
  if ((summary?.differenceRowCount ?? 0) > 0) {
    return 'Revisão das diferenças encontradas entre pedido e retorno';
  }
  return 'Revisão do comparativo de NFs';
}

function buildManualFollowUpMessage(summary) {
  const pendingCount = Number(summary?.pendingRowCount || 0);
  const differenceCount = Number(summary?.differenceRowCount || 0);
  const bodyLine = pendingCount > 0
    ? `Após a conferência entre o relatório enviado e o retorno recebido, identificamos ${pendingCount} pendência(s) que continuam sem comprovação oficial.`
    : differenceCount > 0
      ? `Após a conferência entre o relatório enviado e o retorno recebido, identificamos ${differenceCount} diferença(s) entre as duas bases.`
      : 'Segue o comparativo atualizado para revisão.';
  const requestLine = pendingCount > 0
    ? 'Solicitamos, por gentileza, a tratativa dessas pendências e o reenvio da documentação correspondente.'
    : 'Solicitamos, por gentileza, a revisão das diferenças encontradas e o reenvio da documentação correspondente, quando aplicável.';

  return [
    'Prezados,',
    '',
    bodyLine,
    '',
    requestLine,
    '',
    'Ficamos no aguardo.',
    '',
    'Atenciosamente,',
    'Y3 Gestão Telecom',
  ].join('\n');
}

function resolveFollowUpAttachmentName(files) {
  if (!files) {
    return 'A definir';
  }
  return files.pendingFileName || files.differenceFileName || files.generalFileName || files.workbookName || 'A definir';
}

function getSignatureLabel() {
  return followUpSignatureImage?.files?.[0]?.name || 'Nenhuma imagem selecionada';
}

function updateSignatureHint() {
  if (!followUpSignatureHint) {
    return;
  }
  followUpSignatureHint.textContent = followUpSignatureImage?.files?.[0]
    ? `Imagem selecionada: ${followUpSignatureImage.files[0].name}`
    : 'Se necessário, anexe uma imagem de assinatura para entrar no e-mail.';
}

async function buildSignaturePayload() {
  const file = followUpSignatureImage?.files?.[0];
  if (!file) {
    return null;
  }

  const dataUrl = await readFileAsDataUrl(file);
  const match = /^data:([^;]+);base64,(.+)$/i.exec(dataUrl);
  if (!match) {
    throw new Error('Não foi possível ler a imagem de assinatura selecionada.');
  }

  return {
    name: file.name,
    type: match[1] || file.type || 'application/octet-stream',
    base64: match[2],
  };
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = () => reject(new Error('Falha ao ler a imagem de assinatura.'));
    reader.readAsDataURL(file);
  });
}

async function copyToClipboard(text, successMessage) {
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
    } else {
      const helper = document.createElement('textarea');
      helper.value = text;
      helper.setAttribute('readonly', 'true');
      helper.style.position = 'fixed';
      helper.style.opacity = '0';
      document.body.appendChild(helper);
      helper.select();
      document.execCommand('copy');
      helper.remove();
    }
    showFollowUpFeedback('success', successMessage);
  } catch (error) {
    showFollowUpFeedback('error', 'Não foi possível copiar automaticamente. Tente copiar manualmente.');
  }
}

function showFollowUpFeedback(kind, message) {
  followUpFeedback.classList.remove('is-hidden', 'is-error', 'is-success');
  followUpFeedback.classList.add(kind === 'error' ? 'is-error' : 'is-success');
  followUpFeedback.textContent = message;
}

function hideFollowUpFeedback() {
  followUpFeedback.classList.add('is-hidden');
  followUpFeedback.classList.remove('is-error', 'is-success');
  followUpFeedback.textContent = '';
}

function metaCard(label, value) {
  return `
    <article class="preview-meta-card">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value || '-')}</strong>
    </article>
  `;
}

function formatAttachmentMeta(attachment) {
  const parts = [];
  if (attachment.contentType) {
    parts.push(attachment.contentType);
  }
  if (attachment.size != null) {
    parts.push(formatBytes(attachment.size));
  }
  return parts.join(' | ') || 'Sem metadados adicionais';
}

function formatAttachmentKind(kind) {
  if (kind === 'spreadsheet') {
    return 'Planilha';
  }
  if (kind === 'pdf') {
    return 'PDF';
  }
  return 'Outro';
}

function formatBytes(bytes) {
  const size = Number(bytes);
  if (!Number.isFinite(size) || size < 1024) {
    return `${size || 0} B`;
  }
  if (size < 1024 * 1024) {
    return `${(size / 1024).toFixed(1)} KB`;
  }
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function formatDateTime(value) {
  if (!value) {
    return '-';
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return String(value);
  }

  return new Intl.DateTimeFormat('pt-BR', {
    dateStyle: 'short',
    timeStyle: 'short'
  }).format(date);
}

function apiUrl(path) {
  return `${apiBase}${path}`;
}

async function downloadToBrowser(relativeUrl, fileName, triggerElement = null) {
  if (!relativeUrl) {
    return;
  }

  const target = triggerElement || resultFiles;
  target.classList?.add('is-busy');

  try {
    const response = await fetch(apiUrl(relativeUrl));
    if (!response.ok) {
      throw new Error('Nao foi possivel baixar o arquivo solicitado.');
    }

    const blob = await response.blob();
    const blobUrl = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = blobUrl;
    link.download = fileName || 'arquivo.xlsx';
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(blobUrl), 5000);
  } catch (error) {
    window.open(apiUrl(relativeUrl), '_blank', 'noopener');
    throw error;
  } finally {
    target.classList?.remove('is-busy');
  }
}

function resolveFollowUpAttachmentInfo(files) {
  if (!files) {
    return { relativeUrl: '', absoluteUrl: '', name: '' };
  }

  const relativeUrl = files.pendingFileUrl || files.differenceFileUrl || files.generalFileUrl || files.workbookUrl || '';
  const name = files.pendingFileName || files.differenceFileName || files.generalFileName || files.workbookName || 'arquivo.xlsx';

  return {
    relativeUrl,
    absoluteUrl: relativeUrl ? toAbsoluteAppUrl(relativeUrl) : '',
    name,
  };
}

function toAbsoluteAppUrl(relativeUrl) {
  return new URL(apiUrl(relativeUrl), window.location.href).toString();
}

function buildOutlookWebComposeUrlClient({ to, cc, subject, body }) {
  return `https://outlook.office.com/mail/deeplink/compose?${buildComposeQueryString({ to, cc, subject, body })}`;
}

function buildMailtoUrl({ to, cc, subject, body }) {
  const encodedTo = String(to || '').trim();
  const query = buildComposeQueryString({ cc, subject, body });
  return query ? `mailto:${encodedTo}?${query}` : `mailto:${encodedTo}`;
}

function buildComposeQueryString(fields) {
  return Object.entries(fields || {})
    .map(([key, value]) => [key, String(value || '').trim()])
    .filter(([, value]) => value)
    .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
    .join('&');
}

function buildWebComposeMessage(baseMessage, { attachmentName = '', attachmentUrl = '', signatureName = '' } = {}) {
  return String(baseMessage || '').trim();
}

function buildMailAppMessage(baseMessage, { attachmentName = '', attachmentUrl = '', signatureName = '' } = {}) {
  return String(baseMessage || '').trim();
}

function downloadLocalFile(file) {
  const blobUrl = URL.createObjectURL(file);
  const link = document.createElement('a');
  link.href = blobUrl;
  link.download = file.name || 'arquivo';
  link.style.display = 'none';
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(blobUrl), 5000);
}

function openUrlInNewTab(url) {
  const link = document.createElement('a');
  link.href = url;
  link.target = '_blank';
  link.rel = 'noopener noreferrer';
  link.style.display = 'none';
  document.body.appendChild(link);
  link.click();
  link.remove();
}

function openMailClient(url) {
  const link = document.createElement('a');
  link.href = url;
  link.style.display = 'none';
  document.body.appendChild(link);
  link.click();
  link.remove();
}

function buildApiCandidates() {
  const currentOrigin = normalizeBaseUrl(window.location.origin);
  const candidates = [];
  if (currentOrigin) {
    candidates.push(currentOrigin);
  }

  for (const fallback of ['http://localhost:3210', 'http://127.0.0.1:3210']) {
    if (!candidates.includes(fallback)) {
      candidates.push(fallback);
    }
  }

  return candidates;
}

function normalizeBaseUrl(value) {
  return String(value || '').trim().replace(/\/+$/, '');
}

function compactFileName(value, maxLength = 52) {
  const name = String(value || '');
  if (name.length <= maxLength) {
    return name;
  }

  const extensionIndex = name.lastIndexOf('.');
  const extension = extensionIndex > 0 ? name.slice(extensionIndex) : '';
  const baseName = extension ? name.slice(0, extensionIndex) : name;
  const head = baseName.slice(0, Math.max(18, Math.floor((maxLength - extension.length - 3) / 2)));
  const tail = baseName.slice(-Math.max(12, maxLength - extension.length - head.length - 3));
  return `${head}...${tail}${extension}`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function formatTableCell(value, header = '') {
  if (value == null || value === '') {
    return '-';
  }

  if (isDateColumnHeader(header)) {
    return formatDateCellValue(value);
  }

  return String(value);
}

function isDateColumnHeader(header) {
  const normalized = String(header || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();

  return normalized.includes('emissao')
    || normalized.includes('data fato gerador')
    || normalized.includes('dia de emissao');
}

function formatDateCellValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return buildBrazilianDateString(
      value.getDate(),
      value.getMonth() + 1,
      value.getFullYear(),
      value.getHours(),
      value.getMinutes(),
    );
  }

  const text = String(value).trim();
  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s](\d{2}):(\d{2})(?::\d{2})?)?/);
  if (isoMatch) {
    return buildBrazilianDateString(
      Number(isoMatch[3]),
      Number(isoMatch[2]),
      Number(isoMatch[1]),
      isoMatch[4] == null ? null : Number(isoMatch[4]),
      isoMatch[5] == null ? null : Number(isoMatch[5]),
    );
  }

  const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::\d{2})?)?$/);
  if (!slashMatch) {
    return text;
  }

  let first = Number(slashMatch[1]);
  let second = Number(slashMatch[2]);
  const year = Number(slashMatch[3].length === 2 ? `20${slashMatch[3]}` : slashMatch[3]);
  const hour = slashMatch[4] == null ? null : Number(slashMatch[4]);
  const minute = slashMatch[5] == null ? null : Number(slashMatch[5]);

  if (first <= 12 && second > 12) {
    [first, second] = [second, first];
  }

  return buildBrazilianDateString(first, second, year, hour, minute);
}

function buildBrazilianDateString(day, month, year, hour = null, minute = null) {
  const base = `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${String(year)}`;
  if (hour == null || minute == null) {
    return base;
  }

  return `${base} ${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
}

function setFollowUpState(followUp) {
  hideFollowUpFeedback();
  setFollowUpEnabled(true);
  followUpFrom.value = followUp.from || '';
  followUpTo.value = followUp.to || '';
  followUpCc.value = followUp.cc || '';
  followUpSubject.value = followUp.subject || '';
  followUpMessage.value = followUp.message || '';

  const modeLabel = followUp.mode === 'difference'
    ? 'diferenca(s)'
    : (followUp.itemLabel || 'pendencia(s)');

  followUpHint.textContent = `${followUp.pendingCount || 0} ${modeLabel} entrarao no e-mail com a planilha anexa.`;
  followUpMeta.innerHTML = [
    metaCard('Remetente', followUp.from || 'Conta padrao do Outlook'),
    metaCard(followUp.mode === 'difference' ? 'Diferencas no follow-up' : 'Pendencias no follow-up', `${followUp.pendingCount || 0} linha(s)`),
    metaCard('Planilha anexada', (followUp.attachmentNames || []).join(', ') || 'Nenhum anexo definido'),
    metaCard('Assinatura em imagem', getSignatureLabel()),
  ].join('');
  updateSignatureHint();
}

function setFollowUpPlaceholder(title, description) {
  hideFollowUpFeedback();
  setFollowUpEnabled(false);
  followUpFrom.value = '';
  followUpTo.value = '';
  followUpCc.value = '';
  followUpSubject.value = '';
  followUpMessage.value = '';
  followUpHint.textContent = title;
  followUpMeta.innerHTML = `
    <div class="followup-placeholder">
      <strong>${escapeHtml(title)}</strong>
      <p>${escapeHtml(description)}</p>
    </div>
  `;
  updateSignatureHint();
}

function setFollowUpManualState(summary = null, files = null) {
  hideFollowUpFeedback();
  setFollowUpEnabled(true);
  if (!followUpFrom.value) {
    followUpFrom.value = 'carlos.siqueira@y3gestao.com.br';
  }
  if (!followUpSubject.value) {
    followUpSubject.value = buildManualFollowUpSubject(summary);
  }
  if (!followUpMessage.value) {
    followUpMessage.value = buildManualFollowUpMessage(summary);
  }

  followUpHint.textContent = 'Voce pode editar livremente os campos abaixo. Use o Outlook Web para montar o rascunho no proprio aparelho.';
  followUpMeta.innerHTML = `
    <div class="followup-placeholder">
      <strong>Modo manual habilitado</strong>
      <p>Nao houve uma cobranca automatica pronta nesta execucao, mas voce ainda pode editar os campos, baixar a planilha sugerida e abrir o Outlook Web para concluir o envio.</p>
    </div>
    ${metaCard('Planilha sugerida', resolveFollowUpAttachmentName(files))}
    ${metaCard('Assinatura em imagem', getSignatureLabel())}
  `;
  updateSignatureHint();
}

function updateSignatureHint() {
  if (!followUpSignatureHint) {
    return;
  }

  followUpSignatureHint.textContent = followUpSignatureImage?.files?.[0]
    ? `Imagem selecionada: ${followUpSignatureImage.files[0].name}. Ela sera baixada para insercao manual no Outlook Web.`
    : 'Se necessario, selecione uma imagem para baixar junto e inserir manualmente no Outlook Web.';
}

function setFollowUpBusy(isBusy, action) {
  if (followUpWebButton) {
    followUpWebButton.disabled = isBusy;
  }
  followUpDraftButton.disabled = isBusy;
  followUpSendButton.disabled = isBusy;
  if (copyFollowUpSubjectButton) {
    copyFollowUpSubjectButton.disabled = isBusy;
  }
  if (copyFollowUpMessageButton) {
    copyFollowUpMessageButton.disabled = isBusy;
  }

  if (!isBusy) {
    updateButtonCopy(followUpWebButton, 'Abrir no Outlook Web', 'Abre o rascunho no navegador e baixa a planilha para anexar');
    updateButtonCopy(followUpDraftButton, 'Abrir no aplicativo de e-mail', 'Abre o e-mail no aplicativo padrao deste aparelho e baixa a planilha');
    updateButtonCopy(followUpSendButton, 'Enviar agora pela maquina', 'Tenta enviar automaticamente usando o Outlook da maquina servidora');
    return;
  }

  if (action === 'web') {
    updateButtonCopy(followUpWebButton, 'Abrindo no Outlook Web...', 'Aguarde enquanto o navegador monta o rascunho');
    return;
  }

  if (action === 'send') {
    updateButtonCopy(followUpSendButton, 'Enviando pela maquina...', 'Aguarde enquanto o Outlook processa o follow-up');
  } else {
    updateButtonCopy(followUpDraftButton, 'Abrindo aplicativo...', 'Aguarde enquanto o aplicativo de e-mail recebe a mensagem');
  }
}

function buildManualFollowUpSubject(summary) {
  if ((summary?.pendingRowCount ?? 0) > 0) {
    return 'Reforco de tratativa das NFs pendentes';
  }
  if ((summary?.differenceRowCount ?? 0) > 0) {
    return 'Revisao das diferencas encontradas entre pedido e retorno';
  }
  return 'Revisao do comparativo de NFs';
}

function buildManualFollowUpMessage(summary) {
  const pendingCount = Number(summary?.pendingRowCount || 0);
  const differenceCount = Number(summary?.differenceRowCount || 0);
  const bodyLine = pendingCount > 0
    ? `Apos a conferencia entre o relatorio enviado e o retorno recebido, identificamos ${pendingCount} pendencia(s) que continuam sem comprovacao oficial.`
    : differenceCount > 0
      ? `Apos a conferencia entre o relatorio enviado e o retorno recebido, identificamos ${differenceCount} diferenca(s) entre as duas bases.`
      : 'Segue o comparativo atualizado para revisao.';
  const requestLine = pendingCount > 0
    ? 'Solicitamos, por gentileza, a tratativa dessas pendencias e o reenvio da documentacao correspondente.'
    : 'Solicitamos, por gentileza, a revisao das diferencas encontradas e o reenvio da documentacao correspondente, quando aplicavel.';

  return [
    'Prezados,',
    '',
    bodyLine,
    '',
    requestLine,
    '',
    'Ficamos no aguardo.',
    '',
    'Atenciosamente,',
    'Y3 Gestao Telecom',
  ].join('\n');
}

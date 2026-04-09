let apiBase = '';
let currentRunId = '';
const embedMode = new URLSearchParams(window.location.search).get('embed') === '1';

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
const resultTabPending = document.getElementById('resultTabPending');
const resultTabGeneral = document.getElementById('resultTabGeneral');
const resultPendingTabCount = document.getElementById('resultPendingTabCount');
const resultCombinedTabCount = document.getElementById('resultCombinedTabCount');
const resultPendingCount = document.getElementById('resultPendingCount');
const resultCombinedCount = document.getElementById('resultCombinedCount');
const resultPendingTable = document.getElementById('resultPendingTable');
const resultCombinedTable = document.getElementById('resultCombinedTable');
const warningsCount = document.getElementById('warningsCount');
const warningsList = document.getElementById('warningsList');
const followUpPanel = document.getElementById('followUpPanel');
const followUpHint = document.getElementById('followUpHint');
const followUpTo = document.getElementById('followUpTo');
const followUpCc = document.getElementById('followUpCc');
const followUpSubject = document.getElementById('followUpSubject');
const followUpMessage = document.getElementById('followUpMessage');
const followUpMeta = document.getElementById('followUpMeta');
const followUpDraftButton = document.getElementById('followUpDraftButton');
const followUpSendButton = document.getElementById('followUpSendButton');
const followUpFeedback = document.getElementById('followUpFeedback');

boot();

async function boot() {
  attachPreviewHandlers();
  attachFormHandler();
  attachResultHandlers();
  attachEmbedReporter();
  setPreviewIdle('request', 'Envie o e-mail ou a planilha do pedido para visualizar o que foi solicitado.');
  setPreviewIdle('response', 'Envie o e-mail ou a planilha do retorno para visualizar o que o cliente mandou.');
  await checkHealth();
}

function attachEmbedReporter() {
  if (!embedMode || window.parent === window) {
    return;
  }

  const postHeight = () => {
    const height = Math.max(
      document.body?.scrollHeight || 0,
      document.documentElement?.scrollHeight || 0,
      document.body?.offsetHeight || 0,
      document.documentElement?.offsetHeight || 0,
    );

    window.parent.postMessage({
      type: 'y3-compare-height',
      height,
    }, '*');
  };

  window.addEventListener('load', postHeight);
  window.addEventListener('resize', postHeight);

  if (typeof ResizeObserver === 'function' && document.body) {
    const observer = new ResizeObserver(() => postHeight());
    observer.observe(document.body);
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
      renderError('A API da automacao nao respondeu. Rode `npm.cmd run web` para processar o comparativo.');
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

  followUpDraftButton.addEventListener('click', async () => {
    await sendFollowUp('draft');
  });

  followUpSendButton.addEventListener('click', async () => {
    if (!window.confirm('Isso vai enviar o follow-up agora pelo Outlook. Deseja continuar?')) {
      return;
    }

    await sendFollowUp('send');
  });
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
      throw new Error(data.error || 'Nao foi possivel carregar a previa desse portal.');
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
    <strong>Nao foi possivel carregar a previa</strong>
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
  resultPlaceholder.classList.add('is-hidden');
  resultContent.classList.remove('is-hidden');
  resultTitle.textContent = 'Processando comparativo';
  resultBadge.textContent = 'Em execucao';
  resultBadge.classList.remove('is-error');
  resultInsightTitle.textContent = 'Comparando pedido e retorno';
  resultInsightText.textContent = 'Estamos conferindo o que foi pedido contra a planilha do cliente, os links oficiais e os PDFs do governo.';
  resultMetrics.innerHTML = '';
  resultFiles.innerHTML = '<div class="help-callout"><strong>Execucao em andamento</strong><p>A automacao esta comparando o pedido original com o retorno do cliente.</p></div>';
  resultPendingTabCount.textContent = '...';
  resultCombinedTabCount.textContent = '...';
  resultPendingCount.textContent = '...';
  resultCombinedCount.textContent = '...';
  resultPendingTable.innerHTML = '<div class="empty-inline">As pendencias remanescentes serao exibidas aqui ao final da execucao.</div>';
  resultCombinedTable.innerHTML = '<div class="empty-inline">A base padronizada do retorno sera exibida aqui ao final da execucao.</div>';
  warningsList.innerHTML = '';
  warningsCount.textContent = '...';
  setFollowUpState(null);
  setResultView('pending');
}

function renderSuccess(data) {
  const { summary, files, generalTable, pendingTable, followUp } = data;
  currentRunId = data.runId || '';
  resultPlaceholder.classList.add('is-hidden');
  resultContent.classList.remove('is-hidden');
  resultTitle.textContent = 'Comparativo concluido';
  resultBadge.textContent = 'Concluido';
  resultBadge.classList.remove('is-error');

  resultMetrics.innerHTML = [
    metricCard('Pedido', summary.requestedRowCount ?? 0),
    metricCard('PDFs oficiais', summary.pdfCount),
    metricCard('Pendencias abertas', summary.pendingRowCount ?? 0),
    metricCard('Base de apoio', summary.finalRowCount),
    metricCard('Tabelas lidas', summary.tableCount ?? 0),
    metricCard('Avisos', summary.warnings.length)
  ].join('');

  if ((summary.pendingRowCount ?? 0) > 0) {
    resultInsightTitle.textContent = `${summary.pendingRowCount} pendencia(s) continuam em aberto`;
    resultInsightText.textContent = 'Use a aba "Pendencias remanescentes" para cobrar apenas o que estava no pedido e nao apareceu com comprovacao oficial no retorno do cliente.';
  } else {
    resultInsightTitle.textContent = 'Tudo o que foi pedido foi comprovado';
    resultInsightText.textContent = 'Nenhuma pendencia remanescente foi encontrada. O retorno do cliente apresentou comprovacao oficial para todas as notas solicitadas.';
  }

  resultFiles.innerHTML = [
    files.pendingFileUrl
      ? fileCard('Pendencias remanescentes', files.pendingFileName, files.pendingFileUrl)
      : '',
    files.generalFileUrl
      ? fileCard('Base padronizada do retorno', files.generalFileName, files.generalFileUrl)
      : '',
    fileCard('Pacote completo com abas', files.workbookName, files.workbookUrl),
    fileCard('Resumo do comparativo', files.summaryName, files.summaryUrl),
    `<div class="help-callout"><strong>Leitura recomendada</strong><p>Primeiro baixe as pendencias remanescentes. Depois, use a base padronizada do retorno apenas como apoio para conferencia.</p></div>`,
    `<div class="help-callout"><strong>Follow-up pelo Outlook</strong><p>Se ainda existirem pendencias, use o bloco "Cobranca pronta para Outlook" para abrir um rascunho ou enviar a cobranca automaticamente com a planilha anexa.</p></div>`,
    `<div class="help-callout"><strong>Pasta de saida da automacao</strong><p><code>${escapeHtml(files.outputDirectory)}</code></p></div>`
  ].filter(Boolean).join('');

  resultPendingCount.textContent = `${pendingTable?.rowCount || 0} linha(s)`;
  resultPendingTabCount.textContent = resultPendingCount.textContent;
  resultPendingTable.innerHTML = pendingTable?.rowCount
    ? tableCardMarkup(pendingTable, { showSelector: false, limitRows: 60, note: pendingTable.truncated ? 'Mostrando as primeiras 60 linhas das pendencias remanescentes.' : '' })
    : '<div class="empty-inline">Nenhuma pendencia remanescente foi identificada.</div>';

  resultCombinedCount.textContent = `${generalTable?.rowCount || 0} linha(s)`;
  resultCombinedTabCount.textContent = resultCombinedCount.textContent;
  resultCombinedTable.innerHTML = generalTable
    ? tableCardMarkup(generalTable, { showSelector: false, limitRows: 60, note: generalTable.truncated ? 'Mostrando as primeiras 60 linhas da base padronizada do retorno.' : '' })
    : '<div class="empty-inline">Nenhuma base padronizada foi gerada nesta execucao.</div>';

  warningsCount.textContent = String(summary.warnings.length);
  warningsList.innerHTML = summary.warnings.length
    ? summary.warnings.map((warning) => `<li>${escapeHtml(warning)}</li>`).join('')
    : '<li>Nenhum aviso relevante. O comparativo foi concluido sem alertas automaticos.</li>';

  setFollowUpState((summary.pendingRowCount ?? 0) > 0 ? followUp : null);
  setResultView((pendingTable?.rowCount || 0) > 0 ? 'pending' : 'general');
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
  resultPlaceholder.classList.add('is-hidden');
  resultContent.classList.remove('is-hidden');
  resultTitle.textContent = 'Falha no comparativo';
  resultBadge.textContent = 'Erro';
  resultBadge.classList.add('is-error');
  resultInsightTitle.textContent = 'Nao foi possivel concluir o comparativo';
  resultInsightText.textContent = 'Revise os dois portais de anexos e tente novamente.';
  resultMetrics.innerHTML = '';
  resultFiles.innerHTML = '<div class="help-callout"><strong>Detalhe</strong><p>Revise os arquivos enviados nos dois portais e tente novamente.</p></div>';
  resultPendingTabCount.textContent = '0 linha(s)';
  resultCombinedTabCount.textContent = '0 linha(s)';
  resultPendingCount.textContent = '0 linha(s)';
  resultCombinedCount.textContent = '0 linha(s)';
  resultPendingTable.innerHTML = '<div class="empty-inline">As pendencias remanescentes nao ficaram disponiveis por causa da falha na execucao.</div>';
  resultCombinedTable.innerHTML = '<div class="empty-inline">A base padronizada nao ficou disponivel por causa da falha na execucao.</div>';
  warningsCount.textContent = '1';
  warningsList.innerHTML = `<li>${escapeHtml(message)}</li>`;
  setFollowUpState(null);
  setResultView('pending');
}

function setResultView(view) {
  const normalizedView = view === 'general' ? 'general' : 'pending';
  [resultTabPending, resultTabGeneral].forEach((button) => {
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
    ? limitedRows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(formatTableCell(cell))}</td>`).join('')}</tr>`).join('')
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
    followUpPanel.classList.add('is-hidden');
    followUpTo.value = '';
    followUpCc.value = '';
    followUpSubject.value = '';
    followUpMessage.value = '';
    followUpMeta.innerHTML = '';
    followUpHint.textContent = 'Anexa automaticamente a planilha de pendencias remanescentes';
    return;
  }

  followUpPanel.classList.remove('is-hidden');
  followUpTo.value = followUp.to || '';
  followUpCc.value = followUp.cc || '';
  followUpSubject.value = followUp.subject || '';
  followUpMessage.value = followUp.message || '';
  followUpHint.textContent = `${followUp.pendingCount || 0} NF(s) pendente(s) entrarao no e-mail com a planilha anexa`;
  followUpMeta.innerHTML = [
    metaCard('Pendencias no follow-up', `${followUp.pendingCount || 0} linha(s)`),
    metaCard('Planilha anexada', (followUp.attachmentNames || []).join(', ') || 'Nenhum anexo definido'),
  ].join('');
}

async function sendFollowUp(action) {
  if (!currentRunId) {
    showFollowUpFeedback('error', 'Processe o comparativo antes de tentar abrir ou enviar o follow-up.');
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

  setFollowUpBusy(true, action);

  try {
    const response = await fetch(apiUrl(`/api/runs/${currentRunId}/follow-up/outlook`), {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        action,
        to,
        cc,
        subject,
        message,
      }),
    });

    const data = await response.json();
    if (!response.ok || !data.ok) {
      throw new Error(data.error || 'Nao foi possivel acionar o Outlook para o follow-up.');
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

function setFollowUpBusy(isBusy, action) {
  followUpDraftButton.disabled = isBusy;
  followUpSendButton.disabled = isBusy;

  if (!isBusy) {
    followUpDraftButton.querySelector('span').textContent = 'Abrir rascunho no Outlook';
    followUpDraftButton.querySelector('small').textContent = 'Gera o e-mail com tabela e anexo para revisar antes do envio';
    followUpSendButton.querySelector('span').textContent = 'Enviar agora pelo Outlook';
    followUpSendButton.querySelector('small').textContent = 'Dispara o follow-up automaticamente usando o Outlook desktop';
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

function formatTableCell(value) {
  if (value == null || value === '') {
    return '-';
  }
  return String(value);
}

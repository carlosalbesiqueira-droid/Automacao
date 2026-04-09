let apiBase = '';
const currentFlow = document.body?.dataset.flow || 'standardize';
const isCompareFlow = currentFlow === 'compare';

const form = document.getElementById('automationForm');
const submitButton = document.getElementById('submitButton');
const previewButton = document.getElementById('previewButton');
const emlInput = document.getElementById('emlFile');
const uploadText = document.getElementById('uploadText');
const healthBadge = document.getElementById('healthBadge');

const previewState = document.getElementById('previewState');
const previewContent = document.getElementById('previewContent');
const previewSubject = document.getElementById('previewSubject');
const previewSource = document.getElementById('previewSource');
const previewMeta = document.getElementById('previewMeta');
const previewBody = document.getElementById('previewBody');
const previewAttachments = document.getElementById('previewAttachments');
const previewAttachmentCount = document.getElementById('previewAttachmentCount');
const previewAttachmentHint = document.getElementById('previewAttachmentHint');
const previewTables = document.getElementById('previewTables');
const previewTableCount = document.getElementById('previewTableCount');
const hubViewButtons = Array.from(document.querySelectorAll('[data-hub-view-trigger]'));
const hubPanels = Array.from(document.querySelectorAll('[data-hub-panel]'));
const hubSummaryCards = Array.from(document.querySelectorAll('[data-hub-summary]'));
const compareHubPanel = document.getElementById('compareHubPanel');
const standardizeHubPanel = document.getElementById('standardizeHubPanel');
const compareFrame = document.getElementById('compareFrame');
const reloadCompareFrameButton = document.getElementById('reloadCompareFrameButton');

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
const resultPendingTable = document.getElementById('resultPendingTable');
const resultPendingCount = document.getElementById('resultPendingCount');
const resultCombinedTable = document.getElementById('resultCombinedTable');
const resultCombinedCount = document.getElementById('resultCombinedCount');
const warningsList = document.getElementById('warningsList');
const warningsCount = document.getElementById('warningsCount');

boot();

async function boot() {
  attachPreviewHandlers();
  attachFormHandler();
  attachResultHandlers();
  attachHubHandlers();
  setPreviewIdle('Envie um arquivo .eml para visualizar o e-mail antes do processamento.');
  const apiReady = await checkHealth();
  syncHubViewFromLocation();
  if (apiReady) {
    redirectToCanonicalPanel();
  }
}

function attachHubHandlers() {
  if (!hubViewButtons.length) {
    return;
  }

  hubViewButtons.forEach((button) => {
    button.addEventListener('click', () => {
      setHubView(button.dataset.hubViewTrigger || 'standardize');
    });
  });

  reloadCompareFrameButton?.addEventListener('click', () => {
    if (!compareFrame) {
      return;
    }

    compareFrame.src = './comparativo.html?embed=1';
    reloadCompareFrameButton.disabled = true;
    window.setTimeout(() => {
      reloadCompareFrameButton.disabled = false;
    }, 1200);
  });

  window.addEventListener('hashchange', syncHubViewFromLocation);
  window.addEventListener('message', (event) => {
    const data = event.data || {};
    if (data.type !== 'y3-compare-height' || !compareFrame) {
      return;
    }

    const nextHeight = Number(data.height || 0);
    if (!Number.isFinite(nextHeight) || nextHeight <= 0) {
      return;
    }

    compareFrame.style.height = `${Math.max(980, Math.ceil(nextHeight) + 24)}px`;
  });
}

function syncHubViewFromLocation() {
  if (!hubViewButtons.length) {
    return;
  }

  const fromHash = String(window.location.hash || '').replace(/^#/, '').toLowerCase();
  const view = fromHash === 'comparativo' || fromHash === 'compare' ? 'compare' : 'standardize';
  setHubView(view, { updateHash: false });
}

function setHubView(view, options = {}) {
  const normalizedView = view === 'compare' ? 'compare' : 'standardize';
  const { updateHash = true } = options;

  hubViewButtons.forEach((button) => {
    const isActive = button.dataset.hubViewTrigger === normalizedView;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-selected', isActive ? 'true' : 'false');
  });

  hubPanels.forEach((panel) => {
    panel.classList.toggle('is-active', panel.dataset.hubPanel === normalizedView);
    panel.classList.toggle('is-hidden', panel.dataset.hubPanel !== normalizedView);
  });

  hubSummaryCards.forEach((card) => {
    card.classList.toggle('is-active', card.dataset.hubSummary === normalizedView);
  });

  if (compareFrame && normalizedView === 'compare' && !compareFrame.src) {
    compareFrame.src = './comparativo.html?embed=1';
  }

  if (updateHash) {
    const nextHash = normalizedView === 'compare' ? '#comparativo' : '#padronizacao';
    if (window.location.hash !== nextHash) {
      history.replaceState(null, '', `${window.location.pathname}${window.location.search}${nextHash}`);
    }
  }

  if (normalizedView === 'compare') {
    compareHubPanel?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  } else {
    standardizeHubPanel?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

function attachPreviewHandlers() {
  emlInput.addEventListener('change', async () => {
    const file = emlInput.files?.[0];
    uploadText.innerHTML = file
      ? `Arquivo selecionado: <strong title="${escapeHtml(file.name)}">${escapeHtml(compactFileName(file.name))}</strong>`
      : 'Clique para escolher um arquivo <code>.eml</code> ou arraste-o para esta area.';

    if (file) {
      await requestPreview();
    } else {
      setPreviewIdle('Envie um arquivo .eml para visualizar o e-mail antes do processamento.');
    }
  });

  previewButton.addEventListener('click', async () => {
    await requestPreview();
  });

  previewAttachments.addEventListener('change', (event) => {
    if (event.target.classList.contains('attachment-checkbox')) {
      updateAttachmentSelectionLabel();
    }
  });

  previewTables.addEventListener('change', (event) => {
    if (event.target.classList.contains('table-checkbox')) {
      updateTableSelectionLabel();
    }
  });
}

function attachFormHandler() {
  submitButton.addEventListener('click', (event) => {
    if (event.defaultPrevented) {
      return;
    }

    if (typeof form.requestSubmit === 'function') {
      event.preventDefault();
      form.requestSubmit();
    }
  });

  form.addEventListener('submit', async (event) => {
    event.preventDefault();

    const file = emlInput.files?.[0];
    if (!file) {
      renderError('Selecione um arquivo .eml antes de processar a automacao.');
      setPreviewIdle('Envie um arquivo .eml para visualizar o e-mail antes do processamento.');
      return;
    }

    if (!await ensureApiReady()) {
      renderError('A API da automacao nao respondeu. Rode `npm.cmd run web` para processar o arquivo.');
      return;
    }

    submitButton.disabled = true;
    submitButton.querySelector('span').textContent = 'Processando automacao...';
    submitButton.querySelector('small').textContent = 'Gerando a planilha com anexos e tabelas selecionadas';
    setResultLoading();

    try {
      const response = await fetch(apiUrl('/api/process'), {
        method: 'POST',
        body: buildFormData()
      });

      const data = await response.json();
      if (!response.ok || !data.ok) {
        throw new Error(data.error || 'Falha ao processar a automacao.');
      }

      renderSuccess(data);
    } catch (error) {
      renderError(error.message);
    } finally {
      submitButton.disabled = false;
      submitButton.querySelector('span').textContent = isCompareFlow ? 'Processar comparativo' : 'Processar padronizacao';
      submitButton.querySelector('small').textContent = isCompareFlow
        ? 'Retornar apenas as pendencias que continuam sem comprovacao oficial'
        : 'Gerar base padronizada com a planilha anexada e os PDFs oficiais encontrados';
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
  setPreviewIdle('A automacao nao respondeu. Rode `npm.cmd run web` e abra a pagina pelo painel local.');
  return false;
}

async function ensureApiReady() {
  if (apiBase) {
    return true;
  }

  return checkHealth();
}

async function requestPreview() {
  const file = emlInput.files?.[0];
  if (!file) {
    setPreviewIdle('Selecione um arquivo .eml para carregar a previa do e-mail.');
    return;
  }

  if (!await ensureApiReady()) {
    renderPreviewError('A API da automacao nao respondeu. Rode `npm.cmd run web` e tente novamente.');
    return;
  }

  setPreviewLoading();

  try {
    const response = await fetch(apiUrl('/api/preview-email'), {
      method: 'POST',
      body: buildFormData({ includeSelections: false })
    });

    const data = await response.json();
    if (!response.ok || !data.ok) {
      throw new Error(data.error || 'Nao foi possivel carregar a previa do e-mail.');
    }

    renderPreview(data.preview);
  } catch (error) {
    renderPreviewError(error.message);
  }
}

function buildFormData({ includeSelections = true } = {}) {
  const payload = new FormData();
  payload.set('mode', 'upload');
  payload.set('output', form.elements.output.value || 'output');

  if (emlInput.files?.[0]) {
    payload.set('emlFile', emlInput.files[0]);
  }

  if (!includeSelections) {
    return payload;
  }

  const selectedAttachmentIds = getSelectedAttachmentIds();
  if (document.querySelector('.attachment-checkbox')) {
    payload.set('attachmentSelectionApplied', 'true');
  }
  selectedAttachmentIds.forEach((id) => {
    payload.append('selectedAttachmentIds', id);
  });

  const selectedTableIds = getSelectedTableIds();
  if (document.querySelector('.table-checkbox')) {
    payload.set('tableSelectionApplied', 'true');
  }
  selectedTableIds.forEach((id) => {
    payload.append('selectedTableIds', id);
  });

  return payload;
}

function getSelectedAttachmentIds() {
  return Array.from(document.querySelectorAll('.attachment-checkbox:checked'))
    .map((input) => input.value)
    .filter(Boolean);
}

function getSelectedTableIds() {
  return Array.from(document.querySelectorAll('.table-checkbox:checked'))
    .map((input) => input.value)
    .filter(Boolean);
}

function setPreviewIdle(message) {
  previewState.classList.remove('is-hidden');
  previewContent.classList.add('is-hidden');
  previewState.innerHTML = `
    <strong>Aguardando anexo</strong>
    <p>${escapeHtml(message)}</p>
  `;
}

function setPreviewLoading() {
  previewState.classList.remove('is-hidden');
  previewContent.classList.add('is-hidden');
  previewState.innerHTML = `
    <strong>Lendo o e-mail anexado</strong>
    <p>Estamos extraindo o conteudo do arquivo .eml para mostrar a previa da operacao.</p>
  `;
}

function renderPreview(preview) {
  previewState.classList.add('is-hidden');
  previewContent.classList.remove('is-hidden');

  previewSubject.textContent = preview.subject || '(sem assunto)';
  previewSource.textContent = 'Arquivo .eml';
  previewBody.textContent = preview.text || '(O e-mail nao possui corpo em texto disponivel.)';
  previewAttachmentCount.textContent = `${preview.attachmentCount} anexo(s)`;
  previewTableCount.textContent = `${preview.tableCount || 0} tabela(s)`;

  previewMeta.innerHTML = [
    metaCard('Remetente', preview.from || '-'),
    metaCard('Destinatario', preview.to || '-'),
    metaCard('Data', formatDateTime(preview.date)),
    metaCard('Origem', preview.mailbox || 'upload .eml')
  ].join('');

  previewAttachments.innerHTML = preview.attachments.length
    ? preview.attachments.map((attachment) => attachmentOption(attachment)).join('')
    : '<div class="empty-inline">Nenhum anexo foi identificado neste e-mail.</div>';

  previewTables.innerHTML = preview.tables.length
    ? preview.tables.map((table, index) => tablePreview(table, index)).join('')
    : '<div class="empty-inline">Nenhuma tabela estruturada foi identificada no corpo do e-mail.</div>';

  updateAttachmentSelectionLabel();
  updateTableSelectionLabel();
}

function renderPreviewError(message) {
  previewState.classList.remove('is-hidden');
  previewContent.classList.add('is-hidden');
  previewState.innerHTML = `
    <strong>Nao foi possivel carregar a previa</strong>
    <p>${escapeHtml(message)}</p>
  `;
}

function updateAttachmentSelectionLabel() {
  const total = document.querySelectorAll('.attachment-checkbox').length;
  const selected = document.querySelectorAll('.attachment-checkbox:checked').length;
  previewAttachmentHint.textContent = total
    ? `${selected} de ${total} anexo(s) selecionado(s) para processar`
    : 'Nenhum anexo encontrado';
}

function updateTableSelectionLabel() {
  const total = document.querySelectorAll('.table-checkbox').length;
  const selected = document.querySelectorAll('.table-checkbox:checked').length;
  previewTableCount.textContent = total
    ? `${selected} de ${total} tabela(s) selecionada(s)`
    : '0 tabela(s)';
}

function setResultLoading() {
  resultPlaceholder.classList.add('is-hidden');
  resultContent.classList.remove('is-hidden');
  resultTitle.textContent = 'Processando solicitacao';
  resultBadge.textContent = 'Em execucao';
  resultBadge.classList.remove('is-error');
  if (resultInsightTitle) {
    resultInsightTitle.textContent = 'Comparando comprovacoes';
  }
  if (resultInsightText) {
    resultInsightText.textContent = 'Estamos validando a tabela pedida contra a planilha do cliente, os links oficiais e os PDFs do governo.';
  }
  resultMetrics.innerHTML = '';
  resultFiles.innerHTML = '<div class="help-callout"><strong>Execucao em andamento</strong><p>A automacao esta montando a planilha final com anexos e tabelas selecionadas.</p></div>';
  if (resultPendingTabCount) {
    resultPendingTabCount.textContent = '...';
  }
  if (resultPendingCount) {
    resultPendingCount.textContent = '...';
  }
  if (resultPendingTable) {
    resultPendingTable.innerHTML = '<div class="empty-inline">A diferenciacao das pendencias sera exibida aqui ao final da execucao.</div>';
  }
  if (resultCombinedTabCount) {
    resultCombinedTabCount.textContent = '...';
  }
  if (resultCombinedCount) {
    resultCombinedCount.textContent = '...';
  }
  if (resultCombinedTable) {
    resultCombinedTable.innerHTML = `<div class="empty-inline">${isCompareFlow ? 'A base padronizada de apoio sera exibida aqui ao final da execucao.' : 'A base padronizada sera exibida aqui ao final da execucao.'}</div>`;
  }
  warningsList.innerHTML = '';
  warningsCount.textContent = '...';
  setResultView(isCompareFlow ? 'pending' : 'general');
}

function renderSuccess(data) {
  const { summary, files, generalTable, pendingTable } = data;
  resultPlaceholder.classList.add('is-hidden');
  resultContent.classList.remove('is-hidden');

  resultTitle.textContent = isCompareFlow ? 'Comparativo concluido' : (summary.subject || 'Execucao concluida');
  resultBadge.textContent = 'Concluido';
  resultBadge.classList.remove('is-error');

  resultMetrics.innerHTML = isCompareFlow
    ? [
        metricCard('PDFs oficiais', summary.pdfCount),
        metricCard('Pendencias abertas', summary.pendingRowCount ?? 0),
        metricCard('Base de apoio', summary.finalRowCount),
        metricCard('Tabelas lidas', summary.tableCount ?? 0),
        metricCard('Avisos', summary.warnings.length)
      ].join('')
    : [
        metricCard('PDFs', summary.pdfCount),
        metricCard('Linhas finais', summary.finalRowCount),
        metricCard('Tabelas', summary.tableCount ?? 0),
        metricCard('Avisos', summary.warnings.length)
      ].join('');

  if (resultInsightTitle && resultInsightText) {
    if ((summary.pendingRowCount ?? 0) > 0) {
      const pendingCount = summary.pendingRowCount ?? 0;
      resultInsightTitle.textContent = `${pendingCount} pendencia(s) continuam em aberto`;
      resultInsightText.textContent = 'Use a aba "Diferenciacao das pendencias" para cobrar apenas as notas que foram pedidas e ainda nao tiveram comprovacao oficial por PDF do governo ou link oficial da NF.';
    } else {
      resultInsightTitle.textContent = 'Tudo o que foi pedido foi comprovado';
      resultInsightText.textContent = 'Nenhuma pendencia remanescente foi encontrada. O cliente apresentou comprovacao oficial para todas as notas solicitadas.';
    }
  }

  resultFiles.innerHTML = buildResultFiles(files);

  if (resultPendingTable && pendingTable) {
    resultPendingCount.textContent = `${pendingTable.rowCount} linha(s)`;
    resultPendingTable.innerHTML = pendingTable.rowCount
      ? tableCardMarkup(pendingTable, {
          showSelector: false,
          limitRows: 60,
          note: pendingTable.truncated ? 'Mostrando as primeiras 60 linhas da diferenciacao das pendencias.' : ''
        })
      : '<div class="empty-inline">Nenhuma pendencia remanescente foi identificada. Tudo o que foi pedido encontrou comprovacao oficial.</div>';
  } else {
    if (resultPendingCount) {
      resultPendingCount.textContent = '0 linha(s)';
    }
    if (resultPendingTable) {
      resultPendingTable.innerHTML = '<div class="empty-inline">Nao foi possivel montar a diferenciacao das pendencias nesta execucao.</div>';
    }
  }
  if (resultPendingTabCount && resultPendingCount) {
    resultPendingTabCount.textContent = resultPendingCount.textContent;
  }

  if (resultCombinedTable && generalTable) {
    resultCombinedCount.textContent = `${generalTable.rowCount} linha(s)`;
    resultCombinedTable.innerHTML = tableCardMarkup(generalTable, {
      showSelector: false,
      limitRows: 60,
      note: generalTable.truncated ? 'Mostrando as primeiras 60 linhas da base padronizada gerada pela automacao.' : ''
    });
  } else {
    if (resultCombinedCount) {
      resultCombinedCount.textContent = '0 linha(s)';
    }
    if (resultCombinedTable) {
      resultCombinedTable.innerHTML = '<div class="empty-inline">Nenhuma base padronizada foi gerada nesta execucao.</div>';
    }
  }
  if (resultCombinedTabCount && resultCombinedCount) {
    resultCombinedTabCount.textContent = resultCombinedCount.textContent;
  }

  warningsCount.textContent = String(summary.warnings.length);
  warningsList.innerHTML = summary.warnings.length
    ? summary.warnings.map((warning) => `<li>${escapeHtml(warning)}</li>`).join('')
    : '<li>Nenhum aviso relevante. A leitura foi concluida sem pendencias automaticas.</li>';

  setResultView(isCompareFlow && (pendingTable?.rowCount || 0) > 0 ? 'pending' : 'general');
  resultContent.scrollIntoView({ behavior: 'smooth', block: 'start' });

  const preferredDownloadUrl = isCompareFlow
    ? (files.pendingFileUrl || files.generalFileUrl)
    : files.generalFileUrl;
  const preferredDownloadName = isCompareFlow
    ? (files.pendingFileName || files.generalFileName)
    : files.generalFileName;

  if (preferredDownloadUrl) {
    window.setTimeout(() => {
      downloadToBrowser(preferredDownloadUrl, preferredDownloadName).catch(() => {
        // Mantemos os botoes visiveis como fallback se o navegador bloquear o download.
      });
    }, 200);
  }
}

function renderError(message) {
  resultPlaceholder.classList.add('is-hidden');
  resultContent.classList.remove('is-hidden');
  resultTitle.textContent = 'Falha na execucao';
  resultBadge.textContent = 'Erro';
  resultBadge.classList.add('is-error');
  if (resultInsightTitle) {
    resultInsightTitle.textContent = 'Nao foi possivel concluir o comparativo';
  }
  if (resultInsightText) {
    resultInsightText.textContent = 'Revise o e-mail selecionado, os anexos marcados e tente novamente.';
  }
  resultMetrics.innerHTML = '';
  resultFiles.innerHTML = '<div class="help-callout"><strong>Detalhe</strong><p>Revise os anexos e tabelas selecionados e tente novamente.</p></div>';
  if (resultPendingTabCount) {
    resultPendingTabCount.textContent = '0 linha(s)';
  }
  if (resultPendingCount) {
    resultPendingCount.textContent = '0 linha(s)';
  }
  if (resultPendingTable) {
    resultPendingTable.innerHTML = '<div class="empty-inline">A diferenciacao das pendencias nao ficou disponivel por causa da falha na execucao.</div>';
  }
  if (resultCombinedTabCount) {
    resultCombinedTabCount.textContent = '0 linha(s)';
  }
  if (resultCombinedCount) {
    resultCombinedCount.textContent = '0 linha(s)';
  }
  if (resultCombinedTable) {
    resultCombinedTable.innerHTML = '<div class="empty-inline">A base padronizada nao ficou disponivel por causa da falha na execucao.</div>';
  }
  warningsCount.textContent = '1';
  warningsList.innerHTML = `<li>${escapeHtml(message)}</li>`;
  setResultView(isCompareFlow ? 'pending' : 'general');
}

function setResultView(view) {
  const normalizedView = view === 'general' ? 'general' : 'pending';
  const buttons = [resultTabPending, resultTabGeneral];
  const panels = Array.from(document.querySelectorAll('.result-tab-panel'));

  if (!buttons.some(Boolean) || !panels.length) {
    return;
  }

  buttons.forEach((button) => {
    const isActive = button?.dataset.resultView === normalizedView;
    button?.classList.toggle('is-active', isActive);
    button?.setAttribute('aria-selected', isActive ? 'true' : 'false');
  });

  panels.forEach((panel) => {
    panel.classList.toggle('is-active', panel.dataset.resultPanel === normalizedView);
  });
}

function attachmentOption(attachment) {
  const isRelevant = attachment.kind === 'spreadsheet' || attachment.kind === 'pdf';
  return `
    <label class="attachment-option${isRelevant ? '' : ' attachment-option-muted'}">
      <input class="attachment-checkbox" type="checkbox" value="${escapeHtml(String(attachment.id))}" ${isRelevant ? 'checked' : ''}>
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

function tablePreview(table, index) {
  return tableCardMarkup(table, {
    showSelector: true,
    checkboxId: `table-${index + 1}`,
    limitRows: 30
  });
}

function tableCardMarkup(table, options = {}) {
  const {
    showSelector = false,
    checkboxId = '',
    limitRows = 30,
    note = ''
  } = options;

  const limitedRows = (table.rows || []).slice(0, limitRows);
  const infoNote = note || (table.truncated || (table.rows || []).length > limitRows
    ? `Mostrando apenas as primeiras ${limitRows} linhas desta tabela.`
    : '');

  const selector = showSelector
    ? `
        <label class="table-selector" for="${checkboxId}">
          <input class="table-checkbox" id="${checkboxId}" type="checkbox" value="${escapeHtml(String(table.id))}" checked>
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

  const info = `${table.rowCount || limitedRows.length} linha(s) x ${table.columnCount || (table.headers || []).length} coluna(s)`;

  return `
    <article class="table-card">
      <div class="table-card-header">
        <div>
          <strong>${escapeHtml(table.title || 'Tabela')}</strong>
          <span>${escapeHtml(info)}</span>
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
    <button
      class="result-file"
      type="button"
      data-download-url="${escapeHtml(href)}"
      data-download-name="${escapeHtml(fileName)}"
    >
      <div>
        <strong>${escapeHtml(label)}</strong>
        <small>${escapeHtml(fileName)}</small>
      </div>
      <code>Baixar</code>
    </button>
  `;
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

function buildResultFiles(files) {
  if (isCompareFlow) {
    return [
      files.pendingFileUrl
        ? fileCard('Diferenciacao das pendencias', files.pendingFileName, files.pendingFileUrl)
        : '',
      files.generalFileUrl
        ? fileCard('Base padronizada de apoio', files.generalFileName, files.generalFileUrl)
        : '',
      fileCard('Pacote completo com abas', files.workbookName, files.workbookUrl),
      fileCard('Resumo do e-mail', files.summaryName, files.summaryUrl),
      `<div class="help-callout"><strong>Downloads do navegador</strong><p>O comparativo sera enviado para a pasta padrao de downloads do seu navegador. Se o download automatico nao abrir, clique em um dos cards acima.</p></div>`,
      `<div class="help-callout"><strong>Pasta de saida da automacao</strong><p><code>${escapeHtml(files.outputDirectory)}</code></p></div>`
    ].filter(Boolean).join('');
  }

  return [
    files.generalFileUrl
      ? fileCard('Base padronizada consolidada', files.generalFileName, files.generalFileUrl)
      : '',
    fileCard('Pacote completo com abas', files.workbookName, files.workbookUrl),
    fileCard('Resumo do e-mail', files.summaryName, files.summaryUrl),
    `<div class="help-callout"><strong>Comparativo em outra aba</strong><p>Se quiser validar apenas o que continua faltando, abra <a class="inline-link" href="./comparativo.html" target="_blank" rel="noopener">a tela de comparativo</a>.</p></div>`,
    `<div class="help-callout"><strong>Pasta de saida da automacao</strong><p><code>${escapeHtml(files.outputDirectory)}</code></p></div>`
  ].filter(Boolean).join('');
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

function redirectToCanonicalPanel() {
  const currentOrigin = normalizeBaseUrl(window.location.origin);
  if (!apiBase || !currentOrigin || currentOrigin === apiBase) {
    return;
  }

  const canonicalUrl = `${apiBase}${window.location.pathname || '/'}${window.location.search || ''}${window.location.hash || ''}`;
  window.location.replace(canonicalUrl);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
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

function formatTableCell(value) {
  if (value == null || value === '') {
    return '-';
  }

  return String(value);
}

const state = {
  apiBase: '',
  currentLot: null,
  currentLines: [],
  currentLineId: '',
  history: [],
  pollTimer: null,
  tifluxPollTimer: null,
  tifluxBatchJob: null,
  filters: {
    empresa: '',
    operadora: '',
    status: '',
  },
};

const healthBadge = document.getElementById('healthBadge');
const screenButtons = Array.from(document.querySelectorAll('[data-bot-screen-trigger]'));
const screens = Array.from(document.querySelectorAll('[data-bot-screen]'));
const uploadForm = document.getElementById('botUploadForm');
const uploadButton = document.getElementById('uploadButton');
const spreadsheetFileInput = document.getElementById('spreadsheetFile');
const spreadsheetFileLabel = document.getElementById('spreadsheetFileLabel');
const uploadGuideTitle = document.getElementById('uploadGuideTitle');
const uploadGuideText = document.getElementById('uploadGuideText');
const uploadGuideList = document.getElementById('uploadGuideList');
const uploadMappingPanel = document.getElementById('uploadMappingPanel');

const currentLotBanner = document.getElementById('currentLotBanner');
const processingMetrics = document.getElementById('processingMetrics');
const processingHint = document.getElementById('processingHint');
const processingStrip = document.getElementById('processingStrip');
const processingTableBody = document.getElementById('processingTableBody');
const processingLineCount = document.getElementById('processingLineCount');

const filterEmpresa = document.getElementById('filterEmpresa');
const filterOperadora = document.getElementById('filterOperadora');
const filterStatus = document.getElementById('filterStatus');
const resultMetrics = document.getElementById('resultMetrics');
const resultDownloads = document.getElementById('resultDownloads');
const resultTableBody = document.getElementById('resultTableBody');
const resultLineCount = document.getElementById('resultLineCount');

const lineDetailPanel = document.getElementById('lineDetailPanel');
const historyList = document.getElementById('historyList');
const tifluxBatchForm = document.getElementById('tifluxBatchForm');
const tifluxBatchButton = document.getElementById('tifluxBatchButton');
const tifluxTickets = document.getElementById('tifluxTickets');
const tifluxHistorico = document.getElementById('tifluxHistorico');
const tifluxImpedimento = document.getElementById('tifluxImpedimento');
const tifluxTratativas = document.getElementById('tifluxTratativas');
const tifluxEstagio = document.getElementById('tifluxEstagio');
const tifluxAuthCode = document.getElementById('tifluxAuthCode');
const tifluxFaturaAssumida = document.getElementById('tifluxFaturaAssumida');
const tifluxBoDt = document.getElementById('tifluxBoDt');
const tifluxRpsNf = document.getElementById('tifluxRpsNf');
const tifluxNfPrefeitura = document.getElementById('tifluxNfPrefeitura');
const tifluxAe = document.getElementById('tifluxAe');
const tifluxImportacao = document.getElementById('tifluxImportacao');
const tifluxEnvio = document.getElementById('tifluxEnvio');
const tifluxConcluido = document.getElementById('tifluxConcluido');
const tifluxAssistantTitle = document.getElementById('tifluxAssistantTitle');
const tifluxAssistantText = document.getElementById('tifluxAssistantText');
const tifluxAssistantList = document.getElementById('tifluxAssistantList');
const tifluxBatchHint = document.getElementById('tifluxBatchHint');
const tifluxBatchMetrics = document.getElementById('tifluxBatchMetrics');
const tifluxBatchLineCount = document.getElementById('tifluxBatchLineCount');
const tifluxBatchTableBody = document.getElementById('tifluxBatchTableBody');

boot();

async function boot() {
  attachEvents();
  setScreen('upload');
  await connectApi();
  await loadHistory();
}

function attachEvents() {
  screenButtons.forEach((button) => {
    button.addEventListener('click', () => {
      setScreen(button.dataset.botScreenTrigger || 'upload');
    });
  });

  spreadsheetFileInput.addEventListener('change', () => {
    const file = spreadsheetFileInput.files?.[0];
    spreadsheetFileLabel.innerHTML = file
      ? `Arquivo selecionado: <strong title="${escapeHtml(file.name)}">${escapeHtml(compactFileName(file.name))}</strong>`
      : 'Aceita CSV, XLSX, XLSM e XLS com parser flexivel de colunas.';
  });

  uploadForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    await handleUpload();
  });

  [filterEmpresa, filterOperadora, filterStatus].forEach((input) => {
    input.addEventListener('change', () => {
      state.filters.empresa = filterEmpresa.value;
      state.filters.operadora = filterOperadora.value;
      state.filters.status = filterStatus.value;
      renderResults();
    });
  });

  processingTableBody.addEventListener('click', handleLineActionClick);
  resultTableBody.addEventListener('click', handleLineActionClick);
  resultDownloads.addEventListener('click', handleDownloadClick);
  lineDetailPanel.addEventListener('click', handleDetailClick);
  historyList.addEventListener('click', handleHistoryClick);
  tifluxBatchForm?.addEventListener('submit', async (event) => {
    event.preventDefault();
    await handleTifluxBatchSubmit();
  });
}

async function connectApi() {
  const candidates = buildApiCandidates();
  for (const candidate of candidates) {
    try {
      const response = await fetch(`${candidate}/health`);
      const data = await response.json();
      if (!response.ok || !data.ok) {
        continue;
      }
      state.apiBase = candidate;
      healthBadge.textContent = 'Backend FastAPI online';
      healthBadge.classList.remove('is-error');
      healthBadge.classList.add('is-ok');
      return true;
    } catch {
      // tenta o proximo host
    }
  }

  state.apiBase = '';
  healthBadge.textContent = 'Backend FastAPI offline';
  healthBadge.classList.remove('is-ok');
  healthBadge.classList.add('is-error');
  showUploadGuideError(
    'Nao consegui falar com a API do BOT DE FATURAS. Rode `python scripts/run_bot_faturas_api.py` antes de usar esta tela.'
  );
  return false;
}

async function apiFetch(path, options = {}) {
  if (!state.apiBase) {
    const ready = await connectApi();
    if (!ready) {
      throw new Error('Backend do BOT DE FATURAS indisponivel.');
    }
  }

  const response = await fetch(`${state.apiBase}${path}`, options);
  const contentType = response.headers.get('content-type') || '';
  const payload = contentType.includes('application/json') ? await response.json() : await response.text();

  if (!response.ok) {
    const message = typeof payload === 'string'
      ? payload
      : payload.detail || payload.error || 'Falha ao comunicar com o backend do BOT DE FATURAS.';
    throw new Error(message);
  }

  return payload;
}

async function handleUpload() {
  const file = spreadsheetFileInput.files?.[0];
  if (!file) {
    showUploadGuideError('Selecione uma planilha antes de criar o lote.');
    return;
  }

  uploadButton.disabled = true;
  uploadButton.querySelector('span').textContent = 'Criando lote...';
  uploadButton.querySelector('small').textContent = 'Validando arquivo, normalizando colunas e iniciando a fila';
  showUploadGuideLoading('Recebendo a planilha e criando o lote operacional...');

  const formData = new FormData();
  formData.set('file', file);

  try {
    const payload = await apiFetch('/v1/lots/upload', {
      method: 'POST',
      body: formData,
    });
    state.currentLot = payload.lot;
    state.currentLines = payload.lines || [];
    renderCurrentLot();
    renderResults();
    renderUploadMapping();
    await loadHistory();
    setScreen('processing');
    schedulePoll();
  } catch (error) {
    showUploadGuideError(error.message);
  } finally {
    uploadButton.disabled = false;
    uploadButton.querySelector('span').textContent = 'Criar lote e iniciar fila';
    uploadButton.querySelector('small').textContent = 'Valida o arquivo, separa linhas invalidas e envia as validas para processamento assincrono';
  }
}

function renderCurrentLot() {
  renderUploadGuide();
  renderUploadMapping();
  renderProcessing();
  renderResults();

  if (state.currentLineId) {
    void loadLineDetail(state.currentLineId);
  }
}

function renderUploadGuide() {
  if (!state.currentLot) {
    return;
  }

  const mappingEntries = Object.entries(state.currentLot.mapping || {});
  uploadGuideTitle.textContent = `Lote ${state.currentLot.id} criado`;
  uploadGuideText.textContent = [
    `${state.currentLot.total_linhas || 0} linha(s) recebidas`,
    `${state.currentLot.linhas_validas || 0} validas`,
    `${state.currentLot.linhas_invalidas || 0} invalidas`,
    `Status: ${state.currentLot.status || '-'}`,
  ].join(' | ');

  const items = [
    {
      title: 'Parser e normalizacao',
      text: mappingEntries.length
        ? `${mappingEntries.length} coluna(s) foram reconhecidas automaticamente para o modelo interno padrao.`
        : 'O parser nao conseguiu mapear colunas suficientes. Revise os cabecalhos do arquivo enviado.',
    },
    {
      title: 'Erros de validacao',
      text: state.currentLot.linhas_invalidas
        ? `${state.currentLot.linhas_invalidas} linha(s) foram marcadas como ERRO_VALIDACAO e nao entraram na fila.`
        : 'Nenhuma linha foi bloqueada por validacao obrigatoria nesta remessa.',
    },
    {
      title: 'Fila de execucao',
      text: state.currentLot.linhas_validas
        ? `${state.currentLot.linhas_validas} linha(s) foram encaminhadas para os workers assincronos.`
        : 'Nao houve linhas validas para enfileirar nesta planilha.',
    },
  ];

  uploadGuideList.innerHTML = items.map((item) => `
    <article class="assistant-item">
      <strong>${escapeHtml(item.title)}</strong>
      <p>${escapeHtml(item.text)}</p>
    </article>
  `).join('');
}

function renderUploadMapping() {
  if (!state.currentLot) {
    uploadMappingPanel.innerHTML = '<div class="empty-inline">Depois do upload, o mapeamento de colunas reconhecidas aparece aqui.</div>';
    return;
  }

  const mappingEntries = Object.entries(state.currentLot.mapping || {});
  const warnings = state.currentLot.warnings || [];
  uploadMappingPanel.innerHTML = `
    <article class="table-card">
      <div class="table-card-header">
        <div>
          <strong>Mapeamento de colunas reconhecidas</strong>
          <span>${mappingEntries.length} correspondencia(s)</span>
        </div>
      </div>
      <div class="token-list">
        ${mappingEntries.length
          ? mappingEntries.map(([field, source]) => `<span class="token"><strong>${escapeHtml(field)}</strong>&nbsp;←&nbsp;${escapeHtml(source)}</span>`).join('')
          : '<div class="empty-inline">Nenhuma coluna foi reconhecida com confianca suficiente.</div>'}
      </div>
      ${warnings.length ? `<p class="table-note">${escapeHtml(warnings.join(' | '))}</p>` : ''}
    </article>
  `;
}

function renderProcessing() {
  if (!state.currentLot) {
    currentLotBanner.innerHTML = '<strong>Nenhum lote ativo</strong><p>Assim que um arquivo for enviado, o lote atual aparece aqui com o status consolidado.</p>';
    processingMetrics.innerHTML = '';
    processingHint.textContent = 'Sem lote em execucao no momento';
    processingStrip.innerHTML = '<div class="empty-inline">Nenhum processamento em andamento.</div>';
    processingTableBody.innerHTML = '<tr><td colspan="7">Nenhum lote carregado.</td></tr>';
    processingLineCount.textContent = '0 linha(s)';
    return;
  }

  const stats = computeLineStats(state.currentLines);
  currentLotBanner.innerHTML = `
    <strong>Lote ativo: ${escapeHtml(state.currentLot.id)}</strong>
    <p>${escapeHtml(state.currentLot.source_filename || '-')} | ${escapeHtml(state.currentLot.status || '-')} | ${stats.processed} de ${stats.total} linha(s) finalizadas</p>
  `;
  processingMetrics.innerHTML = [
    metricCard('Total', stats.total),
    metricCard('Na fila', stats.queued),
    metricCard('Em processamento', stats.processing),
    metricCard('Finalizadas', stats.processed),
    metricCard('Sucesso total', stats.successTotal),
    metricCard('Sucesso parcial', stats.successPartial),
    metricCard('Erro', stats.errors),
  ].join('');
  processingHint.textContent = `${stats.queued} na fila | ${stats.processing} em processamento | ${stats.processed} concluidas`;
  processingStrip.innerHTML = [
    miniStat('Na fila', stats.queued, 'queued'),
    miniStat('Processando', stats.processing, 'processing'),
    miniStat('Sucesso total', stats.successTotal, 'success'),
    miniStat('Sucesso parcial', stats.successPartial, 'partial'),
    miniStat('Erro', stats.errors, 'error'),
  ].join('');
  processingLineCount.textContent = `${state.currentLines.length} linha(s)`;
  processingTableBody.innerHTML = state.currentLines.length
    ? state.currentLines.map((line) => `
      <tr>
        <td>${escapeHtml(String(line.numero_linha_origem || '-'))}</td>
        <td>${escapeHtml(line.empresa || '-')}</td>
        <td>${escapeHtml(line.operadora_padronizada || line.operadora_original || '-')}</td>
        <td>${escapeHtml(line.mes_referencia || line.vencimento || '-')}</td>
        <td>${statusChip(line.status_final || line.status_processamento || '-')}</td>
        <td>${escapeHtml(line.erro_codigo || line.erro_descricao || '-')}</td>
        <td>
          <div class="toolbar-inline">
            <button class="secondary-button secondary-button-inline" type="button" data-action="open-line" data-line-id="${escapeHtml(line.id)}">Detalhe</button>
          </div>
        </td>
      </tr>
    `).join('')
    : '<tr><td colspan="7">Nenhuma linha carregada.</td></tr>';
}

function renderResults() {
  if (!state.currentLot) {
    resultMetrics.innerHTML = '';
    resultDownloads.innerHTML = '<div class="empty-inline">Os downloads do lote aparecem aqui quando um lote for carregado.</div>';
    resultTableBody.innerHTML = '<tr><td colspan="8">Nenhum resultado carregado.</td></tr>';
    resultLineCount.textContent = '0 linha(s)';
    return;
  }

  populateFilterOptions();
  const filteredLines = applyFilters(state.currentLines);
  const stats = computeLineStats(state.currentLines);

  resultMetrics.innerHTML = [
    metricCard('Total de linhas', stats.total),
    metricCard('Sucesso total', stats.successTotal),
    metricCard('Sucesso parcial', stats.successPartial),
    metricCard('Erro', stats.errors),
  ].join('');

  resultDownloads.innerHTML = [
    state.currentLot.download_urls?.report
      ? fileActionCard('Baixar planilha de retorno', 'Planilha consolidada do lote', state.currentLot.download_urls.report)
      : '',
    state.currentLot.download_urls?.archive
      ? fileActionCard('Baixar arquivos do lote em ZIP', 'Pacote com os arquivos baixados e evidencias', state.currentLot.download_urls.archive)
      : '',
    state.currentLot.download_urls?.log
      ? fileActionCard('Baixar log consolidado', 'Resumo estruturado do lote processado', state.currentLot.download_urls.log)
      : '',
  ].filter(Boolean).join('');

  resultLineCount.textContent = `${filteredLines.length} linha(s)`;
  resultTableBody.innerHTML = filteredLines.length
    ? filteredLines.map((line) => `
      <tr>
        <td>${escapeHtml(String(line.numero_linha_origem || '-'))}</td>
        <td>${escapeHtml(line.empresa || '-')}</td>
        <td>${escapeHtml(line.operadora_padronizada || line.operadora_original || '-')}</td>
        <td>${statusChip(line.status_final || line.status_processamento || '-')}</td>
        <td>${escapeHtml(line.pdf_status || '-')}</td>
        <td>${escapeHtml(line.ae_status || '-')}</td>
        <td>${escapeHtml(line.erro_codigo || line.erro_descricao || '-')}</td>
        <td>
          <div class="toolbar-inline">
            <button class="secondary-button secondary-button-inline" type="button" data-action="open-line" data-line-id="${escapeHtml(line.id)}">Detalhe</button>
            <button class="secondary-button secondary-button-inline" type="button" data-action="reprocess-line" data-line-id="${escapeHtml(line.id)}">Reprocessar</button>
          </div>
        </td>
      </tr>
    `).join('')
    : '<tr><td colspan="8">Nenhuma linha corresponde aos filtros atuais.</td></tr>';
}

function populateFilterOptions() {
  const unique = (values) => Array.from(new Set(values.filter(Boolean))).sort((left, right) => left.localeCompare(right));
  const empresas = unique(state.currentLines.map((line) => line.empresa || ''));
  const operadoras = unique(state.currentLines.map((line) => line.operadora_padronizada || line.operadora_original || ''));
  const statuses = unique(state.currentLines.map((line) => line.status_final || line.status_processamento || ''));

  syncSelectOptions(filterEmpresa, empresas, state.filters.empresa, 'Todas');
  syncSelectOptions(filterOperadora, operadoras, state.filters.operadora, 'Todas');
  syncSelectOptions(filterStatus, statuses, state.filters.status, 'Todos');
}

function renderHistory() {
  historyList.innerHTML = state.history.length
    ? state.history.map((lot) => `
      <article class="history-card">
        <div class="history-card-header">
          <div>
            <strong>${escapeHtml(lot.id)}</strong>
            <p>${escapeHtml(lot.source_filename || '-')}</p>
          </div>
          ${statusChip(lot.status || '-')}
        </div>
        <div class="history-card-metrics">
          <span>${escapeHtml(`${lot.total_linhas || 0} linhas`)}</span>
          <span>${escapeHtml(`${lot.sucesso_total || 0} total`)}</span>
          <span>${escapeHtml(`${lot.sucesso_parcial || 0} parcial`)}</span>
          <span>${escapeHtml(`${lot.erro_total || 0} erro`)}</span>
        </div>
        <div class="toolbar-inline">
          <button class="secondary-button secondary-button-inline" type="button" data-history-action="open-lot" data-lot-id="${escapeHtml(lot.id)}">Abrir</button>
          <button class="secondary-button secondary-button-inline" type="button" data-history-action="reprocess-lot" data-lot-id="${escapeHtml(lot.id)}">Reprocessar lote</button>
          ${lot.download_urls?.report ? `<button class="secondary-button secondary-button-inline" type="button" data-download-url="${escapeHtml(lot.download_urls.report)}" data-download-name="resultado_${escapeHtml(lot.id)}.xlsx">Planilha</button>` : ''}
          ${lot.download_urls?.archive ? `<button class="secondary-button secondary-button-inline" type="button" data-download-url="${escapeHtml(lot.download_urls.archive)}" data-download-name="lote_${escapeHtml(lot.id)}.zip">ZIP</button>` : ''}
        </div>
      </article>
    `).join('')
    : '<div class="empty-inline">Nenhum lote encontrado no historico ainda.</div>';
}

async function loadHistory() {
  try {
    const payload = await apiFetch('/v1/lots?limit=40');
    state.history = payload.lots || [];
    renderHistory();
  } catch (error) {
    historyList.innerHTML = `<div class="empty-inline">${escapeHtml(error.message)}</div>`;
  }
}

async function refreshCurrentLot() {
  if (!state.currentLot?.id) {
    return;
  }

  try {
    const payload = await apiFetch(`/v1/lots/${state.currentLot.id}`);
    state.currentLot = payload.lot;
    state.currentLines = payload.lines || [];
    renderCurrentLot();
    if ((state.currentLot.status || '').toUpperCase() === 'PROCESSANDO') {
      schedulePoll();
    } else {
      stopPolling();
      await loadHistory();
    }
  } catch (error) {
    stopPolling();
    currentLotBanner.innerHTML = `<strong>Falha ao atualizar lote</strong><p>${escapeHtml(error.message)}</p>`;
  }
}

function schedulePoll() {
  stopPolling();
  if (!state.currentLot?.id || (state.currentLot.status || '').toUpperCase() !== 'PROCESSANDO') {
    return;
  }
  state.pollTimer = window.setTimeout(() => {
    void refreshCurrentLot();
  }, 3000);
}

function stopPolling() {
  if (state.pollTimer) {
    window.clearTimeout(state.pollTimer);
    state.pollTimer = null;
  }
}

async function handleTifluxBatchSubmit() {
  const tickets = String(tifluxTickets?.value || '')
    .split(/\r?\n/)
    .map((item) => item.trim())
    .filter(Boolean);

  const updates = {
    historico_da_fatura: tifluxHistorico?.value || '',
    impedimento: tifluxImpedimento?.value || '',
    tratativas_observacoes: tifluxTratativas?.value || '',
    estagio: tifluxEstagio?.value || '',
    fatura_assumida_data: tifluxFaturaAssumida?.value || '',
    bo_dt_data: tifluxBoDt?.value || '',
    rps_nf_data: tifluxRpsNf?.value || '',
    nf_prefeitura: tifluxNfPrefeitura?.value || '',
    ae_data: tifluxAe?.value || '',
    importacao_data: tifluxImportacao?.value || '',
    envio_data: tifluxEnvio?.value || '',
    concluido_data: tifluxConcluido?.value || '',
  };
  const authCode = String(tifluxAuthCode?.value || '').replace(/\D+/g, '');

  if (!tickets.length) {
    renderTifluxBatchError('Cole ao menos um ticket para processar no TiFlux.');
    return;
  }

  const hasAnyUpdate = Object.values(updates).some((value) => String(value || '').trim());
  if (!hasAnyUpdate) {
    renderTifluxBatchError('Preencha ao menos um campo para atualizar no TiFlux.');
    return;
  }

  tifluxBatchButton.disabled = true;
  tifluxBatchButton.querySelector('span').textContent = 'Enfileirando lote...';
  tifluxBatchButton.querySelector('small').textContent = 'Preparando execução direta no TiFlux';
  renderTifluxBatchLoading(`Enfileirando ${tickets.length} ticket(s) para atualização direta no TiFlux...`);

  try {
    const payload = await apiFetch('/v1/tiflux/batch/run', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tickets, updates, auth_code: authCode }),
    });
    state.tifluxBatchJob = {
      job_id: payload.job_id,
      status: payload.status,
      ticket_count: payload.ticket_count,
      message: payload.message,
      results: [],
    };
    renderTifluxBatchJob();
    scheduleTifluxBatchPoll();
    setScreen('tiflux');
  } catch (error) {
    renderTifluxBatchError(error.message);
  } finally {
    tifluxBatchButton.disabled = false;
    tifluxBatchButton.querySelector('span').textContent = 'Executar lote no TiFlux';
    tifluxBatchButton.querySelector('small').textContent = 'Reaproveita a sessão salva do TiFlux e tenta reautenticar quando necessário.';
  }
}

async function refreshTifluxBatchJob() {
  if (!state.tifluxBatchJob?.job_id) {
    return;
  }
  try {
    const payload = await apiFetch(`/v1/tiflux/google-sheet/jobs/${state.tifluxBatchJob.job_id}`);
    state.tifluxBatchJob = payload;
    renderTifluxBatchJob();
    if (payload.status === 'queued' || payload.status === 'running') {
      scheduleTifluxBatchPoll();
    } else {
      stopTifluxBatchPoll();
    }
  } catch (error) {
    stopTifluxBatchPoll();
    renderTifluxBatchError(error.message);
  }
}

function scheduleTifluxBatchPoll() {
  stopTifluxBatchPoll();
  if (!state.tifluxBatchJob?.job_id) {
    return;
  }
  if (!['queued', 'running'].includes(String(state.tifluxBatchJob.status || ''))) {
    return;
  }
  state.tifluxPollTimer = window.setTimeout(() => {
    void refreshTifluxBatchJob();
  }, 3000);
}

function stopTifluxBatchPoll() {
  if (state.tifluxPollTimer) {
    window.clearTimeout(state.tifluxPollTimer);
    state.tifluxPollTimer = null;
  }
}

function renderTifluxBatchLoading(message) {
  tifluxAssistantTitle.textContent = 'Executando lote TiFlux';
  tifluxAssistantText.textContent = message;
  tifluxAssistantList.innerHTML = `
    <article class="assistant-item">
      <strong>Status</strong>
      <p>O job foi enviado. Assim que a API devolver o andamento, os tickets aparecem abaixo.</p>
    </article>
  `;
  tifluxBatchHint.textContent = 'Lote enfileirado';
}

function renderTifluxBatchError(message) {
  tifluxAssistantTitle.textContent = 'Falha no lote TiFlux';
  tifluxAssistantText.textContent = message;
  tifluxAssistantList.innerHTML = `
    <article class="assistant-item">
      <strong>Como seguir</strong>
      <p>Revise a sessão do TiFlux, o código de verificação e os tickets informados antes de tentar novamente.</p>
    </article>
  `;
  tifluxBatchHint.textContent = 'Falha no processamento';
}

function renderTifluxBatchJob() {
  const job = state.tifluxBatchJob;
  if (!job) {
    return;
  }

  const jobStatus = String(job.status || '-');
  const jobMessage = job.error || job.message || 'Lote TiFlux em andamento.';
  const jobFailedCount = job.failed ?? (jobStatus === 'failed' ? (job.ticket_count || (job.tickets || []).length || 1) : 0);

  tifluxAssistantTitle.textContent = `Job ${job.job_id || '-'}`;
  tifluxAssistantText.textContent = jobMessage;
  tifluxAssistantList.innerHTML = `
    <article class="assistant-item">
      <strong>Status atual</strong>
      <p>${escapeHtml(jobStatus)}</p>
    </article>
    <article class="assistant-item">
      <strong>Tickets</strong>
      <p>${escapeHtml(String(job.ticket_count || job.processed || 0))} ticket(s) neste lote.</p>
    </article>
    ${jobStatus === 'failed' ? `
      <article class="assistant-item">
        <strong>Motivo da falha</strong>
        <p>${escapeHtml(jobMessage)}</p>
      </article>
    ` : ''}
  `;

  tifluxBatchHint.textContent = `${escapeHtml(jobStatus)} | job ${escapeHtml(job.job_id || '-')}`;
  tifluxBatchMetrics.innerHTML = [
    metricCard('Status', jobStatus),
    metricCard('Tickets', job.ticket_count || (job.tickets || []).length || 0),
    metricCard('OK', job.updated || 0),
    metricCard('Falhas', jobFailedCount),
  ].join('');

  const results = Array.isArray(job.results) ? job.results : [];
  tifluxBatchLineCount.textContent = `${results.length || job.ticket_count || 0} ticket(s)`;
  tifluxBatchTableBody.innerHTML = results.length
    ? results.map((item) => `
      <tr>
        <td>${escapeHtml(item.ticket || '-')}</td>
        <td>${statusChip(item.status || '-')}</td>
        <td>${escapeHtml(item.message || '-')}</td>
        <td>${escapeHtml(item.processed_at || '-')}</td>
        <td>${item.evidence ? `<code>${escapeHtml(compactFileName(item.evidence, 36))}</code>` : '-'}</td>
      </tr>
    `).join('')
    : jobStatus === 'failed'
      ? `<tr><td colspan="5">${escapeHtml(jobMessage)}</td></tr>`
      : '<tr><td colspan="5">Aguardando retorno do lote TiFlux.</td></tr>';
}

async function loadLineDetail(lineId, forceScreen = false) {
  try {
    const payload = await apiFetch(`/v1/lines/${lineId}`);
    state.currentLineId = lineId;
    renderLineDetail(payload);
    if (forceScreen) {
      setScreen('detail');
    }
  } catch (error) {
    lineDetailPanel.innerHTML = `<div class="empty-inline">${escapeHtml(error.message)}</div>`;
    if (forceScreen) {
      setScreen('detail');
    }
  }
}

function renderLineDetail(payload) {
  const line = payload.line || {};
  const downloads = line.download_urls || {};
  lineDetailPanel.innerHTML = `
    <article class="table-card">
      <div class="table-card-header">
        <div>
          <strong>Linha ${escapeHtml(String(line.numero_linha_origem || '-'))}</strong>
          <span>${escapeHtml(line.empresa || '-')} | ${escapeHtml(line.operadora_padronizada || line.operadora_original || '-')}</span>
        </div>
        ${statusChip(line.status_final || line.status_processamento || '-')}
      </div>

      <div class="result-metrics">
        ${metricCard('PDF', line.pdf_status || '-')}
        ${metricCard('AE', line.ae_status || '-')}
        ${metricCard('Erro', line.erro_codigo || '-')}
        ${metricCard('Processado em', line.processado_em || '-')}
      </div>

      <div class="toolbar-inline">
        ${downloads.pdf ? actionButton('Baixar PDF', downloads.pdf, 'pdf') : ''}
        ${downloads.ae ? actionButton('Baixar AE/TXT/ZIP/CSV', downloads.ae, 'ae') : ''}
        ${downloads.screenshot ? actionButton('Screenshot', downloads.screenshot, 'png') : ''}
        ${downloads.html ? actionButton('HTML', downloads.html, 'html') : ''}
        ${downloads.log ? actionButton('Log tecnico', downloads.log, 'jsonl') : ''}
        <button class="secondary-button secondary-button-inline" type="button" data-action="reprocess-line" data-line-id="${escapeHtml(line.id)}">Reprocessar linha</button>
      </div>

      <div class="detail-grid-inner">
        <div class="json-panel">
          <strong>Dados originais</strong>
          <pre>${escapeHtml(JSON.stringify(payload.original_data || {}, null, 2))}</pre>
        </div>
        <div class="json-panel">
          <strong>Dados normalizados</strong>
          <pre>${escapeHtml(JSON.stringify(payload.normalized_data || {}, null, 2))}</pre>
        </div>
      </div>

      <div class="help-callout">
        <strong>Erro / observacao</strong>
        <p>${escapeHtml(line.erro_descricao || line.observacao_execucao || 'Nenhuma observacao relevante para esta linha.')}</p>
      </div>

      <div class="json-panel">
        <strong>Log resumido</strong>
        <pre>${escapeHtml((payload.log_preview || []).join('\n') || 'Sem log tecnico disponível ainda.')}</pre>
      </div>
    </article>
  `;
}

async function handleLineActionClick(event) {
  const trigger = event.target.closest('[data-action]');
  if (!trigger) {
    return;
  }

  const action = trigger.dataset.action;
  const lineId = trigger.dataset.lineId || '';
  if (action === 'open-line' && lineId) {
    await loadLineDetail(lineId, true);
  }
  if (action === 'reprocess-line' && lineId) {
    await reprocessLine(lineId);
  }
}

async function handleDetailClick(event) {
  const trigger = event.target.closest('[data-download-url], [data-action]');
  if (!trigger) {
    return;
  }
  if (trigger.dataset.downloadUrl) {
    await downloadFile(trigger.dataset.downloadUrl, trigger.dataset.downloadName || 'arquivo');
    return;
  }
  if (trigger.dataset.action === 'reprocess-line' && trigger.dataset.lineId) {
    await reprocessLine(trigger.dataset.lineId);
  }
}

async function handleHistoryClick(event) {
  const trigger = event.target.closest('[data-history-action], [data-download-url]');
  if (!trigger) {
    return;
  }

  if (trigger.dataset.downloadUrl) {
    await downloadFile(trigger.dataset.downloadUrl, trigger.dataset.downloadName || 'arquivo');
    return;
  }

  const lotId = trigger.dataset.lotId || '';
  if (!lotId) {
    return;
  }
  if (trigger.dataset.historyAction === 'open-lot') {
    const payload = await apiFetch(`/v1/lots/${lotId}`);
    state.currentLot = payload.lot;
    state.currentLines = payload.lines || [];
    renderCurrentLot();
    setScreen('results');
    schedulePoll();
  }
  if (trigger.dataset.historyAction === 'reprocess-lot') {
    const payload = await apiFetch(`/v1/lots/${lotId}/reprocess`, { method: 'POST' });
    state.currentLot = payload.lot;
    state.currentLines = payload.lines || [];
    renderCurrentLot();
    setScreen('processing');
    schedulePoll();
    await loadHistory();
  }
}

async function handleDownloadClick(event) {
  const trigger = event.target.closest('[data-download-url]');
  if (!trigger) {
    return;
  }
  await downloadFile(trigger.dataset.downloadUrl, trigger.dataset.downloadName || 'arquivo');
}

async function reprocessLine(lineId) {
  try {
    const payload = await apiFetch(`/v1/lines/${lineId}/reprocess`, { method: 'POST' });
    state.currentLineId = lineId;
    await refreshCurrentLot();
    renderLineDetail(payload);
    setScreen('detail');
    schedulePoll();
  } catch (error) {
    lineDetailPanel.innerHTML = `<div class="empty-inline">${escapeHtml(error.message)}</div>`;
    setScreen('detail');
  }
}

async function downloadFile(relativeUrl, fallbackName) {
  const response = await fetch(`${state.apiBase}${relativeUrl}`);
  if (!response.ok) {
    throw new Error('Nao foi possivel baixar o arquivo solicitado.');
  }
  const blob = await response.blob();
  const blobUrl = URL.createObjectURL(blob);
  const fileName = fileNameFromHeaders(response.headers.get('content-disposition')) || fallbackName || 'arquivo';
  const link = document.createElement('a');
  link.href = blobUrl;
  link.download = fileName;
  link.style.display = 'none';
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(blobUrl), 5000);
}

function setScreen(screen) {
  const normalized = screen || 'upload';
  screenButtons.forEach((button) => {
    button.classList.toggle('is-active', button.dataset.botScreenTrigger === normalized);
  });
  screens.forEach((section) => {
    section.classList.toggle('is-active', section.dataset.botScreen === normalized);
  });
}

function computeLineStats(lines) {
  return lines.reduce((accumulator, line) => {
    accumulator.total += 1;
    const status = String(line.status_final || line.status_processamento || '');
    if (status === 'NA_FILA') {
      accumulator.queued += 1;
    } else if (status === 'EM_PROCESSAMENTO') {
      accumulator.processing += 1;
    } else {
      accumulator.processed += 1;
    }
    if (status === 'SUCESSO_TOTAL') {
      accumulator.successTotal += 1;
    } else if (status === 'SUCESSO_PARCIAL') {
      accumulator.successPartial += 1;
    } else if (status === 'ERRO_VALIDACAO' || status === 'ERRO_PROCESSAMENTO') {
      accumulator.errors += 1;
    }
    return accumulator;
  }, {
    total: 0,
    queued: 0,
    processing: 0,
    processed: 0,
    successTotal: 0,
    successPartial: 0,
    errors: 0,
  });
}

function applyFilters(lines) {
  return lines.filter((line) => {
    const matchesEmpresa = !state.filters.empresa || line.empresa === state.filters.empresa;
    const lineOperadora = line.operadora_padronizada || line.operadora_original || '';
    const matchesOperadora = !state.filters.operadora || lineOperadora === state.filters.operadora;
    const lineStatus = line.status_final || line.status_processamento || '';
    const matchesStatus = !state.filters.status || lineStatus === state.filters.status;
    return matchesEmpresa && matchesOperadora && matchesStatus;
  });
}

function syncSelectOptions(select, values, currentValue, defaultLabel) {
  const options = [`<option value="">${escapeHtml(defaultLabel)}</option>`]
    .concat(values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`));
  select.innerHTML = options.join('');
  select.value = currentValue || '';
}

function buildApiCandidates() {
  const currentOrigin = String(window.location.origin || '').replace(/\/+$/, '');
  const candidates = [];

  if (currentOrigin) {
    candidates.push(`${currentOrigin}/bot-faturas-api`);
  }

  for (const fallback of ['http://127.0.0.1:8321', 'http://localhost:8321']) {
    if (!candidates.includes(fallback)) {
      candidates.push(fallback);
    }
  }

  return candidates;
}

function showUploadGuideLoading(message) {
  uploadGuideTitle.textContent = 'Processando upload';
  uploadGuideText.textContent = message;
}

function showUploadGuideError(message) {
  uploadGuideTitle.textContent = 'Falha na operacao';
  uploadGuideText.textContent = message;
  uploadGuideList.innerHTML = `
    <article class="assistant-item">
      <strong>Como seguir</strong>
      <p>Revise o arquivo enviado, confirme se a API esta no ar e tente novamente.</p>
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

function miniStat(label, value, tone) {
  return `
    <article class="mini-stat mini-stat-${escapeHtml(tone)}">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(String(value))}</strong>
    </article>
  `;
}

function fileActionCard(label, description, url) {
  return `
    <button class="result-file" type="button" data-download-url="${escapeHtml(url)}" data-download-name="${escapeHtml(label)}">
      <div>
        <strong>${escapeHtml(label)}</strong>
        <small>${escapeHtml(description)}</small>
      </div>
      <code>Baixar</code>
    </button>
  `;
}

function actionButton(label, url, extension) {
  return `<button class="secondary-button secondary-button-inline" type="button" data-download-url="${escapeHtml(url)}" data-download-name="${escapeHtml(label)}.${escapeHtml(extension)}">${escapeHtml(label)}</button>`;
}

function statusChip(value) {
  return `<span class="status-chip status-${escapeHtml(String(value).toLowerCase())}">${escapeHtml(value)}</span>`;
}

function fileNameFromHeaders(contentDisposition) {
  if (!contentDisposition) {
    return '';
  }
  const match = /filename="?([^"]+)"?/i.exec(contentDisposition);
  return match ? match[1] : '';
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
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

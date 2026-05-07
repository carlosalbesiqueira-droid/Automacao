# Automacao Y3 NFS-e

Projeto local em Node.js para:

- ler um e-mail exportado em `.eml`
- identificar planilhas, tabelas HTML, links e PDFs de NFS-e
- padronizar a base final de notas
- comparar "o que pedimos" x "o que o cliente mandou"
- preparar um follow-up pelo Outlook quando ainda existirem pendencias

## O que o sistema entrega

### 1. Padronizacao

Fluxo principal da home:

- le o e-mail completo
- detecta anexos e tabelas do corpo
- processa a planilha enviada fora do padrao
- acessa hyperlinks encontrados na planilha
- cruza as informacoes com PDFs oficiais e links da NF
- gera uma base padronizada unica
- cria hyperlinks no numero da `NFS-e`

Colunas finais da base padronizada:

- `NFS-e`
- `RPS`
- `RPS`
- `Emissao`
- `Data Fato Gerador`
- `Tomador de Servicos`
- `CNPJ`
- `Intermediario de Servicos`
- `Valor Servicos`
- `Valor Deducao`
- `ISS devido`
- `ISS a pagar`
- `Valor Credito`
- `ISS`
- `Situacao`
- `ISS pago por guia NFS-e`
- `Carta de`
- `No da obra`

### 2. Comparativo

Fluxo de comparativo:

- Portal 1: o que pedimos
- Portal 2: o que o cliente mandou
- aceita `.eml`, `.xlsx`, `.xlsm`, `.xls` e `.csv`
- usa PDF oficial ou link oficial da NF como prova principal
- retorna somente as pendencias remanescentes
- gera follow-up para Outlook com a planilha de pendencias anexada

## Estrutura principal

- [public/index.html](public/index.html): tela principal com padronizacao e acesso ao comparativo
- [public/comparativo.html](public/comparativo.html): tela dedicada de comparativo
- [public/faturas.html](public/faturas.html): tela guiada do bot de faturas com Google Sheets ou upload
- [public/app.js](public/app.js): frontend da home
- [public/comparativo.js](public/comparativo.js): frontend do comparativo
- [public/faturas.js](public/faturas.js): frontend do bot de faturas
- [src/server.js](src/server.js): API web local
- [src/process-email.js](src/process-email.js): orquestracao principal
- [src/output/write-workbook.js](src/output/write-workbook.js): geracao das planilhas finais

## Requisitos

- Windows
- Node.js instalado
- Outlook desktop instalado apenas se quiser abrir/enviar follow-up automaticamente

## Instalacao

```powershell
npm.cmd install
```

## Automacao de Faturas em Python

Ambiente local preparado para a automacao operacional do `BOT DE FATURAS`, com backend FastAPI, persistencia local, fila assincrona, workers Playwright e interface web dedicada no painel.

Arquivos principais:

- `requirements-automation.txt`: dependencias-base da automacao
- `requirements-automation-lock.txt`: versoes instaladas no ambiente atual
- `.venv/`: ambiente virtual local do Python
- `bot_faturas_v2/`: modulo FastAPI do novo BOT DE FATURAS
- `scripts/run_bot_faturas_api.py`: inicializacao da API operacional do bot

Ativacao do ambiente:

```powershell
.\.venv\Scripts\Activate.ps1
```

Validacao rapida:

```powershell
python --version
python -m pip --version
python -m playwright --version
```

Criacao das abas do bot de faturas:

```powershell
.\.venv\Scripts\python.exe .\scripts\setup_bot_faturas.py
```

Diagnostico da base do bot:

```powershell
.\.venv\Scripts\python.exe -m bot_faturas.cli validate
```

Dry-run do bot sem acessar portais:

```powershell
.\.venv\Scripts\python.exe -m bot_faturas.cli run --limit 5 --dry-run
```

Execucao real do bot:

```powershell
.\.venv\Scripts\python.exe -m bot_faturas.cli run --limit 5
```

Pacotes preparados para:

- Google Sheets e formatacao de abas
- login e download em portais via navegador automatizado
- tratamento de planilhas, PDFs, HTML e links
- normalizacao, retries, validacao e logging
- API FastAPI para upload de planilhas, historico, detalhe por linha e reprocessamento
- persistencia local em banco SQLite e armazenamento de evidencias por lote/linha

Estrutura adicionada para o bot de faturas:

- `bot_faturas/config.py`: configuracao do bot por ambiente
- `bot_faturas/sheets.py`: leitura e atualizacao da aba `BOT_FATURAS`
- `bot_faturas/handlers.py`: estrategia generica de portal e perfis conhecidos
- `bot_faturas/runner.py`: orquestracao do processamento
- `bot_faturas/cli.py`: comandos `validate` e `run`
- `scripts/run_bot_faturas.py`: atalho para executar a CLI
- `bot_faturas_v2/api.py`: API FastAPI do BOT DE FATURAS
- `bot_faturas_v2/database.py`: persistencia de lotes e linhas em SQLite
- `bot_faturas_v2/parser.py`: parser flexivel de CSV/XLSX com normalizacao de colunas
- `bot_faturas_v2/queue.py`: fila assincrona com workers Playwright
- `bot_faturas_v2/connectors/`: conectores isolados por operadora/plataforma
- `public/faturas.html`: nova area operacional com upload, processamento, resultados, detalhe e historico
- `public/faturas.js`: frontend do novo fluxo FastAPI
- `scripts/tiflux_historico.py`: automacao Playwright para atualizar `Historico da Fatura` em tickets do TiFlux

Uso pelo painel web:

- acesse `http://localhost:3210/faturas.html` ou clique em `Abrir bot de faturas` na home
- inicie a API FastAPI do bot em paralelo
- envie uma planilha CSV/XLSX
- acompanhe o lote em tempo real nas abas de processamento, resultado, detalhe e historico

### Subida da API FastAPI do BOT DE FATURAS

Com o ambiente virtual ativo, rode:

```powershell
.\.venv\Scripts\python.exe .\scripts\run_bot_faturas_api.py
```

Por padrao a API responde em:

```text
http://127.0.0.1:8321
```

O painel web em `http://localhost:3210/faturas.html` tenta se conectar automaticamente a essa API.

## Automacao TiFlux de Historico

Script focado no fluxo mostrado na gravacao:

- abrir o ticket
- entrar em `Area de Faturas`
- clicar em editar
- preencher `Historico da Fatura`
- salvar

Primeiro uso recomendado:

```powershell
.\.venv\Scripts\python.exe .\scripts\tiflux_historico.py 550910 "Texto do historico" --email "carlos.siqueira@y3gestao.com.br" --show-browser
```

Nos proximos usos, a sessao salva em `storage/tiflux/session.json` tende a evitar novo login:

```powershell
.\.venv\Scripts\python.exe .\scripts\tiflux_historico.py 550910 "Texto do historico"
```

Observacoes:

- se o TiFlux pedir codigo por e-mail, conclua a autenticacao na janela aberta e pressione `ENTER` no terminal
- se quiser rodar em background, voce pode passar `--headless`
- se a URL do seu ambiente mudar de `entities_3622`, use `--entity-path`
- screenshots de evidencias ficam em `output/tiflux/`

## Planilha Google para TiFlux

Agora o projeto tambem suporta um fluxo de preenchimento do TiFlux a partir de uma Google Sheet, pensado para outras pessoas usarem sem terminal.

Campos suportados na planilha:

- `NUMERO_TICKET`
- `Historico da Fatura`
- `Impedimento`
- `Tratativas/observacoes`
- `Fatura Assumida (Data)`
- `BO+DT (Data)`
- `RPS+NF (Data)`
- `NF Prefeitura`
- `AE (Data)`
- `Importacao (Data)`
- `Envio (Data)`
- `Concluido (Data)`

Colunas de retorno preenchidas automaticamente pela automacao:

- `STATUS_EXECUCAO`
- `MENSAGEM_EXECUCAO`
- `PROCESSADO_EM`
- `CAMPOS_APLICADOS`
- `EVIDENCIA`

Como funciona:

1. a pessoa preenche a aba com o numero do ticket na primeira coluna
2. preenche somente os campos que deseja alterar no TiFlux
3. clica no botao/menu da planilha
4. a API local/publica le a aba, usa a sessao do seu usuario TiFlux e devolve o status na propria planilha

Endpoint FastAPI criado para esse fluxo:

- `POST /v1/tiflux/google-sheet/run`
- `GET /v1/tiflux/google-sheet/jobs/{job_id}`
- `GET /v1/tiflux/google-sheet/template`

Requisitos especificos:

- configurar `TIFLUX_EMAIL` e `TIFLUX_PASSWORD` no ambiente da API
- manter a credencial Google disponivel por arquivo ou por variavel `GOOGLE_SERVICE_ACCOUNT_JSON`
- compartilhar a Google Sheet com o e-mail da service account
- manter a API FastAPI rodando
- se o acesso for pela planilha Google, expor a API por um endereco acessivel externamente
  - exemplo: URL publica do painel + proxy `/bot-faturas-api`

Planilha operacional atual:

- arquivo: `Automacao`
- aba unica: `TIFLUX_PREENCHIMENTO`
- checkbox-botao: `S2`
- status do disparo: `T2`
- ultimo job: `U2`

Fluxo da aba:

1. preencher `NUMERO_TICKET`
2. preencher apenas os campos desejados
3. deixar o restante em branco
4. marcar a caixa `S2`
5. aguardar o retorno nas colunas de status da propria linha

Observacao sobre selecao no TiFlux:

- `Impedimento` e `Tratativas/observacoes` aceitam dropdown na planilha
- na automacao do TiFlux esses campos sao preenchidos digitando o valor e confirmando com `Enter`

Apps Script para a planilha:

- modelo pronto em [google_apps_script_tiflux.gs](scripts/google_apps_script_tiflux.gs)
- ajuste `TIFLUX_API_BASE` para o seu endereco publico
- no Google Sheets, abra `Extensoes > Apps Script`, cole o conteudo e salve
- atualize a planilha para aparecer o menu `TiFlux Y3`
- se quiser botao visual, insira um desenho e associe a funcao `processarAbaTiflux`

Observacoes:

- o processamento escreve o retorno linha a linha na propria aba
- a automacao tenta reaproveitar a sessao salva em `storage/tiflux/session.json`
- se a sessao expirar, o login tenta reaproveitar o codigo mais recente recebido no Outlook/Notificacoes do Windows
- se a conta TiFlux estiver bloqueada, o erro volta para a planilha na coluna `MENSAGEM_EXECUCAO`

## Como rodar

### Painel web

```powershell
npm.cmd start
```

ou

```powershell
npm.cmd run web
```

Depois abra:

```text
http://localhost:3210
```

### Link publico temporario

Se quiser compartilhar o painel rodando no seu Windows por um link externo:

```powershell
npm.cmd run public-link
```

ou use o atalho:

```powershell
.\link-publico-y3.cmd
```

Esse fluxo usa Quick Tunnel do Cloudflare e gera um link `trycloudflare.com`.

Observacoes:

- o seu computador precisa continuar ligado
- o painel precisa continuar rodando
- o link muda quando o tunel e reiniciado
- isso e indicado para teste e uso rapido, nao para ambiente permanente

### Publicacao estavel no Railway

O projeto agora tambem esta preparado para deploy em um unico servico no Railway, com:

- painel Node.js exposto no dominio publico do Railway
- API FastAPI do `BOT DE FATURAS` rodando no mesmo deploy, atras de proxy interno
- healthcheck composto em `/api/health/full`
- persistencia local em `storage/faturas` preparada para ser montada em volume do Railway

Arquivos de deploy incluidos:

- `Dockerfile`
- `railway.json`
- `scripts/start_railway.sh`

Fluxo recomendado no Railway:

1. criar um novo projeto a partir deste repositório
2. habilitar um volume persistente e montar em `/data`
3. definir as variaveis necessarias em `Variables`
4. publicar o servico
5. abrir a URL publica `*.up.railway.app`

Variaveis importantes para o Railway:

```text
PORT=3210
BOT_FATURAS_API_HOST=127.0.0.1
BOT_FATURAS_API_PORT=8321
BOT_FATURAS_API_BASE=http://127.0.0.1:8321
BOT_FATURAS_STORAGE_DIR=/data/storage/faturas
BOT_FATURAS_DB_PATH=/data/storage/faturas/bot_faturas.db
BOT_FATURAS_ENCRYPTION_KEY=<gerar-uma-chave-fernet-ou-deixar-em-branco>
GOOGLE_SERVICE_ACCOUNT_FILE=/app/google-service-account.json
GOOGLE_SERVICE_ACCOUNT_JSON=<cole-o-json-completo-da-service-account-em-uma-linha>
TIFLUX_EMAIL=carlos.siqueira@y3gestao.com.br
TIFLUX_PASSWORD=<senha-do-tiflux>
TIFLUX_BASE_URL=https://app.tiflux.com
TIFLUX_ENTITY_PATH=entities_3622
TIFLUX_SHEET_TITLE=Automação
TIFLUX_WORKSHEET_TITLE=TIFLUX_PREENCHIMENTO
```

Observacoes:

- a URL do Railway e estavel enquanto o servico existir, diferente do `trycloudflare.com`
- o volume e importante para preservar planilhas, ZIPs, evidencias e banco SQLite entre deploys
- se quiser, depois voce ainda pode conectar um dominio proprio por cima
- a URL que o Apps Script deve chamar fica assim:
  - `https://SEU-PROJETO.up.railway.app/bot-faturas-api`

### CLI

```powershell
npm.cmd run cli -- --eml "C:\caminho\email.eml"
```

ou

```powershell
npm.cmd run cli -- --subject "ASSUNTO DO EMAIL"
```

## Scripts uteis

- `npm start`: sobe o painel web
- `npm run web`: sobe o painel web
- `npm run cli -- ...`: roda o processamento por linha de comando
- `npm test`: roda os testes
- `npm run check`: valida sintaxe dos arquivos principais e roda os testes

## Saidas geradas

Cada execucao cria uma pasta em `output/` com timestamp, contendo por exemplo:

- `planilha_padronizada_*.xlsx`
- `arquivo_geral_*.xlsx`
- `pendencias_remanescentes_*.xlsx`
- `resumo_email.txt` ou `resumo_comparativo.txt`
- `anexos/`

## Outlook

O envio automatico de follow-up depende do Outlook desktop do Windows.

Se o Outlook nao estiver instalado ou registrado corretamente na maquina, o sistema mostra um erro amigavel no painel em vez de falhar silenciosamente.

## Observacoes importantes

- o projeto foi pensado para operacao local
- arquivos reais de e-mail e saidas geradas nao devem ser versionados
- a automacao tenta preencher o maximo possivel a partir de planilha, links e PDFs, mas ainda pode emitir avisos quando algum dado vier ausente ou inconsistente na origem

## Checklist de release

Antes de subir para o GitHub, rode:

```powershell
npm.cmd run check
```

e confirme:

- painel abrindo em `http://localhost:3210`
- padronizacao processando um `.eml` real
- comparativo processando duas entradas reais
- hyperlinks da coluna `NFS-e` funcionando no Excel
- Outlook desktop instalado, se o follow-up automatico fizer parte do seu uso

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
- [public/app.js](public/app.js): frontend da home
- [public/comparativo.js](public/comparativo.js): frontend do comparativo
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

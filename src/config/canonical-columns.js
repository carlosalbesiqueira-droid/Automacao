const CANONICAL_COLUMNS = [
  { key: 'nfse', title: 'NFS-e', width: 14, type: 'text' },
  { key: 'rps', title: 'RPS', width: 18, type: 'text' },
  { key: 'dps', title: 'RPS', width: 14, type: 'text' },
  { key: 'emissao', title: 'Emissão', width: 18, type: 'dateTime' },
  { key: 'dataFatoGerador', title: 'Data Fato Gerador', width: 18, type: 'date' },
  { key: 'tomador', title: 'Tomador de Serviços', width: 34, type: 'text' },
  { key: 'cnpj', title: 'CNPJ', width: 20, type: 'text' },
  { key: 'intermediario', title: 'Intermediário de Serviços', width: 26, type: 'text' },
  { key: 'valorServico', title: 'Valor Serviços', width: 15, type: 'currency' },
  { key: 'valorDeducao', title: 'Valor Dedução', width: 15, type: 'currency' },
  { key: 'issDevido', title: 'ISS devido', width: 14, type: 'currency' },
  { key: 'issPagar', title: 'ISS a pagar', width: 14, type: 'currency' },
  { key: 'valorCredito', title: 'Valor Crédito', width: 14, type: 'currency' },
  { key: 'issRetido', title: 'ISS', width: 10, type: 'text' },
  { key: 'situacao', title: 'Situação', width: 14, type: 'text' },
  { key: 'issPagoGuia', title: 'ISS pago por guia NFS-e', width: 24, type: 'text' },
  { key: 'cartaDe', title: 'Carta de', width: 18, type: 'text' },
  { key: 'numeroObra', title: 'Nº da obra', width: 18, type: 'text' }
];

const COLUMN_ALIASES = {
  nfse: ['nfse', 'nfs-e', 'nfs e', 'nfs', 'nota fiscal', 'numero da nota', 'numero nota', 'rep'],
  rps: ['rps', 'rps completo', 'serie rps', 'rps serie', 'numero rps'],
  dps: ['dps', 'nf', 'numero nf', 'numero da nf', 'nota'],
  emissao: ['emissao', 'emissao rps', 'data emissao', 'data de emissao', 'data emissao nf', 'data hora emissao'],
  dataFatoGerador: ['data fato gerador', 'fato gerador', 'data servico', 'data rps', 'competencia'],
  tomador: ['tomador de servicos', 'tomador', 'cliente', 'razao social', 'tomador servico'],
  cnpj: ['cnpj', 'cpf/cnpj', 'cpf cnpj', 'cnpj cliente', 'cnpj tomador', 'cnpj arcelor'],
  intermediario: ['intermediario de servicos', 'intermediario', 'atividade', 'atividade servico', 'codigo atividade'],
  valorServico: ['valor servico', 'valor servicos', 'valor servico r$', 'valor servicos r$', 'valor rps', 'valor da nota', 'valor total', 'valor total da nota'],
  valorDeducao: ['valor deducao', 'valor deducoes', 'deducoes', 'valor deducoes r$'],
  issDevido: ['iss devido', 'valor iss', 'valor do iss', 'iss nf'],
  issPagar: ['iss a pagar', 'iss pagar'],
  valorCredito: ['valor credito', 'credito', 'credito nota salvador'],
  issRetido: ['iss', 'iss retido', 'retencao iss'],
  situacao: ['situacao', 'status'],
  issPagoGuia: ['iss pago por guia nfs-e', 'iss pago guia', 'guia paga', 'iss recolhido'],
  cartaDe: ['carta de', 'carta de correcao', 'carta de correção'],
  numeroObra: ['nº da obra', 'no da obra', 'numero da obra', 'n da obra', 'inscricao da obra', 'inscrição da obra']
};

module.exports = {
  CANONICAL_COLUMNS,
  COLUMN_ALIASES
};

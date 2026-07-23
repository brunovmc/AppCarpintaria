const ABA_INVESTIMENTOS = 'INVESTIMENTOS';
const INVESTIMENTOS_CACHE_SCOPE = 'INVESTIMENTOS_LISTA_ATIVOS';
const INVESTIMENTOS_CACHE_TTL_SEC = 90;

const INVESTIMENTOS_SCHEMA = [
  'ID',
  'descricao',
  'tipo_investimento',
  'investidor',
  'referencia_investimento',
  'valor_total_investimento',
  'recebido_por',
  'data_investimento',
  'forma_pagamento',
  'parcelas',
  'parcelas_detalhe_json',
  'client_request_id',
  'observacao',
  'ativo',
  'criado_em'
];

function normalizarClientRequestIdInvestimento_(valor) {
  return String(valor || '').trim().replace(/[^a-zA-Z0-9:_-]+/g, '').slice(0, 120);
}

function validarValorListaInvestimento_(valor, lista, nomeCampo, obrigatorio) {
  const bruto = String(valor || '').trim();
  if (!bruto) {
    if (obrigatorio) throw new Error(`${nomeCampo} e obrigatorio.`);
    return '';
  }
  const opcoes = Array.isArray(lista) ? lista.map(item => String(item || '').trim()).filter(Boolean) : [];
  if (opcoes.length === 0) return bruto;
  const encontrado = opcoes.find(item => item.toUpperCase() === bruto.toUpperCase());
  if (!encontrado) throw new Error(`Selecione uma opcao valida em ${nomeCampo}.`);
  return encontrado;
}

function normalizarPayloadInvestimento(payload) {
  const dados = payload || {};
  const descricao = String(dados.descricao || '').trim().slice(0, 200);
  if (!descricao) throw new Error('Descricao do investimento e obrigatoria.');

  const valorTotal = round2Financeiro(parseNumeroBR(dados.valor_total_investimento));
  if (valorTotal <= 0) throw new Error('Valor total do investimento deve ser maior que zero.');

  const dataInvestimento = normalizarDataFinanceiro(
    dados.data_investimento,
    true,
    'Data do investimento'
  );
  const validacoes = typeof obterValidacoes === 'function' ? obterValidacoes() : {};
  const tipoInvestimento = validarValorListaInvestimento_(
    dados.tipo_investimento || 'APORTE_SOCIO',
    validacoes?.tiposInvestimento,
    'Tipo de investimento',
    true
  );
  const investidor = validarValorListaInvestimento_(
    dados.investidor,
    validacoes?.investidores,
    'Investidor',
    true
  );
  const recebidoPor = validarPagoPorFinanceiro(dados.recebido_por, false);
  const formaPagamento = validarFormaPagamentoFinanceiro(dados.forma_pagamento, false);
  const parcelas = normalizarParcelasFinanceiro(dados.parcelas, formaPagamento);
  const parcelasDetalhe = normalizarParcelasDetalhePayloadFinanceiro(
    dados.parcelas_detalhe ?? dados.parcelas_detalhe_json,
    parcelas,
    dataInvestimento || new Date(),
    valorTotal
  );

  return {
    descricao,
    tipo_investimento: tipoInvestimento,
    investidor,
    referencia_investimento: String(dados.referencia_investimento || '').trim().slice(0, 160),
    valor_total_investimento: valorTotal,
    recebido_por: recebidoPor,
    data_investimento: dataInvestimento,
    forma_pagamento: formaPagamento,
    parcelas,
    parcelas_detalhe_json: serializarParcelasDetalheFinanceiro(parcelasDetalhe),
    observacao: String(dados.observacao || '').trim().slice(0, 500)
  };
}

function formatarDataInvestimentoSeguro_(valor, formato) {
  const data = parseDataFinanceiro(valor);
  return data ? Utilities.formatDate(data, Session.getScriptTimeZone(), formato) : '';
}

function lerCacheInvestimentos_() {
  return appCacheGetJson(INVESTIMENTOS_CACHE_SCOPE);
}

function salvarCacheInvestimentos_(lista) {
  return appCachePutJson(
    INVESTIMENTOS_CACHE_SCOPE,
    Array.isArray(lista) ? lista : [],
    INVESTIMENTOS_CACHE_TTL_SEC
  );
}

function limparCacheInvestimentos() {
  return appCacheRemove(INVESTIMENTOS_CACHE_SCOPE);
}

function recarregarCacheInvestimentos() {
  limparCacheInvestimentos();
  const dados = listarInvestimentos(true);
  return {
    ok: true,
    scope: INVESTIMENTOS_CACHE_SCOPE,
    ttl_segundos: INVESTIMENTOS_CACHE_TTL_SEC,
    total_itens: Array.isArray(dados) ? dados.length : 0
  };
}

function listarInvestimentos(forcarRecarregar, ambiente) {
  if (typeof executarComAmbienteBancoDadosAutorizado_ !== 'function') {
    throw new Error('Controle de ambiente indisponivel.');
  }
  return executarComAmbienteBancoDadosAutorizado_(
    ambiente,
    () => listarInvestimentosNoAmbienteAtual_(forcarRecarregar)
  );
}

function listarInvestimentosNoAmbienteAtual_(forcarRecarregar) {
  if (!forcarRecarregar) {
    const cached = lerCacheInvestimentos_();
    if (Array.isArray(cached)) return cached;
  }
  const sheet = getSheet(ABA_INVESTIMENTOS);
  if (!sheet) {
    salvarCacheInvestimentos_([]);
    return [];
  }
  const base = rowsToObjects(sheet)
    .filter(item => String(item.ativo).toLowerCase() === 'true')
    .map(item => ({
      ...item,
      valor_total_investimento: round2Financeiro(parseNumeroBR(item.valor_total_investimento)),
      parcelas: Math.max(1, Math.floor(parseNumeroBR(item.parcelas) || 1)),
      client_request_id: String(item.client_request_id || '').trim(),
      data_investimento: formatarDataInvestimentoSeguro_(item.data_investimento, 'yyyy-MM-dd'),
      criado_em: formatarDataInvestimentoSeguro_(item.criado_em, 'yyyy-MM-dd HH:mm')
    }));
  const lista = (typeof enriquecerInvestimentosComResumoPagamento === 'function'
    ? enriquecerInvestimentosComResumoPagamento(base)
    : base
  ).sort((a, b) => {
    const da = parseDataFinanceiro(a.data_investimento)?.getTime() || 0;
    const db = parseDataFinanceiro(b.data_investimento)?.getTime() || 0;
    if (da !== db) return db - da;
    return (parseDataFinanceiro(b.criado_em)?.getTime() || 0) -
      (parseDataFinanceiro(a.criado_em)?.getTime() || 0);
  });
  salvarCacheInvestimentos_(lista);
  return lista;
}

function buscarInvestimentoPorClientRequestId_(sheet, clientRequestId) {
  const requestId = normalizarClientRequestIdInvestimento_(clientRequestId);
  if (!sheet || !requestId) return null;
  return rowsToObjects(sheet).find(item =>
    String(item.client_request_id || '').trim() === requestId &&
    String(item.ativo).toLowerCase() === 'true'
  ) || null;
}

function criarInvestimento(payload, opcoesInternas) {
  assertCanWrite('Criacao de investimento');
  const dados = normalizarPayloadInvestimento(payload);
  const clientRequestId = normalizarClientRequestIdInvestimento_(payload?.client_request_id);
  const lock = opcoesInternas?.lockJaAdquirido === true ? null : LockService.getScriptLock();
  if (lock) lock.waitLock(10000);
  try {
    const ss = getDataSpreadsheet();
    let sheet = ss.getSheetByName(ABA_INVESTIMENTOS);
    if (!sheet) sheet = ss.insertSheet(ABA_INVESTIMENTOS);
    ensureSchema(sheet, INVESTIMENTOS_SCHEMA);
    const existente = buscarInvestimentoPorClientRequestId_(sheet, clientRequestId);
    if (existente) {
      return listarInvestimentosNoAmbienteAtual_(true).find(item => item.ID === existente.ID) || existente;
    }
    const novo = {
      ...dados,
      ID: gerarId('INV'),
      client_request_id: clientRequestId,
      ativo: true,
      criado_em: new Date()
    };
    if (!insert(ABA_INVESTIMENTOS, novo, INVESTIMENTOS_SCHEMA)) return null;
    gerarParcelasFinanceirasOrigem(ORIGEM_TIPO_INVESTIMENTO, novo.ID, {
      item: novo,
      total_previsto: novo.valor_total_investimento,
      natureza: NATUREZA_RECEBIMENTO
    });
    return listarInvestimentosNoAmbienteAtual_(true).find(item => item.ID === novo.ID) || null;
  } finally {
    if (lock) lock.releaseLock();
  }
}

function atualizarInvestimento(id, payload) {
  assertCanWrite('Atualizacao de investimento');
  const investimentoId = String(id || '').trim();
  if (!investimentoId) throw new Error('ID do investimento nao informado.');
  const sheet = getSheet(ABA_INVESTIMENTOS);
  if (!sheet) throw new Error('Aba INVESTIMENTOS nao encontrada.');
  const atual = rowsToObjects(sheet).find(item => String(item.ID || '').trim() === investimentoId);
  if (!atual || String(atual.ativo).toLowerCase() !== 'true') {
    throw new Error('Investimento nao encontrado ou inativo.');
  }
  const dados = normalizarPayloadInvestimento(payload);
  const ok = updateById(ABA_INVESTIMENTOS, 'ID', investimentoId, dados, INVESTIMENTOS_SCHEMA);
  if (ok) regerarParcelasFinanceirasOrigemComPagamentos(ORIGEM_TIPO_INVESTIMENTO, investimentoId);
  return ok;
}

function deletarInvestimento(id) {
  assertCanWrite('Exclusao de investimento');
  const investimentoId = String(id || '').trim();
  const sheet = getSheet(ABA_INVESTIMENTOS);
  if (!sheet) return false;
  const atual = rowsToObjects(sheet).find(item => String(item.ID || '').trim() === investimentoId);
  if (!atual || String(atual.ativo).toLowerCase() !== 'true') return false;
  const recebido = calcularTotalPagoOrigemFinanceiro(ORIGEM_TIPO_INVESTIMENTO, investimentoId, true);
  if (recebido > 0.009) throw new Error('Desfaca os recebimentos antes de excluir o investimento.');
  const ok = updateById(
    ABA_INVESTIMENTOS,
    'ID',
    investimentoId,
    { ativo: false },
    INVESTIMENTOS_SCHEMA
  );
  if (ok) limparParcelasFinanceirasOrigem(ORIGEM_TIPO_INVESTIMENTO, investimentoId);
  return ok;
}

function registrarRecebimentoInvestimento(investimentoId, payload) {
  assertCanWrite('Registro de recebimento de investimento');
  const id = String(investimentoId || '').trim();
  if (!id) throw new Error('Investimento invalido.');
  const pagamentoCriado = registrarPagamento(ORIGEM_TIPO_INVESTIMENTO, id, payload);
  return {
    pagamentoCriado,
    investimentoAtualizado: listarInvestimentosNoAmbienteAtual_(true).find(item => item.ID === id) || null
  };
}

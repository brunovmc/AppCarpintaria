const ABA_ESTOQUE = 'ESTOQUE';
const ESTOQUE_CACHE_SCOPE = 'ESTOQUE_LISTA_ATIVOS';
const ESTOQUE_CACHE_TTL_SEC = 90;

const ESTOQUE_SCHEMA = [
  'ID',
  'tipo',
  'item',
  'unidade',
  'valor_unit',
  'custo_unitario',
  'preco_venda',
  'ativo',
  'criado_em',
  'quantidade',
  'comprimento_cm',
  'largura_cm',
  'espessura_cm',
  'categoria',
  'fornecedor',
  'pago_por',
  'potencia',
  'voltagem',
  'comprado_em',
  'data_pagamento',
  'forma_pagamento',
  'parcelas',
  'vida_util_mes',
  'observacao',
  'origem_tipo',
  'origem_id',
  'op_id'
];

function validarPagoPorEstoque(pagoPor) {
  const valor = String(pagoPor || '').trim();
  if (!valor) return '';

  const validacoes = obterValidacoes();
  const lista = Array.isArray(validacoes?.pagosPor) ? validacoes.pagosPor : [];
  if (lista.length === 0) return valor;

  const match = lista.find(v => String(v || '').trim().toUpperCase() === valor.toUpperCase());
  if (!match) {
    throw new Error('Pago por invalido.');
  }
  return match;
}

function categoriaValidaParaTipoEstoque(tipo, categoria) {
  const t = String(tipo || '').trim().toUpperCase();
  const c = String(categoria || '').trim().toUpperCase();
  if (!t || !c) return false;

  const validacoes = obterValidacoes();
  const mapa = validacoes?.categoriasPorTipo || {};
  const categoriasTipo = Array.isArray(mapa[t]) ? mapa[t] : [];

  if (categoriasTipo.length > 0) {
    return categoriasTipo.some(v => String(v || '').trim().toUpperCase() === c);
  }

  const categoriasGerais = Array.isArray(validacoes?.categorias) ? validacoes.categorias : [];
  return categoriasGerais.some(v => String(v || '').trim().toUpperCase() === c);
}

function normalizarPayloadMadeiraEstoque(payload) {
  const dados = { ...(payload || {}) };
  const tipo = String(dados.tipo || '').trim().toUpperCase();
  const custoBase = (dados.custo_unitario !== undefined && dados.custo_unitario !== null && String(dados.custo_unitario).trim() !== '')
    ? dados.custo_unitario
    : dados.valor_unit;
  const custoUnitario = parseNumeroBR(custoBase);
  if (custoUnitario < 0) {
    throw new Error('Custo unitario nao pode ser negativo.');
  }
  const precoBruto = String(dados.preco_venda ?? '').trim();
  const precoVenda = precoBruto === '' ? '' : parseNumeroBR(precoBruto);
  if (precoVenda !== '' && precoVenda < 0) {
    throw new Error('Preco de venda nao pode ser negativo.');
  }

  dados.custo_unitario = custoUnitario;
  dados.preco_venda = precoVenda;
  // Compatibilidade com partes do app que ainda usam valor_unit.
  dados.valor_unit = custoUnitario;
  dados.origem_tipo = String(dados.origem_tipo || '').trim();
  dados.origem_id = String(dados.origem_id || '').trim();
  dados.op_id = String(dados.op_id || '').trim();
  dados.pago_por = validarPagoPorEstoque(dados.pago_por);
  dados.data_pagamento = normalizarDataFinanceiro(dados.data_pagamento, false, 'Data de pagamento');
  const formaPagamento = validarFormaPagamentoFinanceiro(dados.forma_pagamento, false);
  dados.forma_pagamento = formaPagamento;
  dados.parcelas = normalizarParcelasFinanceiro(dados.parcelas, formaPagamento);

  if (!categoriaValidaParaTipoEstoque(tipo, dados.categoria)) {
    throw new Error('Categoria invalida para o tipo selecionado.');
  }

  if (tipo !== 'MADEIRA') return dados;

  const comprimento = parseNumeroBR(dados.comprimento_cm);
  const largura = parseNumeroBR(dados.largura_cm);
  const espessura = parseNumeroBR(dados.espessura_cm);

  if (comprimento <= 0 || largura <= 0 || espessura <= 0) {
    throw new Error('Para MADEIRA, informe comprimento, largura e espessura validos.');
  }

  dados.comprimento_cm = comprimento;
  dados.largura_cm = largura;
  dados.espessura_cm = espessura;
  dados.quantidade = Number(((comprimento * largura * espessura) / 1000000).toFixed(2));
  dados.unidade = 'M3';
  return dados;
}

function formatarDataEstoqueSeguro(valor, formato) {
  if (!valor) return '';
  let d = null;
  if (typeof parseDataFinanceiro === 'function') {
    d = parseDataFinanceiro(valor);
  }
  if (!d) {
    const fallback = new Date(valor);
    d = isNaN(fallback.getTime()) ? null : fallback;
  }
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), formato);
}


function lerCacheListaEstoque() {
  return appCacheGetJson(ESTOQUE_CACHE_SCOPE);
}

function salvarCacheListaEstoque(lista) {
  appCachePutJson(ESTOQUE_CACHE_SCOPE, Array.isArray(lista) ? lista : [], ESTOQUE_CACHE_TTL_SEC);
}

function limparCacheEstoque() {
  return appCacheRemove(ESTOQUE_CACHE_SCOPE);
}

function recarregarCacheEstoque() {
  limparCacheEstoque();
  const dados = listarEstoque(true);
  return {
    ok: true,
    scope: ESTOQUE_CACHE_SCOPE,
    ttl_segundos: ESTOQUE_CACHE_TTL_SEC,
    total_itens: Array.isArray(dados) ? dados.length : 0
  };
}

function listarEstoque(forcarRecarregar) {
  if (!forcarRecarregar) {
    const cached = lerCacheListaEstoque();
    if (Array.isArray(cached)) {
      return cached;
    }
  }

  const sheet = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
  if (!sheet) {
    salvarCacheListaEstoque([]);
    return [];
  }

  const rows = rowsToObjects(sheet);

  const lista = rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .map(i => ({
      ...i,
      custo_unitario: parseNumeroBR(i.custo_unitario || i.valor_unit),
      preco_venda: String(i.preco_venda || '').trim() === '' ? '' : parseNumeroBR(i.preco_venda),
      valor_unit: parseNumeroBR(i.custo_unitario || i.valor_unit),
      criado_em: formatarDataEstoqueSeguro(i.criado_em, 'yyyy-MM-dd HH:mm'),
      comprado_em: formatarDataEstoqueSeguro(i.comprado_em, 'yyyy-MM-dd'),
      data_pagamento: formatarDataEstoqueSeguro(i.data_pagamento, 'yyyy-MM-dd')
    }));
  salvarCacheListaEstoque(lista);
  return lista;
}

function testeListarEstoqueDireto() {
  const itens = listarEstoque();
  Logger.log(itens);
}

function testeDebugEstoque() {
  const sheet = getSheet('ESTOQUE');
  const todos = rowsToObjects(sheet);
  Logger.log('TODOS:');
  Logger.log(todos);

  const filtrados = listarEstoque();
  Logger.log('FILTRADOS:');
  Logger.log(filtrados);
}



function criarItemEstoque(payload) {
  const dados = normalizarPayloadMadeiraEstoque(payload);
  const novo = {
    ...dados,
    ID: gerarId('EST'),
    ativo: true,
    criado_em: new Date()
  };

  return insert(ABA_ESTOQUE, novo, ESTOQUE_SCHEMA);
}


function testeCriarItem(){
  criarItemEstoque({
  tipo: 'MADEIRA',
  item: 'TESTE',
  unidade: 'M3',
  valor_unit: 9999
});
}

function obterItemEstoque(id) {
  const sheet = getSheet(ABA_ESTOQUE);
  if (!sheet) return null;

  const item = rowsToObjects(sheet).find(i => i.ID === id);
  if (!item) return null;

  return {
    ID: item.ID,
    tipo: item.tipo || '',
    item: item.item || '',
    categoria: item.categoria || '',
    unidade: item.unidade || '',
    quantidade: item.quantidade || '',
    comprimento_cm: item.comprimento_cm || '',
    largura_cm: item.largura_cm || '',
    espessura_cm: item.espessura_cm || '',
    valor_unit: parseNumeroBR(item.custo_unitario || item.valor_unit),
    custo_unitario: parseNumeroBR(item.custo_unitario || item.valor_unit),
    preco_venda: String(item.preco_venda || '').trim() === '' ? '' : parseNumeroBR(item.preco_venda),
    fornecedor: item.fornecedor || '',
    pago_por: item.pago_por || '',
    forma_pagamento: item.forma_pagamento || '',
    parcelas: item.parcelas || '',
    observacao: item.observacao || '',
    origem_tipo: item.origem_tipo || '',
    origem_id: item.origem_id || '',
    op_id: item.op_id || '',

    potencia: item.potencia || '',
    voltagem: item.voltagem || '',
    vida_util_mes: item.vida_util_mes || '',

    criado_em: item.criado_em
      ? formatarDataEstoqueSeguro(item.criado_em, 'yyyy-MM-dd HH:mm')
      : '',

    comprado_em: item.comprado_em
      ? formatarDataEstoqueSeguro(item.comprado_em, 'yyyy-MM-dd')
      : '',
    data_pagamento: item.data_pagamento
      ? formatarDataEstoqueSeguro(item.data_pagamento, 'yyyy-MM-dd')
      : ''
  };
}



function atualizarItemEstoque(id, payload) {
  const dados = normalizarPayloadMadeiraEstoque(payload);
  return updateById(
    ABA_ESTOQUE,
    'ID',
    id,
    dados,
    ESTOQUE_SCHEMA
  );
}



function deletarItemEstoque(id) {
  return updateById(
    ABA_ESTOQUE,
    'ID',
    id,
    { ativo: false },
    ESTOQUE_SCHEMA
  );
}

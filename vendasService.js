const ABA_VENDAS = 'VENDAS';
const VENDAS_CACHE_SCOPE = 'VENDAS_LISTA_ATIVAS';
const VENDAS_CACHE_TTL_SEC = 90;

const VENDAS_SCHEMA = [
  'ID',
  'estoque_id',
  'tipo',
  'categoria',
  'item',
  'unidade',
  'quantidade',
  'valor_unit_venda',
  'valor_total_venda',
  'recebido_por',
  'data_venda',
  'forma_pagamento',
  'parcelas',
  'plano_recebimento',
  'data_entrega_prevista',
  'observacao',
  'estoque_baixado',
  'data_baixa_estoque',
  'ativo',
  'criado_em'
];

function ehItemVendavelEstoque(item) {
  const tipo = String(item?.tipo || '').trim().toUpperCase();
  const categoria = String(item?.categoria || '').trim().toUpperCase();
  return String(item?.ativo).toLowerCase() === 'true' &&
    tipo === 'PRODUTO' &&
    categoria === 'FINALIZADO';
}

function validarRecebidoPorVenda(recebidoPor) {
  return validarPagoPorFinanceiro(recebidoPor, false);
}

function normalizarPlanoRecebimentoVenda(plano) {
  const p = String(plano || '').trim().toUpperCase();
  if (!p) return 'PADRAO';
  if (p === 'PADRAO' || p === 'ENCOMENDA_50_50') {
    return p;
  }
  throw new Error('Plano de recebimento invalido.');
}

function obterItemEstoqueVendavelPorId(estoqueId) {
  const sheet = getSheet(ABA_ESTOQUE);
  if (!sheet) throw new Error('Aba ESTOQUE nao encontrada.');
  const id = String(estoqueId || '').trim();
  const item = rowsToObjects(sheet).find(i => String(i.ID || '').trim() === id);
  if (!item || !ehItemVendavelEstoque(item)) {
    throw new Error('Item de estoque nao e vendavel. Apenas PRODUTO FINALIZADO pode ser vendido.');
  }
  return item;
}

function normalizarPayloadVenda(payload, vendaExistente) {
  const dados = { ...(payload || {}) };
  const estoqueId = String(dados.estoque_id || '').trim();
  if (!estoqueId) {
    throw new Error('Selecione um item de estoque para vender.');
  }

  const estoqueItem = obterItemEstoqueVendavelPorId(estoqueId);
  const quantidade = round2Financeiro(parseNumeroBR(dados.quantidade));
  if (quantidade <= 0) {
    throw new Error('Quantidade da venda deve ser maior que zero.');
  }

  const saldoAtual = round2Financeiro(parseNumeroBR(estoqueItem.quantidade));
  const quantidadeAnterior = vendaExistente
    ? round2Financeiro(parseNumeroBR(vendaExistente.quantidade))
    : 0;
  const estoqueMesmoItem = vendaExistente && String(vendaExistente.estoque_id || '').trim() === estoqueId;
  const estoqueJaBaixado = vendaExistente && String(vendaExistente.estoque_baixado).toLowerCase() === 'true';
  const saldoDisponivel = (vendaExistente && !estoqueJaBaixado && estoqueMesmoItem)
    ? round2Financeiro(saldoAtual + quantidadeAnterior)
    : saldoAtual;

  if (quantidade > saldoDisponivel + 0.009) {
    throw new Error('Quantidade de venda maior que o saldo disponivel no estoque.');
  }

  const dataVenda = normalizarDataFinanceiro(dados.data_venda, true, 'Data de venda');
  const formaPagamento = validarFormaPagamentoFinanceiro(dados.forma_pagamento, false);
  const planoRecebimento = normalizarPlanoRecebimentoVenda(dados.plano_recebimento);
  const parcelasBase = normalizarParcelasFinanceiro(dados.parcelas, formaPagamento);
  const parcelas = planoRecebimento === 'ENCOMENDA_50_50' ? 2 : parcelasBase;
  const dataEntregaPrevista = normalizarDataFinanceiro(
    dados.data_entrega_prevista,
    false,
    'Data de entrega prevista'
  );
  let valorTotal = round2Financeiro(parseNumeroBR(dados.valor_total_venda));
  if (valorTotal <= 0) {
    const precoBase = parseNumeroBR(estoqueItem.preco_venda || 0);
    valorTotal = round2Financeiro(Math.max(0, precoBase * quantidade));
  }
  if (valorTotal <= 0) {
    throw new Error('Valor total da venda deve ser maior que zero.');
  }

  const valorUnit = round2Financeiro(valorTotal / quantidade);

  return {
    estoque_id: estoqueId,
    tipo: 'PRODUTO',
    categoria: 'FINALIZADO',
    item: String(estoqueItem.item || '').trim(),
    unidade: String(estoqueItem.unidade || 'UN').trim() || 'UN',
    quantidade,
    valor_unit_venda: valorUnit,
    valor_total_venda: valorTotal,
    recebido_por: validarRecebidoPorVenda(dados.recebido_por),
    data_venda: dataVenda,
    forma_pagamento: formaPagamento,
    parcelas,
    plano_recebimento: planoRecebimento,
    data_entrega_prevista: dataEntregaPrevista,
    observacao: String(dados.observacao || '').trim()
  };
}

function formatarDataVendaSeguro(valor, formato) {
  if (!valor) return '';
  const d = parseDataFinanceiro(valor);
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), formato);
}

function lerCacheListaVendas() {
  return appCacheGetJson(VENDAS_CACHE_SCOPE);
}

function salvarCacheListaVendas(lista) {
  appCachePutJson(VENDAS_CACHE_SCOPE, Array.isArray(lista) ? lista : [], VENDAS_CACHE_TTL_SEC);
}

function limparCacheVendas() {
  return appCacheRemove(VENDAS_CACHE_SCOPE);
}

function recarregarCacheVendas() {
  limparCacheVendas();
  const dados = listarVendas(true);
  return {
    ok: true,
    scope: VENDAS_CACHE_SCOPE,
    ttl_segundos: VENDAS_CACHE_TTL_SEC,
    total_itens: Array.isArray(dados) ? dados.length : 0
  };
}

function listarVendas(forcarRecarregar) {
  if (!forcarRecarregar) {
    const cached = lerCacheListaVendas();
    if (Array.isArray(cached)) {
      return cached;
    }
  }

  const sheet = getSheet(ABA_VENDAS);
  if (!sheet) {
    salvarCacheListaVendas([]);
    return [];
  }

  const rows = rowsToObjects(sheet);
  const base = rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .map(i => ({
      ...i,
      quantidade: round2Financeiro(parseNumeroBR(i.quantidade)),
      valor_unit_venda: round2Financeiro(parseNumeroBR(i.valor_unit_venda)),
      valor_total_venda: round2Financeiro(parseNumeroBR(i.valor_total_venda)),
      parcelas: Number(i.parcelas || 1),
      estoque_baixado: String(i.estoque_baixado).toLowerCase() === 'true',
      criado_em: formatarDataVendaSeguro(i.criado_em, 'yyyy-MM-dd HH:mm'),
      data_venda: formatarDataVendaSeguro(i.data_venda, 'yyyy-MM-dd'),
      data_entrega_prevista: formatarDataVendaSeguro(i.data_entrega_prevista, 'yyyy-MM-dd'),
      data_baixa_estoque: formatarDataVendaSeguro(i.data_baixa_estoque, 'yyyy-MM-dd')
    }));

  const enriquecida = (typeof enriquecerVendasComResumoPagamento === 'function')
    ? enriquecerVendasComResumoPagamento(base)
    : base;
  const lista = enriquecida
    .map(v => ({
      ...v,
      parcelas_detalhe: (typeof listarParcelasFinanceirasOrigem === 'function')
        ? listarParcelasFinanceirasOrigem(ORIGEM_TIPO_VENDA, v.ID)
        : []
    }))
    .sort((a, b) => {
      const da = parseDataFinanceiro(a.data_venda)?.getTime() || 0;
      const db = parseDataFinanceiro(b.data_venda)?.getTime() || 0;
      if (da !== db) return db - da;
      const ca = parseDataFinanceiro(a.criado_em)?.getTime() || 0;
      const cb = parseDataFinanceiro(b.criado_em)?.getTime() || 0;
      return cb - ca;
    });

  salvarCacheListaVendas(lista);
  return lista;
}

function listarItensEstoqueVendaveis() {
  const sheet = getSheet(ABA_ESTOQUE);
  if (!sheet) return [];
  return rowsToObjects(sheet)
    .filter(i => ehItemVendavelEstoque(i))
    .map(i => ({
      ID: i.ID,
      item: i.item || '',
      unidade: i.unidade || 'UN',
      categoria: i.categoria || '',
      quantidade: round2Financeiro(parseNumeroBR(i.quantidade)),
      preco_venda: round2Financeiro(parseNumeroBR(i.preco_venda || 0))
    }))
    .filter(i => i.quantidade > 0)
    .sort((a, b) => a.item.localeCompare(b.item));
}

function criarVenda(payload) {
  const dados = normalizarPayloadVenda(payload);
  const novo = {
    ...dados,
    ID: gerarId('VND'),
    estoque_baixado: false,
    data_baixa_estoque: '',
    ativo: true,
    criado_em: new Date()
  };

  const ok = insert(ABA_VENDAS, novo, VENDAS_SCHEMA);
  if (!ok) return null;
  gerarParcelasFinanceirasOrigem(ORIGEM_TIPO_VENDA, novo.ID);
  return listarVendas(true).find(i => i.ID === novo.ID) || null;
}

function atualizarVenda(id, payload) {
  const vendaId = String(id || '').trim();
  if (!vendaId) {
    throw new Error('ID da venda nao informado.');
  }

  const sheet = getSheet(ABA_VENDAS);
  if (!sheet) {
    throw new Error('Aba VENDAS nao encontrada.');
  }
  const atual = rowsToObjects(sheet).find(i => String(i.ID || '').trim() === vendaId);
  if (!atual || String(atual.ativo).toLowerCase() !== 'true') {
    throw new Error('Venda nao encontrada ou inativa.');
  }

  const estoqueJaBaixado = String(atual.estoque_baixado).toLowerCase() === 'true';
  if (estoqueJaBaixado) {
    const novoEstoqueId = String(payload?.estoque_id || atual.estoque_id || '').trim();
    const novaQtd = round2Financeiro(parseNumeroBR(payload?.quantidade || atual.quantidade || 0));
    const estoqueAtualId = String(atual.estoque_id || '').trim();
    const qtdAtual = round2Financeiro(parseNumeroBR(atual.quantidade));
    if (novoEstoqueId !== estoqueAtualId || Math.abs(novaQtd - qtdAtual) > 0.009) {
      throw new Error('Nao e permitido alterar item/quantidade apos baixa de estoque da venda.');
    }
  }

  const dados = normalizarPayloadVenda(payload, atual);
  const ok = updateById(
    ABA_VENDAS,
    'ID',
    vendaId,
    dados,
    VENDAS_SCHEMA
  );
  if (ok) {
    regerarParcelasFinanceirasOrigemComPagamentos(ORIGEM_TIPO_VENDA, vendaId);
  }
  return ok;
}

function deletarVenda(id) {
  const vendaId = String(id || '').trim();
  const sheet = getSheet(ABA_VENDAS);
  if (!sheet) return false;
  const atual = rowsToObjects(sheet).find(i => String(i.ID || '').trim() === vendaId);
  if (!atual) return false;
  if (String(atual.estoque_baixado).toLowerCase() === 'true') {
    throw new Error('Nao e permitido excluir venda com estoque ja baixado.');
  }

  const ok = updateById(
    ABA_VENDAS,
    'ID',
    vendaId,
    { ativo: false },
    VENDAS_SCHEMA
  );
  if (ok) {
    limparParcelasFinanceirasOrigem(ORIGEM_TIPO_VENDA, vendaId);
  }
  return ok;
}

function aplicarBaixaEstoqueVendaNoPrimeiroRecebimento(vendaId) {
  const id = String(vendaId || '').trim();
  if (!id) throw new Error('Venda invalida.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheetVendas = getSheet(ABA_VENDAS);
    if (!sheetVendas) throw new Error('Aba VENDAS nao encontrada.');
    const venda = rowsToObjects(sheetVendas).find(i => String(i.ID || '').trim() === id);
    if (!venda || String(venda.ativo).toLowerCase() !== 'true') {
      throw new Error('Venda nao encontrada ou inativa.');
    }
    if (String(venda.estoque_baixado).toLowerCase() === 'true') {
      return { ok: true, ja_baixado: true };
    }

    const totalPago = calcularTotalPagoOrigemFinanceiro(ORIGEM_TIPO_VENDA, id, true);
    if (totalPago <= 0.009) {
      return { ok: true, ja_baixado: false };
    }

    const estoqueId = String(venda.estoque_id || '').trim();
    const sheetEstoque = getSheet(ABA_ESTOQUE);
    if (!sheetEstoque) throw new Error('Aba ESTOQUE nao encontrada.');
    const estoqueRows = rowsToObjects(sheetEstoque);
    const item = estoqueRows.find(i => String(i.ID || '').trim() === estoqueId);
    if (!item || String(item.ativo).toLowerCase() !== 'true') {
      throw new Error('Item de estoque da venda nao encontrado ou inativo.');
    }
    if (!ehItemVendavelEstoque(item)) {
      throw new Error('Item de estoque da venda deixou de ser vendavel.');
    }

    const quantidadeVenda = round2Financeiro(parseNumeroBR(venda.quantidade));
    const saldoAtual = round2Financeiro(parseNumeroBR(item.quantidade));
    if (saldoAtual + 0.009 < quantidadeVenda) {
      throw new Error('Saldo insuficiente para baixar estoque da venda.');
    }

    const novoSaldo = round2Financeiro(saldoAtual - quantidadeVenda);
    updateById(
      ABA_ESTOQUE,
      'ID',
      estoqueId,
      { quantidade: novoSaldo },
      ESTOQUE_SCHEMA
    );

    const dataBaixa = new Date();
    updateById(
      ABA_VENDAS,
      'ID',
      id,
      {
        estoque_baixado: true,
        data_baixa_estoque: dataBaixa
      },
      VENDAS_SCHEMA
    );

    return {
      ok: true,
      ja_baixado: false,
      estoqueAtualizado: {
        ID: estoqueId,
        quantidade: novoSaldo
      },
      vendaAtualizada: {
        ID: id,
        estoque_baixado: true,
        data_baixa_estoque: formatarDataVendaSeguro(dataBaixa, 'yyyy-MM-dd')
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function registrarRecebimentoVenda(vendaId, payload) {
  const id = String(vendaId || '').trim();
  if (!id) throw new Error('Venda invalida.');

  const pagamentoCriado = registrarPagamento(ORIGEM_TIPO_VENDA, id, payload);
  const baixa = aplicarBaixaEstoqueVendaNoPrimeiroRecebimento(id);
  const vendaAtualizada = listarVendas(true).find(i => i.ID === id) || null;

  return {
    pagamentoCriado,
    baixaEstoque: baixa,
    vendaAtualizada
  };
}

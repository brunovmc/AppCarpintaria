const DASHBOARD_EXECUTIVO_CACHE_SCOPE = 'DASHBOARD_EXECUTIVO_FINANCEIRO';
const DASHBOARD_EXECUTIVO_CACHE_TTL_SEC = 120;
const DASHBOARD_EXECUTIVO_VERSION = 'v2';

function limparCacheDashboardExecutivoFinanceiro() {
  return appCacheRemove(DASHBOARD_EXECUTIVO_CACHE_SCOPE);
}

function normalizarTextoDashboardExecutivo(valor, fallback) {
  const texto = String(valor || '').trim();
  return texto || String(fallback || '').trim();
}

function normalizarChaveDashboardExecutivo(valor) {
  return normalizarTextoSemAcentoFinanceiro(valor)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function deslocarMesDashboardExecutivo(referenciaYm, delta) {
  const ref = getMesReferenciaFinanceiro(referenciaYm);
  const ano = Number(ref.slice(0, 4));
  const mes = Number(ref.slice(5, 7));
  const data = new Date(ano, mes - 1 + Number(delta || 0), 1);
  return `${data.getFullYear()}-${String(data.getMonth() + 1).padStart(2, '0')}`;
}

function fimDoDiaDashboardExecutivo(data) {
  const inicio = inicioDoDiaFinanceiro(data);
  if (!inicio) return null;
  return new Date(inicio.getFullYear(), inicio.getMonth(), inicio.getDate(), 23, 59, 59, 999);
}

function diferencaPercentualDashboardExecutivo(atual, anterior) {
  const atualSeguro = round2Financeiro(atual);
  const anteriorSeguro = round2Financeiro(anterior);
  return {
    valor_anterior: anteriorSeguro,
    variacao_absoluta: round2Financeiro(atualSeguro - anteriorSeguro),
    variacao_percentual: Math.abs(anteriorSeguro) <= 0.009
      ? null
      : round2Financeiro(((atualSeguro - anteriorSeguro) / Math.abs(anteriorSeguro)) * 100)
  };
}

function criarMapaDashboardExecutivo(lista, campoId) {
  const mapa = {};
  (Array.isArray(lista) ? lista : []).forEach(item => {
    const id = String(item?.[campoId] || '').trim();
    if (id) mapa[id] = item;
  });
  return mapa;
}

function obterNaturezaSeguraDashboardExecutivo(item) {
  try {
    return normalizarNaturezaFinanceiro(
      item?.natureza || getNaturezaOrigemFinanceiro(item?.origem_tipo)
    );
  } catch (error) {
    return '';
  }
}

function obterTipoOrigemSeguroDashboardExecutivo(valor) {
  try {
    return normalizarOrigemTipoFinanceiro(valor);
  } catch (error) {
    return '';
  }
}

function criarMetaOrigemDashboardExecutivo(contexto, origemTipo, origemId) {
  const tipo = String(origemTipo || '').trim().toUpperCase();
  const id = String(origemId || '').trim();
  if (!id) return null;

  if (tipo === ORIGEM_TIPO_COMPRA) {
    const item = contexto.comprasPorId[id];
    if (!item) return null;
    const tipoItem = normalizarTextoDashboardExecutivo(item.tipo, 'Sem tipo');
    const categoria = normalizarTextoDashboardExecutivo(item.categoria, 'Sem categoria');
    return {
      origem_tipo: tipo,
      origem_id: id,
      origem_aba: 'compras',
      titulo: `Compra: ${normalizarTextoDashboardExecutivo(item.item, id)}`,
      detalhe: `${tipoItem} | ${categoria}`,
      categoria,
      item: normalizarTextoDashboardExecutivo(item.item, id),
      pago_por: normalizarTextoDashboardExecutivo(item.pago_por, ''),
      fornecedor: normalizarTextoDashboardExecutivo(item.fornecedor, 'Sem fornecedor')
    };
  }

  if (tipo === ORIGEM_TIPO_DESPESA) {
    const item = contexto.despesasPorId[id];
    if (!item) return null;
    const categoria = normalizarTextoDashboardExecutivo(item.categoria, 'Sem categoria');
    const fornecedor = normalizarTextoDashboardExecutivo(item.fornecedor, 'Sem fornecedor');
    return {
      origem_tipo: tipo,
      origem_id: id,
      origem_aba: 'despesas',
      titulo: `Despesa: ${normalizarTextoDashboardExecutivo(item.descricao, id)}`,
      detalhe: `${categoria} | ${fornecedor}`,
      categoria,
      item: normalizarTextoDashboardExecutivo(item.descricao, id),
      pago_por: normalizarTextoDashboardExecutivo(item.pago_por, ''),
      fornecedor
    };
  }

  if (tipo === ORIGEM_TIPO_VENDA) {
    const item = contexto.vendasPorId[id];
    if (!item) return null;
    const nomeItem = normalizarTextoDashboardExecutivo(item.item, id);
    return {
      origem_tipo: tipo,
      origem_id: id,
      origem_aba: 'vendas',
      titulo: `Venda: ${nomeItem}`,
      detalhe: `Qtd: ${round2Financeiro(parseNumeroBR(item.quantidade))} ${normalizarTextoDashboardExecutivo(item.unidade, '')}`.trim(),
      categoria: normalizarTextoDashboardExecutivo(item.categoria, 'Sem categoria'),
      item: nomeItem,
      pago_por: normalizarTextoDashboardExecutivo(item.recebido_por, ''),
      fornecedor: ''
    };
  }

  if (tipo === ORIGEM_TIPO_ESTOQUE) {
    const item = contexto.estoquePorId[id];
    if (!item) return null;
    const tipoItem = normalizarTextoDashboardExecutivo(item.tipo, 'Sem tipo');
    const categoria = normalizarTextoDashboardExecutivo(item.categoria, 'Sem categoria');
    return {
      origem_tipo: tipo,
      origem_id: id,
      origem_aba: 'estoque',
      titulo: `Estoque: ${normalizarTextoDashboardExecutivo(item.item, id)}`,
      detalhe: `${tipoItem} | ${categoria}`,
      categoria,
      item: normalizarTextoDashboardExecutivo(item.item, id),
      tipo_item: tipoItem,
      pago_por: normalizarTextoDashboardExecutivo(item.pago_por, ''),
      fornecedor: normalizarTextoDashboardExecutivo(item.fornecedor, '')
    };
  }

  if (tipo === ORIGEM_TIPO_PRODUCAO) {
    const item = contexto.producoesPorId[id];
    if (!item) return null;
    const nome = normalizarTextoDashboardExecutivo(
      item.nome_ordem,
      normalizarTextoDashboardExecutivo(item.nome_produto, id)
    );
    return {
      origem_tipo: tipo,
      origem_id: id,
      origem_aba: 'producao',
      titulo: `Producao: ${nome}`,
      detalhe: `Status: ${normalizarTextoDashboardExecutivo(item.status, 'Sem status')}`,
      categoria: normalizarTextoDashboardExecutivo(item.status, 'Sem status'),
      item: nome,
      pago_por: '',
      fornecedor: ''
    };
  }

  return null;
}

function lerParcelasAtivasDashboardExecutivo(contexto) {
  const sheet = getSheet(ABA_PARCELAS_FINANCEIRAS);
  if (!sheet) return [];

  return rowsToObjects(sheet)
    .filter(item => String(item.ativo).toLowerCase() === 'true')
    .map(item => {
      const origemTipo = obterTipoOrigemSeguroDashboardExecutivo(item.origem_tipo);
      const origemId = String(item.origem_id || '').trim();
      const meta = criarMetaOrigemDashboardExecutivo(contexto, origemTipo, origemId);
      if (!meta) return null;
      const valorPrevisto = round2Financeiro(parseNumeroBR(item.valor_previsto));
      const valorPago = round2Financeiro(parseNumeroBR(item.valor_pago));
      return {
        ID: String(item.ID || '').trim(),
        origem_tipo: origemTipo,
        origem_id: origemId,
        natureza: obterNaturezaSeguraDashboardExecutivo({
          natureza: item.natureza,
          origem_tipo: origemTipo
        }),
        parcela_numero: Number(item.parcela_numero || 0),
        parcelas_total: Number(item.parcelas_total || 0),
        data_prevista: formatarDataYmdFinanceiroSafe(item.data_prevista),
        valor_previsto: valorPrevisto,
        valor_pago: valorPago,
        valor_pendente: round2Financeiro(Math.max(0, valorPrevisto - valorPago)),
        meta
      };
    })
    .filter(Boolean);
}

function criarObrigacoesDashboardExecutivo(contexto) {
  const obrigacoes = [];
  const origensComParcelas = {};

  contexto.parcelas.forEach(parcela => {
    const chave = `${parcela.origem_tipo}|${parcela.origem_id}`;
    origensComParcelas[chave] = true;
    if (parcela.valor_pendente <= 0.009 || !parcela.natureza) return;
    obrigacoes.push({
      id: parcela.ID || `PAR:${chave}:${parcela.parcela_numero}`,
      origem_tipo: parcela.origem_tipo,
      origem_id: parcela.origem_id,
      natureza: parcela.natureza,
      data_prevista: parcela.data_prevista,
      valor: parcela.valor_pendente,
      parcela_numero: parcela.parcela_numero,
      parcelas_total: parcela.parcelas_total,
      meta: parcela.meta
    });
  });

  const adicionarFallback = (lista, origemTipo, campoData) => {
    (Array.isArray(lista) ? lista : []).forEach(item => {
      const origemId = String(item.ID || '').trim();
      if (!origemId || origensComParcelas[`${origemTipo}|${origemId}`]) return;
      const valor = round2Financeiro(parseNumeroBR(item.total_pendente));
      if (valor <= 0.009) return;
      const meta = criarMetaOrigemDashboardExecutivo(contexto, origemTipo, origemId);
      if (!meta) return;
      obrigacoes.push({
        id: `ORIG:${origemTipo}:${origemId}`,
        origem_tipo: origemTipo,
        origem_id: origemId,
        natureza: getNaturezaOrigemFinanceiro(origemTipo),
        data_prevista: formatarDataYmdFinanceiroSafe(item[campoData]),
        valor,
        parcela_numero: 0,
        parcelas_total: 0,
        meta
      });
    });
  };

  adicionarFallback(contexto.compras, ORIGEM_TIPO_COMPRA, 'data_vencimento');
  adicionarFallback(contexto.despesas, ORIGEM_TIPO_DESPESA, 'data_vencimento');
  // VENDAS nao possui vencimento no cabecalho da origem. Sem uma parcela
  // financeira, a data da venda nao deve ser tratada como vencimento.
  adicionarFallback(contexto.vendas, ORIGEM_TIPO_VENDA, 'data_vencimento');
  return obrigacoes;
}

function criarContextoDashboardExecutivo(forcarRecarregar) {
  const force = !!forcarRecarregar;
  // Pagamentos vem primeiro para que uma atualizacao forcada alimente com dados
  // frescos os resumos financeiros usados pelas listas de origem.
  const pagamentos = typeof listarPagamentos === 'function' ? listarPagamentos(force) : [];
  const contexto = {
    compras: typeof listarCompras === 'function' ? listarCompras(force) : [],
    despesas: typeof listarDespesasGerais === 'function' ? listarDespesasGerais(force) : [],
    vendas: typeof listarVendas === 'function' ? listarVendas(force) : [],
    pagamentos,
    estoque: typeof listarEstoque === 'function' ? listarEstoque(force) : [],
    producoes: typeof listarProducao === 'function' ? listarProducao(force) : []
  };

  contexto.comprasPorId = criarMapaDashboardExecutivo(contexto.compras, 'ID');
  contexto.despesasPorId = criarMapaDashboardExecutivo(contexto.despesas, 'ID');
  contexto.vendasPorId = criarMapaDashboardExecutivo(contexto.vendas, 'ID');
  contexto.estoquePorId = criarMapaDashboardExecutivo(contexto.estoque, 'ID');
  contexto.producoesPorId = criarMapaDashboardExecutivo(contexto.producoes, 'producao_id');
  contexto.parcelas = lerParcelasAtivasDashboardExecutivo(contexto);
  contexto.obrigacoes = criarObrigacoesDashboardExecutivo(contexto);
  contexto.eventos = contexto.pagamentos
    .map(item => {
      const origemTipo = obterTipoOrigemSeguroDashboardExecutivo(item.origem_tipo);
      const origemId = String(item.origem_id || '').trim();
      const meta = criarMetaOrigemDashboardExecutivo(contexto, origemTipo, origemId);
      const natureza = obterNaturezaSeguraDashboardExecutivo({
        natureza: item.natureza,
        origem_tipo: origemTipo
      });
      const valor = round2Financeiro(parseNumeroBR(item.valor_pago));
      const data = formatarDataYmdFinanceiroSafe(item.data_pagamento);
      if (!meta || !natureza || valor <= 0 || !data) return null;
      return {
        ...item,
        origem_tipo: origemTipo,
        origem_id: origemId,
        natureza,
        valor,
        data,
        meta
      };
    })
    .filter(Boolean);
  return contexto;
}

function criarAgingDashboardExecutivo(obrigacoes, hoje) {
  const definicoes = [
    { id: 'vencido', label: 'Vencido' },
    { id: '0_7', label: '0 a 7 dias' },
    { id: '8_30', label: '8 a 30 dias' },
    { id: '31_60', label: '31 a 60 dias' },
    { id: '61_90', label: '61 a 90 dias' },
    { id: 'mais_90', label: 'Mais de 90 dias' },
    { id: 'sem_data', label: 'Sem data' }
  ];
  const mapa = {};
  definicoes.forEach(item => {
    mapa[item.id] = { ...item, receber: 0, pagar: 0, total_itens: 0 };
  });

  (Array.isArray(obrigacoes) ? obrigacoes : []).forEach(item => {
    const data = inicioDoDiaFinanceiro(item.data_prevista);
    let bucket = 'sem_data';
    if (data) {
      const diff = Math.floor((data.getTime() - hoje.getTime()) / 86400000);
      if (diff < 0) bucket = 'vencido';
      else if (diff <= 7) bucket = '0_7';
      else if (diff <= 30) bucket = '8_30';
      else if (diff <= 60) bucket = '31_60';
      else if (diff <= 90) bucket = '61_90';
      else bucket = 'mais_90';
    }
    const campo = item.natureza === NATUREZA_RECEBIMENTO ? 'receber' : 'pagar';
    mapa[bucket][campo] = round2Financeiro(mapa[bucket][campo] + item.valor);
    mapa[bucket].total_itens += 1;
  });

  return definicoes.map(item => mapa[item.id]);
}

function criarProjecaoDashboardExecutivo(obrigacoes, hoje) {
  const semanas = [];
  for (let indice = 0; indice < 13; indice += 1) {
    const inicio = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + (indice * 7));
    const fim = new Date(inicio.getFullYear(), inicio.getMonth(), inicio.getDate() + 6);
    semanas.push({
      id: `semana_${indice + 1}`,
      indice: indice + 1,
      inicio: formatarDataYmdFinanceiro(inicio),
      fim: formatarDataYmdFinanceiro(fim),
      recebimentos: 0,
      pagamentos: 0,
      liquido: 0,
      acumulado: 0
    });
  }

  const limite = fimDoDiaDashboardExecutivo(new Date(
    hoje.getFullYear(),
    hoje.getMonth(),
    hoje.getDate() + 90
  ));
  (Array.isArray(obrigacoes) ? obrigacoes : []).forEach(item => {
    const data = inicioDoDiaFinanceiro(item.data_prevista);
    if (!data || data < hoje || data > limite) return;
    const indice = Math.min(12, Math.floor((data.getTime() - hoje.getTime()) / 86400000 / 7));
    const campo = item.natureza === NATUREZA_RECEBIMENTO ? 'recebimentos' : 'pagamentos';
    semanas[indice][campo] = round2Financeiro(semanas[indice][campo] + item.valor);
  });

  let acumulado = 0;
  semanas.forEach(item => {
    item.liquido = round2Financeiro(item.recebimentos - item.pagamentos);
    acumulado = round2Financeiro(acumulado + item.liquido);
    item.acumulado = acumulado;
  });
  return semanas;
}

function agregarTopDashboardExecutivo(mapa, limite) {
  const lista = Object.keys(mapa || {})
    .map(chave => ({
      id: normalizarChaveDashboardExecutivo(chave) || 'sem_identificacao',
      label: chave,
      valor: round2Financeiro(mapa[chave] || 0),
      membros: [chave]
    }))
    .filter(item => item.valor > 0)
    .sort((a, b) => b.valor - a.valor || a.label.localeCompare(b.label));
  const maximo = Math.max(1, Number(limite || 7));
  if (lista.length <= maximo) return lista;
  const principais = lista.slice(0, maximo);
  const cauda = lista.slice(maximo);
  const outros = cauda.reduce((acc, item) => acc + item.valor, 0);
  principais.push({
    id: 'outros',
    label: 'Outros',
    valor: round2Financeiro(outros),
    membros: cauda.reduce((acc, item) => acc.concat(item.membros || [item.label]), [])
  });
  return principais;
}

function adicionarPercentuaisDashboardExecutivo(lista) {
  const total = (Array.isArray(lista) ? lista : []).reduce((acc, item) => acc + Number(item.valor || 0), 0);
  return (Array.isArray(lista) ? lista : []).map(item => ({
    ...item,
    percentual: total <= 0 ? 0 : round2Financeiro((Number(item.valor || 0) / total) * 100)
  }));
}

function criarHistoricoDashboardExecutivo(contexto, referencia) {
  const meses = [];
  const porMes = {};
  for (let indice = 11; indice >= 0; indice -= 1) {
    const mes = deslocarMesDashboardExecutivo(referencia, -indice);
    porMes[mes] = { mes, vendas: 0, recebimentos: 0, pagamentos: 0, fluxo_liquido: 0, media_movel_3: 0 };
    meses.push(porMes[mes]);
  }

  contexto.vendas.forEach(item => {
    const mes = gerarChaveMesFinanceiro(item.data_venda);
    if (!porMes[mes]) return;
    porMes[mes].vendas = round2Financeiro(porMes[mes].vendas + parseNumeroBR(item.valor_total_venda));
  });
  contexto.eventos.forEach(item => {
    const mes = gerarChaveMesFinanceiro(item.data);
    if (!porMes[mes]) return;
    const campo = item.natureza === NATUREZA_RECEBIMENTO ? 'recebimentos' : 'pagamentos';
    porMes[mes][campo] = round2Financeiro(porMes[mes][campo] + item.valor);
  });

  meses.forEach((item, indice) => {
    item.fluxo_liquido = round2Financeiro(item.recebimentos - item.pagamentos);
    const inicioMedia = Math.max(0, indice - 2);
    const janela = meses.slice(inicioMedia, indice + 1);
    item.media_movel_3 = round2Financeiro(
      janela.reduce((acc, mes) => acc + mes.fluxo_liquido, 0) / janela.length
    );
  });
  return meses;
}

function criarCapitalDashboardExecutivo(contexto, hoje) {
  const custoPorTipo = {};
  let estoqueCusto = 0;
  let produtosCusto = 0;
  let insumosCusto = 0;
  let potencialEstoque = 0;
  let itensSemCusto = 0;

  contexto.estoque.forEach(item => {
    const quantidade = Math.max(0, parseNumeroBR(item.quantidade));
    const custoUnitario = Math.max(0, parseNumeroBR(item.custo_unitario || item.valor_unit));
    const tipo = normalizarTextoDashboardExecutivo(item.tipo, 'Sem tipo');
    const custo = round2Financeiro(quantidade * custoUnitario);
    if (quantidade > 0 && custoUnitario <= 0) itensSemCusto += 1;
    estoqueCusto = round2Financeiro(estoqueCusto + custo);
    custoPorTipo[tipo] = round2Financeiro((custoPorTipo[tipo] || 0) + custo);
    if (normalizarTextoSemAcentoFinanceiro(tipo) === 'PRODUTO') {
      produtosCusto = round2Financeiro(produtosCusto + custo);
      potencialEstoque = round2Financeiro(
        potencialEstoque + round2Financeiro(quantidade * Math.max(0, parseNumeroBR(item.preco_venda)))
      );
    } else {
      insumosCusto = round2Financeiro(insumosCusto + custo);
    }
  });

  const statusAtivo = status => {
    const valor = normalizarTextoSemAcentoFinanceiro(status);
    return valor === 'EM PRODUCAO' || valor === 'FINALIZACAO';
  };
  const ordensAtivas = contexto.producoes.filter(item => statusAtivo(item.status));
  const idsAtivos = {};
  ordensAtivas.forEach(item => { idsAtivos[String(item.producao_id || '').trim()] = true; });

  let producaoCusto = 0;
  const consumoSheet = getSheet(ABA_PRODUCAO_CONSUMO);
  if (consumoSheet) {
    rowsToObjects(consumoSheet)
      .filter(item => String(item.ativo).toLowerCase() === 'true')
      .forEach(item => {
        if (!idsAtivos[String(item.producao_id || '').trim()]) return;
        producaoCusto = round2Financeiro(
          producaoCusto + Math.max(0, parseNumeroBR(item.total_snapshot))
        );
      });
  }

  let potencialProducao = 0;
  let ordensAtrasadas = 0;
  const porStatus = {};
  ordensAtivas.forEach(item => {
    const status = normalizarTextoDashboardExecutivo(item.status, 'Sem status');
    porStatus[status] = Number(porStatus[status] || 0) + 1;
    const planejada = Math.max(0, parseNumeroBR(item.qtd_planejada));
    const restante = Math.max(0, parseNumeroBR(item.qtd_restante));
    const potencialPlanejado = Math.max(0, parseNumeroBR(item.valor_previsto_venda_pecas));
    const proporcao = planejada > 0 ? Math.min(1, restante / planejada) : 0;
    potencialProducao = round2Financeiro(potencialProducao + (potencialPlanejado * proporcao));
    const prevista = inicioDoDiaFinanceiro(item.data_prevista_termino);
    if (prevista && prevista < hoje) ordensAtrasadas += 1;
  });

  return {
    estoque_custo_total: estoqueCusto,
    estoque_por_tipo: adicionarPercentuaisDashboardExecutivo(agregarTopDashboardExecutivo(custoPorTipo, 8)),
    produtos_custo: produtosCusto,
    insumos_custo: insumosCusto,
    producao_custo_em_andamento: producaoCusto,
    capital_operacional_total: round2Financeiro(estoqueCusto + producaoCusto),
    potencial_venda_estoque: potencialEstoque,
    potencial_venda_producao: potencialProducao,
    potencial_venda_total: round2Financeiro(potencialEstoque + potencialProducao),
    ordens_ativas: ordensAtivas.length,
    ordens_atrasadas: ordensAtrasadas,
    producao_por_status: Object.keys(porStatus)
      .map(status => ({
        id: normalizarChaveDashboardExecutivo(status),
        label: status,
        quantidade: porStatus[status]
      }))
      .sort((a, b) => b.quantidade - a.quantidade || a.label.localeCompare(b.label)),
    itens_sem_custo: itensSemCusto
  };
}

function criarSociosDashboardExecutivo(contexto, inicio, fim) {
  const totais = { bruno: 0, zizu: 0, investimento: 0 };
  let investimentoAcumulado = 0;
  let naoClassificado = 0;
  contexto.eventos.forEach(evento => {
    if (evento.natureza !== NATUREZA_PAGAMENTO) return;
    const rateio = calcularRateioLinhaPagoPorFinanceiro(evento.meta?.pago_por, evento.valor);
    investimentoAcumulado = round2Financeiro(investimentoAcumulado + rateio.investimento);
    const data = parseDataFinanceiro(evento.data);
    if (!data || data < inicio || data >= fim) return;
    const totalRateado = round2Financeiro(rateio.bruno + rateio.zizu + rateio.investimento);
    if (totalRateado <= 0.009) {
      naoClassificado = round2Financeiro(naoClassificado + evento.valor);
    }
    totais.bruno = round2Financeiro(totais.bruno + rateio.bruno);
    totais.zizu = round2Financeiro(totais.zizu + rateio.zizu);
    totais.investimento = round2Financeiro(totais.investimento + rateio.investimento);
  });
  const itens = adicionarPercentuaisDashboardExecutivo([
    { id: 'bruno', label: 'Bruno', valor: totais.bruno },
    { id: 'zizu', label: 'Zizu', valor: totais.zizu },
    { id: 'investimento', label: 'Investimento', valor: totais.investimento }
  ]);
  return {
    total_periodo: round2Financeiro(totais.bruno + totais.zizu + totais.investimento),
    bruno: totais.bruno,
    zizu: totais.zizu,
    investimento: totais.investimento,
    nao_classificado: naoClassificado,
    investimento_acumulado: investimentoAcumulado,
    itens
  };
}

function obterDashboardExecutivoFinanceiro(referenciaYm, forcarRecarregar) {
  const intervalo = getIntervaloMesFinanceiro(referenciaYm);
  const referencia = intervalo.ref;
  if (!forcarRecarregar) {
    const cached = appCacheGetJson(DASHBOARD_EXECUTIVO_CACHE_SCOPE);
    if (cached && cached.version === DASHBOARD_EXECUTIVO_VERSION && cached.referencia === referencia) {
      return cached;
    }
  }

  const contexto = criarContextoDashboardExecutivo(!!forcarRecarregar);
  const historico = criarHistoricoDashboardExecutivo(contexto, referencia);
  const atual = historico[historico.length - 1] || {};
  const anterior = historico[historico.length - 2] || {};
  const hoje = inicioDoDiaFinanceiro(new Date());
  const aging = criarAgingDashboardExecutivo(contexto.obrigacoes, hoje);
  const projecao = criarProjecaoDashboardExecutivo(contexto.obrigacoes, hoje);

  const receberTotal = round2Financeiro(contexto.obrigacoes
    .filter(item => item.natureza === NATUREZA_RECEBIMENTO)
    .reduce((acc, item) => acc + item.valor, 0));
  const pagarTotal = round2Financeiro(contexto.obrigacoes
    .filter(item => item.natureza === NATUREZA_PAGAMENTO)
    .reduce((acc, item) => acc + item.valor, 0));
  const vencido = aging.find(item => item.id === 'vencido') || {};

  const gastosCategoria = {};
  contexto.eventos.forEach(evento => {
    const data = parseDataFinanceiro(evento.data);
    if (
      evento.natureza !== NATUREZA_PAGAMENTO ||
      !data || data < intervalo.inicio || data >= intervalo.fim
    ) return;
    const categoria = normalizarTextoDashboardExecutivo(evento.meta?.categoria, 'Sem categoria');
    gastosCategoria[categoria] = round2Financeiro((gastosCategoria[categoria] || 0) + evento.valor);
  });

  const vendasProduto = {};
  contexto.vendas.forEach(venda => {
    const data = parseDataFinanceiro(venda.data_venda);
    if (!data || data < intervalo.inicio || data >= intervalo.fim) return;
    const item = normalizarTextoDashboardExecutivo(venda.item, 'Sem identificacao');
    vendasProduto[item] = round2Financeiro(
      (vendasProduto[item] || 0) + Math.max(0, parseNumeroBR(venda.valor_total_venda))
    );
  });

  const gastosLista = adicionarPercentuaisDashboardExecutivo(
    agregarTopDashboardExecutivo(gastosCategoria, 7)
  );
  const vendasLista = adicionarPercentuaisDashboardExecutivo(
    agregarTopDashboardExecutivo(vendasProduto, 7)
  );
  const chavesProdutosVendidos = Object.keys(vendasProduto)
    .filter(chave => round2Financeiro(vendasProduto[chave]) > 0);
  const totalVendasProdutos = chavesProdutosVendidos
    .reduce((acc, chave) => acc + round2Financeiro(vendasProduto[chave]), 0);
  const socios = criarSociosDashboardExecutivo(contexto, intervalo.inicio, intervalo.fim);
  const capital = criarCapitalDashboardExecutivo(contexto, hoje);
  const resultado = {
    version: DASHBOARD_EXECUTIVO_VERSION,
    referencia,
    gerado_em: new Date().toISOString(),
    posicao_em: formatarDataYmdFinanceiro(hoje),
    kpis: {
      vendas: {
        valor: round2Financeiro(atual.vendas),
        ...diferencaPercentualDashboardExecutivo(atual.vendas, anterior.vendas)
      },
      recebimentos: {
        valor: round2Financeiro(atual.recebimentos),
        ...diferencaPercentualDashboardExecutivo(atual.recebimentos, anterior.recebimentos)
      },
      pagamentos: {
        valor: round2Financeiro(atual.pagamentos),
        ...diferencaPercentualDashboardExecutivo(atual.pagamentos, anterior.pagamentos)
      },
      fluxo_liquido: {
        valor: round2Financeiro(atual.fluxo_liquido),
        ...diferencaPercentualDashboardExecutivo(atual.fluxo_liquido, anterior.fluxo_liquido)
      }
    },
    historico_12_meses: historico,
    projecao_13_semanas: projecao,
    posicao: {
      receber_total: receberTotal,
      pagar_total: pagarTotal,
      receber_vencido: round2Financeiro(vencido.receber),
      pagar_vencido: round2Financeiro(vencido.pagar),
      aging
    },
    direcionadores: {
      gastos_categoria: gastosLista,
      vendas_produto: vendasLista,
      media_vendas_produto: chavesProdutosVendidos.length
        ? round2Financeiro(totalVendasProdutos / chavesProdutosVendidos.length)
        : 0
    },
    capital,
    socios,
    qualidade: {
      sem_vencimento_receber: round2Financeiro((aging.find(item => item.id === 'sem_data') || {}).receber),
      sem_vencimento_pagar: round2Financeiro((aging.find(item => item.id === 'sem_data') || {}).pagar),
      itens_sem_custo: capital.itens_sem_custo,
      pagamentos_sem_responsavel: socios.nao_classificado
    }
  };

  appCachePutJson(
    DASHBOARD_EXECUTIVO_CACHE_SCOPE,
    resultado,
    DASHBOARD_EXECUTIVO_CACHE_TTL_SEC
  );
  return resultado;
}

function criarLinhaEventoDashboardExecutivo(evento, valorOverride, detalheExtra) {
  const pagamento = evento.natureza === NATUREZA_PAGAMENTO;
  return {
    linha_id: `PGT:${String(evento.ID || '').trim()}`,
    tipo_linha: 'pagamento',
    titulo: evento.meta.titulo,
    detalhe: `${evento.meta.detalhe} | ${normalizarTextoDashboardExecutivo(detalheExtra, pagamento ? 'Pagamento' : 'Recebimento')}`,
    data: evento.data,
    valor: round2Financeiro(valorOverride === undefined ? evento.valor : valorOverride),
    origem_tipo: evento.meta.origem_tipo,
    origem_id: evento.meta.origem_id,
    origem_aba: evento.meta.origem_aba,
    pagamento_id: String(evento.ID || '').trim(),
    pode_editar_origem: true,
    pode_excluir_origem: true,
    pode_remover_pagamento: true,
    eh_contagem: false
  };
}

function criarLinhaObrigacaoDashboardExecutivo(item) {
  const parcela = item.parcela_numero > 0
    ? ` | Parcela ${item.parcela_numero}/${item.parcelas_total || item.parcela_numero}`
    : '';
  return {
    linha_id: `PEND:${item.id}`,
    tipo_linha: 'parcela_pendente',
    titulo: item.meta.titulo,
    detalhe: `${item.meta.detalhe}${parcela} | Pendente`,
    data: item.data_prevista,
    valor: round2Financeiro(item.valor),
    origem_tipo: item.meta.origem_tipo,
    origem_id: item.meta.origem_id,
    origem_aba: item.meta.origem_aba,
    pagamento_id: '',
    pode_editar_origem: true,
    pode_excluir_origem: true,
    pode_remover_pagamento: false,
    eh_contagem: false
  };
}

function obterDetalhesDashboardExecutivoFinanceiro(referenciaYm, contextoDetalhe, filtroInput) {
  const intervalo = getIntervaloMesFinanceiro(referenciaYm);
  const contexto = criarContextoDashboardExecutivo(false);
  const chave = String(contextoDetalhe || '').trim().toLowerCase();
  const filtro = filtroInput && typeof filtroInput === 'object' ? filtroInput : {};
  const hoje = inicioDoDiaFinanceiro(new Date());
  let titulo = 'Detalhes';
  let itens = [];

  const noMes = dataValor => {
    const data = parseDataFinanceiro(dataValor);
    return !!data && data >= intervalo.inicio && data < intervalo.fim;
  };
  const eventosDoMes = contexto.eventos.filter(item => noMes(item.data));
  const eventosPorNatureza = natureza => eventosDoMes.filter(item => item.natureza === natureza);
  const bucketObrigacao = item => {
    const data = inicioDoDiaFinanceiro(item.data_prevista);
    if (!data) return 'sem_data';
    const diff = Math.floor((data.getTime() - hoje.getTime()) / 86400000);
    if (diff < 0) return 'vencido';
    if (diff <= 7) return '0_7';
    if (diff <= 30) return '8_30';
    if (diff <= 60) return '31_60';
    if (diff <= 90) return '61_90';
    return 'mais_90';
  };
  const statusProducaoAtivo = status => {
    const valor = normalizarTextoSemAcentoFinanceiro(status);
    return valor === 'EM PRODUCAO' || valor === 'FINALIZACAO';
  };
  const criarLinhaEstoque = (item, valor, detalheExtra) => {
    const meta = criarMetaOrigemDashboardExecutivo(contexto, ORIGEM_TIPO_ESTOQUE, item.ID);
    if (!meta) return null;
    return {
      linha_id: `EST:${item.ID}`,
      tipo_linha: 'estoque',
      titulo: meta.titulo,
      detalhe: `${meta.detalhe} | ${round2Financeiro(parseNumeroBR(item.quantidade))} ${normalizarTextoDashboardExecutivo(item.unidade, '')}${detalheExtra ? ` | ${detalheExtra}` : ''}`,
      data: formatarDataYmdFinanceiroSafe(item.comprado_em || item.criado_em),
      valor: round2Financeiro(valor),
      origem_tipo: meta.origem_tipo,
      origem_id: meta.origem_id,
      origem_aba: meta.origem_aba,
      pagamento_id: '',
      pode_editar_origem: true,
      pode_excluir_origem: true,
      pode_remover_pagamento: false,
      eh_contagem: false
    };
  };
  const linhasCustoEstoque = () => contexto.estoque.map(item => criarLinhaEstoque(
    item,
    Math.max(0, parseNumeroBR(item.quantidade)) * Math.max(0, parseNumeroBR(item.custo_unitario || item.valor_unit)),
    'Valor a custo'
  )).filter(item => item && item.valor > 0);
  const linhasCustoProducao = () => {
    const ativos = {};
    contexto.producoes.filter(item => statusProducaoAtivo(item.status)).forEach(item => {
      ativos[String(item.producao_id || '').trim()] = item;
    });
    const totais = {};
    const sheet = getSheet(ABA_PRODUCAO_CONSUMO);
    if (sheet) {
      rowsToObjects(sheet)
        .filter(item => String(item.ativo).toLowerCase() === 'true')
        .forEach(item => {
          const id = String(item.producao_id || '').trim();
          if (!ativos[id]) return;
          totais[id] = round2Financeiro((totais[id] || 0) + Math.max(0, parseNumeroBR(item.total_snapshot)));
        });
    }
    return Object.keys(totais).map(id => {
      const op = ativos[id];
      const meta = criarMetaOrigemDashboardExecutivo(contexto, ORIGEM_TIPO_PRODUCAO, id);
      if (!meta) return null;
      return {
        linha_id: `PRODCUSTO:${id}`,
        tipo_linha: 'producao',
        titulo: meta.titulo,
        detalhe: `${meta.detalhe} | Materiais consumidos`,
        data: formatarDataYmdFinanceiroSafe(op.data_inicio || op.criado_em),
        valor: round2Financeiro(totais[id]),
        origem_tipo: meta.origem_tipo,
        origem_id: meta.origem_id,
        origem_aba: meta.origem_aba,
        pagamento_id: '',
        pode_editar_origem: false,
        pode_excluir_origem: false,
        pode_remover_pagamento: false,
        eh_contagem: false
      };
    }).filter(item => item && item.valor > 0);
  };
  const linhasPotencialVenda = () => {
    const estoque = contexto.estoque
      .filter(item => normalizarTextoSemAcentoFinanceiro(item.tipo) === 'PRODUTO')
      .map(item => criarLinhaEstoque(
        item,
        Math.max(0, parseNumeroBR(item.quantidade)) * Math.max(0, parseNumeroBR(item.preco_venda)),
        'Potencial pelo preco de venda'
      ))
      .filter(item => item && item.valor > 0);
    const producao = contexto.producoes.filter(item => statusProducaoAtivo(item.status)).map(item => {
      const id = String(item.producao_id || '').trim();
      const meta = criarMetaOrigemDashboardExecutivo(contexto, ORIGEM_TIPO_PRODUCAO, id);
      if (!meta) return null;
      const planejada = Math.max(0, parseNumeroBR(item.qtd_planejada));
      const restante = Math.max(0, parseNumeroBR(item.qtd_restante));
      const proporcao = planejada > 0 ? Math.min(1, restante / planejada) : 0;
      const valor = round2Financeiro(Math.max(0, parseNumeroBR(item.valor_previsto_venda_pecas)) * proporcao);
      return {
        linha_id: `PRODPOT:${id}`,
        tipo_linha: 'producao',
        titulo: meta.titulo,
        detalhe: `${meta.detalhe} | Potencial das pecas restantes`,
        data: formatarDataYmdFinanceiroSafe(item.data_prevista_termino),
        valor,
        origem_tipo: meta.origem_tipo,
        origem_id: meta.origem_id,
        origem_aba: meta.origem_aba,
        pagamento_id: '',
        pode_editar_origem: false,
        pode_excluir_origem: false,
        pode_remover_pagamento: false,
        eh_contagem: false
      };
    }).filter(item => item && item.valor > 0);
    return estoque.concat(producao);
  };

  if (chave === 'kpi_vendas') {
    titulo = 'Vendas registradas no mes';
    itens = contexto.vendas.filter(item => noMes(item.data_venda)).map(item => {
      const meta = criarMetaOrigemDashboardExecutivo(contexto, ORIGEM_TIPO_VENDA, item.ID);
      return meta ? {
        linha_id: `VND:${item.ID}`,
        tipo_linha: 'origem',
        titulo: meta.titulo,
        detalhe: `${meta.detalhe} | Venda registrada`,
        data: formatarDataYmdFinanceiroSafe(item.data_venda),
        valor: round2Financeiro(parseNumeroBR(item.valor_total_venda)),
        origem_tipo: meta.origem_tipo,
        origem_id: meta.origem_id,
        origem_aba: meta.origem_aba,
        pagamento_id: '',
        pode_editar_origem: true,
        pode_excluir_origem: true,
        pode_remover_pagamento: false,
        eh_contagem: false
      } : null;
    }).filter(Boolean);
  } else if (chave === 'kpi_recebimentos' || chave === 'kpi_pagamentos' || chave === 'kpi_fluxo') {
    const natureza = chave === 'kpi_recebimentos' ? NATUREZA_RECEBIMENTO : NATUREZA_PAGAMENTO;
    titulo = chave === 'kpi_recebimentos'
      ? 'Recebimentos realizados no mes'
      : (chave === 'kpi_pagamentos' ? 'Pagamentos realizados no mes' : 'Fluxo financeiro realizado no mes');
    const base = chave === 'kpi_fluxo' ? eventosDoMes : eventosPorNatureza(natureza);
    itens = base.map(item => criarLinhaEventoDashboardExecutivo(
      item,
      chave === 'kpi_fluxo' && item.natureza === NATUREZA_PAGAMENTO ? -item.valor : item.valor,
      item.natureza === NATUREZA_PAGAMENTO ? 'Pagamento realizado' : 'Recebimento realizado'
    ));
  } else if (chave === 'historico_mes') {
    const mes = getMesReferenciaFinanceiro(filtro.mes || intervalo.ref);
    const serie = String(filtro.serie || 'fluxo').trim().toLowerCase();
    const janelaMedia = Math.max(1, Math.min(3, Number(filtro.janela || 3)));
    titulo = serie === 'media'
      ? `Composicao da media movel encerrada em ${mes}`
      : `Movimento realizado em ${mes}`;
    if (serie === 'vendas') {
      return obterDetalhesDashboardExecutivoFinanceiro(mes, 'kpi_vendas', {});
    }
    const refInicial = serie === 'media'
      ? deslocarMesDashboardExecutivo(mes, -(janelaMedia - 1))
      : mes;
    const base = contexto.eventos.filter(item => {
      const refEvento = gerarChaveMesFinanceiro(item.data);
      return !!refEvento && refEvento >= refInicial && refEvento <= mes;
    });
    itens = base
      .filter(item => serie === 'recebimentos'
        ? item.natureza === NATUREZA_RECEBIMENTO
        : (serie === 'pagamentos' ? item.natureza === NATUREZA_PAGAMENTO : true))
      .map(item => criarLinhaEventoDashboardExecutivo(
        item,
        (serie === 'fluxo' || serie === 'media') && item.natureza === NATUREZA_PAGAMENTO
          ? -item.valor / (serie === 'media' ? janelaMedia : 1)
          : item.valor / (serie === 'media' ? janelaMedia : 1),
        serie === 'media'
          ? 'Contribuicao para a media movel'
          : (item.natureza === NATUREZA_PAGAMENTO ? 'Pagamento realizado' : 'Recebimento realizado')
      ));
  } else if (chave === 'aging') {
    const natureza = String(filtro.natureza || '').toUpperCase() === 'RECEBER'
      ? NATUREZA_RECEBIMENTO
      : NATUREZA_PAGAMENTO;
    const bucket = String(filtro.bucket || 'vencido').trim();
    titulo = `${natureza === NATUREZA_RECEBIMENTO ? 'Contas a receber' : 'Contas a pagar'} - ${normalizarTextoDashboardExecutivo(filtro.label, bucket)}`;
    itens = contexto.obrigacoes
      .filter(item => item.natureza === natureza && bucketObrigacao(item) === bucket)
      .map(criarLinhaObrigacaoDashboardExecutivo);
  } else if (chave === 'posicao_receber' || chave === 'posicao_pagar') {
    const natureza = chave === 'posicao_receber' ? NATUREZA_RECEBIMENTO : NATUREZA_PAGAMENTO;
    titulo = natureza === NATUREZA_RECEBIMENTO ? 'Posicao de contas a receber' : 'Posicao de contas a pagar';
    itens = contexto.obrigacoes.filter(item => item.natureza === natureza).map(criarLinhaObrigacaoDashboardExecutivo);
  } else if (chave === 'projecao_semana') {
    const inicioSemana = filtro.acumulado
      ? hoje
      : inicioDoDiaFinanceiro(filtro.inicio);
    const fimSemana = fimDoDiaDashboardExecutivo(filtro.fim);
    const naturezaFiltro = String(filtro.natureza || '').trim().toUpperCase();
    titulo = filtro.acumulado
      ? `Movimento projetado acumulado ate ${formatarDataYmdFinanceiroSafe(filtro.fim)}`
      : `Movimento projetado - ${normalizarTextoDashboardExecutivo(filtro.label, 'semana')}`;
    itens = contexto.obrigacoes
      .filter(item => {
        const data = inicioDoDiaFinanceiro(item.data_prevista);
        if (!data || !inicioSemana || !fimSemana || data < inicioSemana || data > fimSemana) return false;
        if (naturezaFiltro === 'RECEBER') return item.natureza === NATUREZA_RECEBIMENTO;
        if (naturezaFiltro === 'PAGAR') return item.natureza === NATUREZA_PAGAMENTO;
        return true;
      })
      .map(item => {
        const linha = criarLinhaObrigacaoDashboardExecutivo(item);
        if (!naturezaFiltro && item.natureza === NATUREZA_PAGAMENTO) {
          linha.valor = round2Financeiro(-linha.valor);
        }
        return linha;
      });
  } else if (chave === 'gasto_categoria') {
    const label = String(filtro.label || '').trim();
    const membros = Array.isArray(filtro.membros) ? filtro.membros.map(String) : [];
    titulo = `Pagamentos - ${label || 'categoria'}`;
    itens = eventosPorNatureza(NATUREZA_PAGAMENTO)
      .filter(item => {
        const categoria = normalizarTextoDashboardExecutivo(item.meta?.categoria, 'Sem categoria');
        return membros.length ? membros.includes(categoria) : categoria === label;
      })
      .map(item => criarLinhaEventoDashboardExecutivo(item, item.valor, 'Pagamento realizado'));
  } else if (chave === 'venda_produto') {
    const label = String(filtro.label || '').trim();
    const membros = Array.isArray(filtro.membros) ? filtro.membros.map(String) : [];
    titulo = `Vendas - ${label || 'produto'}`;
    itens = contexto.vendas
      .filter(item => noMes(item.data_venda))
      .filter(item => {
        const nomeItem = normalizarTextoDashboardExecutivo(item.item, 'Sem identificacao');
        return membros.length ? membros.includes(nomeItem) : nomeItem === label;
      })
      .map(item => {
        const meta = criarMetaOrigemDashboardExecutivo(contexto, ORIGEM_TIPO_VENDA, item.ID);
        if (!meta) return null;
        return {
          linha_id: `VND:${item.ID}`,
          tipo_linha: 'origem',
          titulo: meta.titulo,
          detalhe: `${meta.detalhe} | Venda registrada`,
          data: formatarDataYmdFinanceiroSafe(item.data_venda),
          valor: round2Financeiro(parseNumeroBR(item.valor_total_venda)),
          origem_tipo: meta.origem_tipo,
          origem_id: meta.origem_id,
          origem_aba: meta.origem_aba,
          pagamento_id: '',
          pode_editar_origem: true,
          pode_excluir_origem: true,
          pode_remover_pagamento: false,
          eh_contagem: false
        };
      }).filter(Boolean);
  } else if (chave === 'capital_operacional' || chave === 'capital_estoque' || chave === 'capital_producao' || chave === 'capital_potencial') {
    if (chave === 'capital_operacional') {
      titulo = 'Capital operacional aplicado a custo';
      itens = linhasCustoEstoque().concat(linhasCustoProducao());
    } else if (chave === 'capital_estoque') {
      titulo = 'Estoque avaliado a custo';
      itens = linhasCustoEstoque();
    } else if (chave === 'capital_producao') {
      titulo = 'Custos consumidos nas ordens em andamento';
      itens = linhasCustoProducao();
    } else {
      titulo = 'Potencial de venda do estoque e da producao';
      itens = linhasPotencialVenda();
    }
  } else if (chave === 'capital_tipo') {
    const label = String(filtro.label || '').trim();
    const membros = Array.isArray(filtro.membros) ? filtro.membros.map(String) : [];
    titulo = `Estoque a custo - ${label || 'tipo'}`;
    itens = contexto.estoque
      .filter(item => {
        const tipo = normalizarTextoDashboardExecutivo(item.tipo, 'Sem tipo');
        return membros.length ? membros.includes(tipo) : tipo === label;
      })
      .map(item => criarLinhaEstoque(
        item,
        Math.max(0, parseNumeroBR(item.quantidade)) * Math.max(0, parseNumeroBR(item.custo_unitario || item.valor_unit)),
        'Valor a custo'
      )).filter(Boolean);
  } else if (chave === 'producao_status') {
    const label = String(filtro.label || '').trim();
    titulo = `Ordens de producao - ${label || 'status'}`;
    itens = contexto.producoes
      .filter(item => normalizarTextoDashboardExecutivo(item.status, 'Sem status') === label)
      .map(item => {
        const meta = criarMetaOrigemDashboardExecutivo(contexto, ORIGEM_TIPO_PRODUCAO, item.producao_id);
        if (!meta) return null;
        const planejada = Math.max(0, parseNumeroBR(item.qtd_planejada));
        const restante = Math.max(0, parseNumeroBR(item.qtd_restante));
        const proporcao = planejada > 0 ? Math.min(1, restante / planejada) : 0;
        const potencialRestante = round2Financeiro(
          Math.max(0, parseNumeroBR(item.valor_previsto_venda_pecas)) * proporcao
        );
        return {
          linha_id: `PROD:${item.producao_id}`,
          tipo_linha: 'producao',
          titulo: meta.titulo,
          detalhe: `${meta.detalhe} | Restante: ${round2Financeiro(restante)} | Potencial de venda`,
          data: formatarDataYmdFinanceiroSafe(item.data_prevista_termino),
          valor: potencialRestante,
          origem_tipo: meta.origem_tipo,
          origem_id: meta.origem_id,
          origem_aba: meta.origem_aba,
          pagamento_id: '',
          pode_editar_origem: false,
          pode_excluir_origem: false,
          pode_remover_pagamento: false,
          eh_contagem: false
        };
      }).filter(Boolean);
  } else if (chave === 'socio' || chave === 'socio_acumulado') {
    const socio = String(filtro.socio || '').trim().toLowerCase();
    titulo = chave === 'socio_acumulado'
      ? `Historico de pagamentos atribuidos a ${normalizarTextoDashboardExecutivo(filtro.label, socio)}`
      : `Pagamentos atribuidos a ${normalizarTextoDashboardExecutivo(filtro.label, socio)}`;
    const eventosSocio = chave === 'socio_acumulado'
      ? contexto.eventos.filter(evento => evento.natureza === NATUREZA_PAGAMENTO)
      : eventosPorNatureza(NATUREZA_PAGAMENTO);
    itens = eventosSocio.map(evento => {
      const rateio = calcularRateioLinhaPagoPorFinanceiro(evento.meta?.pago_por, evento.valor);
      const valor = round2Financeiro(rateio[socio] || 0);
      return valor > 0 ? criarLinhaEventoDashboardExecutivo(evento, valor, 'Parcela atribuida') : null;
    }).filter(Boolean);
  } else {
    throw new Error('Contexto de detalhes do dashboard invalido.');
  }

  itens.sort((a, b) => {
    const da = parseDataFinanceiro(a.data)?.getTime() || 0;
    const db = parseDataFinanceiro(b.data)?.getTime() || 0;
    if (da !== db) return db - da;
    return Math.abs(Number(b.valor || 0)) - Math.abs(Number(a.valor || 0));
  });
  return {
    referencia: intervalo.ref,
    contexto: chave,
    card_key: chave,
    card_label: titulo,
    campo_valor: 'valor',
    total: round2Financeiro(itens.reduce((acc, item) => acc + Number(item.valor || 0), 0)),
    total_itens: itens.length,
    itens
  };
}

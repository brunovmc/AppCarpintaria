const COMPROVANTE_CLASSIFICACAO_NAO_CLASSIFICADO = 'NAO_CLASSIFICADO';
const COMPROVANTE_CLASSIFICACAO_DESPESA = 'DESPESA_SIMPLES';
const COMPROVANTE_CLASSIFICACAO_COMPRA = 'COMPRA';
const COMPROVANTE_CLASSIFICACAO_MISTO = 'MISTO';
const COMPROVANTE_STATUS_PENDENTE = 'PENDENTE';
const COMPROVANTE_STATUS_PARCIAL = 'PARCIAL';
const COMPROVANTE_STATUS_CONCILIADO = 'CONCILIADO';

function executarComAmbienteComprovantesFinanceiros_(ambiente, callback) {
  if (typeof executarComAmbienteBancoDadosAutorizado_ !== 'function') {
    throw new Error('Controle de ambiente indisponivel.');
  }
  return executarComAmbienteBancoDadosAutorizado_(ambiente, () => callback());
}

function obterContextoComprovantesFinanceiros(statusFiltro, ambiente) {
  return executarComAmbienteComprovantesFinanceiros_(ambiente, () =>
    obterContextoComprovantesFinanceirosAtual_(statusFiltro)
  );
}

function obterContextoComprovantesFinanceirosAtual_(statusFiltro) {
  if (typeof assertCanRead === 'function') assertCanRead('Conciliacao de comprovantes financeiros');
  tentarReconciliarEstruturaComprovantesDriveNoAcesso_();
  const catalogo = montarCatalogoFinanceiroComprovantes_();
  const pagamentos = listarPagamentos(true);
  const todos = listarInboxDespesasNoAmbienteAtual_('TODOS')
    .filter(item => String(item.status || '').trim().toUpperCase() !== 'DESCARTADO')
    .map(item => enriquecerComprovanteFinanceiro_(item, pagamentos, catalogo));
  const filtro = String(statusFiltro || 'PENDENTES').trim().toUpperCase();
  const comprovantes = todos.filter(item => {
    const status = String(item.status_conciliacao || item.status || '').trim().toUpperCase();
    if (filtro === 'TODOS' || filtro === 'ALL') return true;
    if (filtro === 'CONCILIADOS' || filtro === 'CONCILIADO') {
      return status === COMPROVANTE_STATUS_CONCILIADO || status === 'CONFIRMADO';
    }
    return [COMPROVANTE_STATUS_PENDENTE, COMPROVANTE_STATUS_PARCIAL, 'ERRO'].includes(status);
  });
  const destinosMap = {};
  catalogo.pendentes.slice(0, 250).forEach(destino => {
    destinosMap[destino.chave] = destino;
  });
  comprovantes.forEach(comprovante => {
    (comprovante.sugestoes || []).forEach(sugestao => {
      const destino = catalogo.pendentes.find(item => item.chave === sugestao.chave);
      if (destino) destinosMap[destino.chave] = destino;
    });
  });
  return {
    comprovantes,
    destinos: Object.values(destinosMap),
    resumo: {
      total: todos.length,
      pendentes: todos.filter(item => item.status_conciliacao === COMPROVANTE_STATUS_PENDENTE).length,
      parciais: todos.filter(item => item.status_conciliacao === COMPROVANTE_STATUS_PARCIAL).length,
      conciliados: todos.filter(item => item.status_conciliacao === COMPROVANTE_STATUS_CONCILIADO).length,
      erros: todos.filter(item => item.status_conciliacao === 'ERRO').length
    }
  };
}

function montarCatalogoFinanceiroComprovantes_() {
  const compras = typeof listarCompras === 'function' ? listarCompras(true) : [];
  const sheetDespesas = getSheet(ABA_DESPESAS_GERAIS);
  const despesas = sheetDespesas
    ? rowsToObjects(sheetDespesas).filter(item => String(item.ativo).toLowerCase() === 'true')
    : [];
  const origens = {};
  compras.forEach(item => {
    const id = String(item.ID || '').trim();
    if (!id) return;
    origens[`COMPRA|${id}`] = {
      origem_tipo: 'COMPRA', origem_id: id,
      origem_rotulo: String(item.item || 'Compra').trim() || 'Compra',
      fornecedor: String(item.fornecedor || '').trim(),
      pago_por: String(item.pago_por || '').trim(),
      data_origem: formatarDataYmdFinanceiroSafe(item.comprado_em || item.criado_em),
      observacao: String(item.observacao || '').trim()
    };
  });
  despesas.forEach(item => {
    const id = String(item.ID || '').trim();
    if (!id) return;
    origens[`DESPESA_GERAL|${id}`] = {
      origem_tipo: 'DESPESA_GERAL', origem_id: id,
      origem_rotulo: String(item.descricao || 'Despesa').trim() || 'Despesa',
      fornecedor: String(item.fornecedor || '').trim(),
      pago_por: String(item.pago_por || '').trim(),
      data_origem: formatarDataYmdFinanceiroSafe(item.data_competencia || item.criado_em),
      observacao: String(item.observacao || '').trim()
    };
  });

  const sheetParcelas = getSheet(ABA_PARCELAS_FINANCEIRAS);
  const parcelas = sheetParcelas ? rowsToObjects(sheetParcelas) : [];
  const todos = parcelas
    .filter(parcela => {
      const tipo = String(parcela.origem_tipo || '').trim().toUpperCase();
      return String(parcela.ativo).toLowerCase() === 'true' &&
        ['COMPRA', 'DESPESA_GERAL'].includes(tipo) &&
        String(parcela.natureza || NATUREZA_PAGAMENTO).trim().toUpperCase() === NATUREZA_PAGAMENTO;
    })
    .map(parcela => {
      const tipo = String(parcela.origem_tipo || '').trim().toUpperCase();
      const origemId = String(parcela.origem_id || '').trim();
      const origem = origens[`${tipo}|${origemId}`];
      if (!origem) return null;
      const previsto = round2Financeiro(parseNumeroBR(parcela.valor_previsto));
      const pago = round2Financeiro(parseNumeroBR(parcela.valor_pago));
      const pendente = round2Financeiro(Math.max(0, previsto - pago));
      const numero = Math.max(1, Math.floor(parseNumeroBR(parcela.parcela_numero) || 1));
      const totalParcelas = Math.max(numero, Math.floor(parseNumeroBR(parcela.parcelas_total) || 1));
      const dataPrevista = formatarDataYmdFinanceiroSafe(parcela.data_prevista);
      const prefixo = tipo === 'COMPRA' ? 'Compra' : 'Despesa';
      return {
        ...origem,
        parcela_id: String(parcela.ID || '').trim(),
        parcela_numero: numero,
        parcelas_total: totalParcelas,
        data_prevista: dataPrevista,
        valor_previsto: previsto,
        valor_pago: pago,
        valor_pendente: pendente,
        chave: `${tipo}|${origemId}|${String(parcela.ID || '').trim()}`,
        label: `${prefixo} · ${origem.origem_rotulo} · parcela ${numero}/${totalParcelas} · ${formatarMoedaComprovante_(pendente)}`
      };
    })
    .filter(Boolean);

  const mapaPorParcela = {};
  todos.forEach(item => { mapaPorParcela[item.parcela_id] = item; });
  const pendentes = todos
    .filter(item => item.valor_pendente > 0.009)
    .sort((a, b) => {
      const da = parseDataFinanceiro(a.data_prevista)?.getTime() || Number.MAX_SAFE_INTEGER;
      const db = parseDataFinanceiro(b.data_prevista)?.getTime() || Number.MAX_SAFE_INTEGER;
      return da - db || b.valor_pendente - a.valor_pendente || a.label.localeCompare(b.label);
    });
  return { todos, pendentes, mapaPorParcela, origens };
}

function enriquecerComprovanteFinanceiro_(item, pagamentos, catalogo) {
  const comprovanteId = String(item.ID || '').trim();
  const alocacoes = (Array.isArray(pagamentos) ? pagamentos : [])
    .filter(pagamento => String(pagamento.comprovante_id || '').trim() === comprovanteId)
    .map(pagamento => {
      const parcela = catalogo.mapaPorParcela[String(pagamento.parcela_alvo_id || '').trim()] || null;
      const origemTipo = String(pagamento.origem_tipo || '').trim().toUpperCase();
      const origemId = String(pagamento.origem_id || '').trim();
      const origem = catalogo.origens[`${origemTipo}|${origemId}`] || {};
      return {
        pagamento_id: String(pagamento.ID || '').trim(),
        origem_tipo: origemTipo,
        origem_id: origemId,
        origem_rotulo: String(origem.origem_rotulo || origemId).trim(),
        parcela_id: String(pagamento.parcela_alvo_id || '').trim(),
        parcela_numero: Number(parcela?.parcela_numero || 1),
        parcelas_total: Number(parcela?.parcelas_total || 1),
        valor_alocado: round2Financeiro(parseNumeroBR(pagamento.valor_pago)),
        data_pagamento: formatarDataYmdFinanceiroSafe(pagamento.data_pagamento),
        forma_pagamento: String(pagamento.forma_pagamento || '').trim()
      };
    });
  const statusPersistido = String(item.status || '').trim().toUpperCase();
  const despesaLegadoId = String(item.despesa_id_confirmada || '').trim();
  const confirmadoLegado = statusPersistido === 'CONFIRMADO' && despesaLegadoId && alocacoes.length === 0;
  if (confirmadoLegado) {
    const origem = catalogo.origens[`DESPESA_GERAL|${despesaLegadoId}`] || {};
    alocacoes.push({
      pagamento_id: '',
      origem_tipo: 'DESPESA_GERAL',
      origem_id: despesaLegadoId,
      origem_rotulo: String(origem.origem_rotulo || item.descricao || despesaLegadoId).trim(),
      parcela_id: '',
      parcela_numero: 1,
      parcelas_total: 1,
      valor_alocado: round2Financeiro(parseNumeroBR(item.valor_total)),
      data_pagamento: formatarDataYmdFinanceiroSafe(item.data_pagamento),
      forma_pagamento: String(item.forma_pagamento || '').trim(),
      legado: true
    });
  }
  const valorAlocado = round2Financeiro(alocacoes.reduce((acc, alocacao) => acc + alocacao.valor_alocado, 0));
  const valorTotal = round2Financeiro(parseNumeroBR(item.valor_total));
  const saldo = round2Financeiro(Math.max(0, valorTotal - valorAlocado));
  let statusConciliacao = COMPROVANTE_STATUS_PENDENTE;
  if (statusPersistido === 'ERRO' && alocacoes.length === 0) statusConciliacao = 'ERRO';
  else if (saldo <= 0.009 && valorTotal > 0) statusConciliacao = COMPROVANTE_STATUS_CONCILIADO;
  else if (valorAlocado > 0) statusConciliacao = COMPROVANTE_STATUS_PARCIAL;

  const sugestoes = saldo > 0.009
    ? catalogo.pendentes
      .map(destino => pontuarDestinoComprovante_(item, saldo, destino))
      .filter(destino => destino.score >= 18)
      .sort((a, b) => b.score - a.score || a.label.localeCompare(b.label))
      .slice(0, 5)
    : [];
  return {
    ...item,
    status_conciliacao: statusConciliacao,
    valor_alocado: valorAlocado,
    saldo_nao_alocado: saldo,
    alocacoes,
    sugestoes
  };
}

function pontuarDestinoComprovante_(comprovante, saldo, destino) {
  let score = 0;
  const motivos = [];
  const pendente = round2Financeiro(destino.valor_pendente);
  const diferenca = Math.abs(saldo - pendente);
  if (diferenca <= 0.009) {
    score += 55;
    motivos.push('valor exato');
  } else if (pendente <= saldo + 0.009) {
    score += 38;
    motivos.push('parcela cabe no comprovante');
  } else if (saldo < pendente) {
    const proporcao = pendente > 0 ? saldo / pendente : 0;
    score += Math.round(20 + Math.min(15, proporcao * 15));
    motivos.push('possivel pagamento parcial');
  }

  const textoComprovante = `${comprovante.fornecedor || ''} ${comprovante.descricao || ''}`;
  const textoDestino = `${destino.fornecedor || ''} ${destino.origem_rotulo || ''} ${destino.observacao || ''}`;
  const similaridade = similaridadeTextoComprovante_(textoComprovante, textoDestino);
  if (similaridade >= 0.15) {
    const pontosNome = Math.round(Math.min(25, similaridade * 32));
    score += pontosNome;
    motivos.push(similaridade >= 0.55 ? 'nome muito semelhante' : 'nome relacionado');
  }

  const dataComprovante = parseDataFinanceiro(comprovante.data_pagamento || comprovante.data_competencia);
  const dataDestino = parseDataFinanceiro(destino.data_prevista || destino.data_origem);
  if (dataComprovante && dataDestino) {
    const dias = Math.round(Math.abs(dataComprovante.getTime() - dataDestino.getTime()) / 86400000);
    if (dias === 0) { score += 15; motivos.push('mesma data'); }
    else if (dias <= 3) { score += 12; motivos.push(`datas proximas (${dias}d)`); }
    else if (dias <= 7) { score += 8; motivos.push(`datas proximas (${dias}d)`); }
    else if (dias <= 30) score += 3;
  }
  if (comprovante.pago_por && destino.pago_por &&
      normalizarTextoComprovante_(comprovante.pago_por) === normalizarTextoComprovante_(destino.pago_por)) {
    score += 5;
    motivos.push('mesmo pagador');
  }
  return { ...destino, score: Math.min(100, score), motivos };
}

function normalizarTextoComprovante_(valor) {
  return String(valor || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function similaridadeTextoComprovante_(a, b) {
  const ignorar = new Set(['de', 'da', 'do', 'das', 'dos', 'e', 'pagamento', 'pix', 'banco', 'compra', 'despesa']);
  const tokens = valor => [...new Set(normalizarTextoComprovante_(valor)
    .replace(/[^a-z0-9]+/g, ' ').split(/\s+/)
    .filter(token => token.length >= 3 && !ignorar.has(token)))];
  const aa = tokens(a);
  const bb = tokens(b);
  if (aa.length === 0 || bb.length === 0) return 0;
  const setB = new Set(bb);
  const intersecao = aa.filter(token => setB.has(token)).length;
  const uniao = new Set([...aa, ...bb]).size;
  return uniao > 0 ? intersecao / uniao : 0;
}

function formatarMoedaComprovante_(valor) {
  return `R$ ${round2Financeiro(valor).toFixed(2).replace('.', ',')}`;
}

function alocarComprovanteFinanceiro(comprovanteId, payload) {
  const ambiente = payload?.ambiente || payload?._db_env || payload?.db_env || '';
  return executarComAmbienteComprovantesFinanceiros_(ambiente, () =>
    executarComLockComprovante_(() => alocarComprovanteFinanceiroAtual_(comprovanteId, payload))
  );
}

function salvarRevisaoComprovanteFinanceiro(comprovanteId, payload) {
  const ambiente = payload?.ambiente || payload?._db_env || payload?.db_env || '';
  return executarComAmbienteComprovantesFinanceiros_(ambiente, () =>
    executarComLockComprovante_(() => {
      assertCanWrite('Revisao de comprovante financeiro');
      const comprovante = obterInboxDespesaPorId_(comprovanteId);
      if (!comprovante) throw new Error('Comprovante nao encontrado.');
      if (String(comprovante.status || '').trim().toUpperCase() === 'DESCARTADO') {
        throw new Error('Comprovante descartado.');
      }
      persistirEdicaoComprovanteFinanceiro_(comprovante, payload);
      return { ok: true, comprovante: atualizarEstadoComprovanteFinanceiro_(comprovanteId) };
    })
  );
}

function alocarComprovanteFinanceiroAtual_(comprovanteId, payload) {
  assertCanWrite('Vinculo de comprovante financeiro');
  let comprovante = obterInboxDespesaPorId_(comprovanteId);
  if (!comprovante) throw new Error('Comprovante nao encontrado.');
  if (String(comprovante.status || '').toUpperCase() === 'DESCARTADO') throw new Error('Comprovante descartado.');
  comprovante = persistirEdicaoComprovanteFinanceiro_(comprovante, payload);
  const valor = round2Financeiro(parseNumeroBR(payload?.valor_alocado));
  if (valor <= 0) throw new Error('Informe um valor maior que zero para o vinculo.');
  const origemTipo = String(payload?.origem_tipo || '').trim().toUpperCase();
  const origemId = String(payload?.origem_id || '').trim();
  const parcelaId = String(payload?.parcela_id || payload?.parcela_alvo_id || '').trim();
  if (!['COMPRA', 'DESPESA_GERAL'].includes(origemTipo) || !origemId || !parcelaId) {
    throw new Error('Selecione uma compra ou despesa e uma parcela valida.');
  }
  const dataPagamento = normalizarDataFinanceiro(payload?.data_pagamento || comprovante.data_pagamento, true, 'Data de pagamento');
  const formaPagamento = validarFormaPagamentoFinanceiro(payload?.forma_pagamento || comprovante.forma_pagamento, true);
  const clientRequestIdFornecido = normalizarClientRequestIdComprovante_(payload?.client_request_id);
  const payloadPagamento = clientRequestId => ({
    parcela_alvo_id: parcelaId,
    data_pagamento: dataPagamento,
    valor_pago: valor,
    forma_pagamento: formaPagamento,
    observacao: String(payload?.observacao || `Vinculado ao comprovante ${comprovante.ID}`).trim(),
    comprovante_id: comprovante.ID,
    client_request_id: clientRequestId
  });
  if (clientRequestIdFornecido) {
    const existente = listarPagamentos(true).find(pagamento =>
      String(pagamento.client_request_id || '').trim() === clientRequestIdFornecido
    );
    if (existente) {
      const pagamento = registrarPagamento(origemTipo, origemId, {
        ...payloadPagamento(clientRequestIdFornecido),
        parcela_alvo_id: String(existente.parcela_alvo_id || parcelaId).trim()
      });
      const comprovanteAtualizado = atualizarEstadoComprovanteFinanceiro_(comprovante.ID);
      return { ok: true, pagamento, comprovante: comprovanteAtualizado };
    }
  }
  const catalogo = montarCatalogoFinanceiroComprovantes_();
  const pagamentos = listarPagamentos(true);
  const resumo = enriquecerComprovanteFinanceiro_(comprovante, pagamentos, catalogo);
  if (valor > resumo.saldo_nao_alocado + 0.009) throw new Error('Valor maior que o saldo disponivel do comprovante.');
  const destino = catalogo.mapaPorParcela[parcelaId];
  if (!destino || destino.origem_tipo !== origemTipo || destino.origem_id !== origemId) {
    throw new Error('Parcela selecionada nao pertence ao lancamento informado.');
  }
  if (valor > destino.valor_pendente + 0.009) throw new Error('Valor maior que o saldo pendente da parcela.');
  const clientRequestId = clientRequestIdFornecido ||
    gerarClientRequestIdAlocacaoComprovante_(
      comprovante.ID, origemTipo, origemId, parcelaId, resumo.saldo_nao_alocado, destino.valor_pendente, valor
    );
  const pagamento = registrarPagamento(origemTipo, origemId, payloadPagamento(clientRequestId));
  if (!pagamento || !String(pagamento.ID || '').trim()) {
    throw new Error('O pagamento nao foi persistido. Tente novamente.');
  }
  const comprovanteAtualizado = atualizarEstadoComprovanteFinanceiro_(comprovante.ID);
  return {
    ok: true,
    pagamento,
    comprovante: comprovanteAtualizado
  };
}

function registrarSaldoComprovanteComoDespesa(comprovanteId, payload) {
  const ambiente = payload?.ambiente || payload?._db_env || payload?.db_env || '';
  return executarComAmbienteComprovantesFinanceiros_(ambiente, () =>
    executarComLockComprovante_(() => registrarSaldoComprovanteComoDespesaAtual_(comprovanteId, payload))
  );
}

function registrarSaldoComprovanteComoDespesaAtual_(comprovanteId, payload) {
  assertCanWrite('Registro de comprovante como despesa simples');
  let comprovante = obterInboxDespesaPorId_(comprovanteId);
  if (!comprovante) throw new Error('Comprovante nao encontrado.');
  comprovante = persistirEdicaoComprovanteFinanceiro_(comprovante, payload);
  const baseKeyFornecida = normalizarClientRequestIdComprovante_(payload?.client_request_id);
  const valorSolicitado = round2Financeiro(parseNumeroBR(payload?.valor_alocado));
  if (baseKeyFornecida) {
    const pagamentoExistente = listarPagamentos(true).find(pagamento =>
      String(pagamento.client_request_id || '').trim() === `${baseKeyFornecida}_PAG`.slice(0, 120)
    );
    if (pagamentoExistente) {
      const despesaIdExistente = String(pagamentoExistente.origem_id || '').trim();
      const sheetDespesas = getSheet(ABA_DESPESAS_GERAIS);
      const despesaExistente = sheetDespesas
        ? rowsToObjects(sheetDespesas).find(item =>
          String(item.ID || '').trim() === despesaIdExistente && String(item.ativo).toLowerCase() === 'true'
        )
        : null;
      if (!despesaExistente || String(pagamentoExistente.origem_tipo || '').trim().toUpperCase() !== 'DESPESA_GERAL') {
        throw new Error('A operacao repetida aponta para uma despesa que nao esta mais disponivel.');
      }
      const dataPagamentoExistente = normalizarDataFinanceiro(payload?.data_pagamento || comprovante.data_pagamento, true, 'Data de pagamento');
      const formaPagamentoExistente = validarFormaPagamentoFinanceiro(payload?.forma_pagamento || comprovante.forma_pagamento, true);
      const valorRepeticao = valorSolicitado > 0
        ? valorSolicitado
        : round2Financeiro(parseNumeroBR(pagamentoExistente.valor_pago));
      const pagamento = registrarPagamento('DESPESA_GERAL', despesaIdExistente, {
        parcela_alvo_id: String(pagamentoExistente.parcela_alvo_id || '').trim(),
        data_pagamento: dataPagamentoExistente,
        valor_pago: valorRepeticao,
        forma_pagamento: formaPagamentoExistente,
        observacao: `Despesa registrada pelo comprovante ${comprovante.ID}`,
        comprovante_id: comprovante.ID,
        client_request_id: `${baseKeyFornecida}_PAG`.slice(0, 120)
      });
      const comprovanteAtualizado = atualizarEstadoComprovanteFinanceiro_(comprovante.ID);
      return {
        ok: true,
        despesa_id: despesaIdExistente,
        despesa: despesaExistente,
        pagamento,
        comprovante: comprovanteAtualizado
      };
    }
  }
  const resumo = obterComprovanteFinanceiroEnriquecido_(comprovante.ID);
  const valor = round2Financeiro(parseNumeroBR(payload?.valor_alocado || resumo.saldo_nao_alocado));
  if (valor <= 0 || valor > resumo.saldo_nao_alocado + 0.009) {
    throw new Error('Valor da despesa deve ser maior que zero e respeitar o saldo do comprovante.');
  }
  const descricao = String(payload?.descricao || comprovante.descricao || '').trim();
  if (!descricao) throw new Error('Descricao da despesa e obrigatoria.');
  const categoria = validarCategoriaDespesaFinanceiro(payload?.categoria || comprovante.categoria, true);
  const pagoPor = validarPagoPorFinanceiro(payload?.pago_por || comprovante.pago_por, false);
  const dataPagamento = normalizarDataFinanceiro(payload?.data_pagamento || comprovante.data_pagamento, true, 'Data de pagamento');
  const dataCompetencia = normalizarDataFinanceiro(payload?.data_competencia || comprovante.data_competencia || dataPagamento, true, 'Data de competencia');
  const formaPagamento = validarFormaPagamentoFinanceiro(payload?.forma_pagamento || comprovante.forma_pagamento, true);
  const baseKey = baseKeyFornecida ||
    gerarClientRequestIdAlocacaoComprovante_(
      comprovante.ID, 'DESPESA_GERAL', 'NOVA', '', resumo.saldo_nao_alocado, valor, valor
    );
  const despesa = criarDespesaGeral({
    descricao,
    categoria,
    fornecedor: String(payload?.fornecedor || comprovante.fornecedor || '').trim(),
    pago_por: pagoPor,
    valor_total: valor,
    data_competencia: dataCompetencia,
    data_vencimento: String(payload?.data_vencimento || comprovante.data_vencimento || dataPagamento).trim(),
    data_pagamento: dataPagamento,
    forma_pagamento: formaPagamento,
    parcelas: 1,
    parcelas_detalhe: [],
    fixo: false,
    observacao: String(payload?.observacao || comprovante.observacao || '').trim(),
    client_request_id: `${baseKey}_DES`.slice(0, 120)
  }, { lockJaAdquirido: true });
  const despesaId = String(despesa?.ID || '').trim();
  if (!despesaId) throw new Error('Nao foi possivel criar a despesa simples.');
  const parcela = listarParcelasFinanceirasOrigem('DESPESA_GERAL', despesaId)[0] || null;
  const pagamento = registrarPagamento('DESPESA_GERAL', despesaId, {
    parcela_alvo_id: String(parcela?.ID || '').trim(),
    data_pagamento: dataPagamento,
    valor_pago: valor,
    forma_pagamento: formaPagamento,
    observacao: `Despesa registrada pelo comprovante ${comprovante.ID}`,
    comprovante_id: comprovante.ID,
    client_request_id: `${baseKey}_PAG`.slice(0, 120)
  });
  if (!pagamento || !String(pagamento.ID || '').trim()) {
    throw new Error('A despesa foi criada, mas o pagamento nao foi persistido. Tente novamente para concluir.');
  }
  updateById(ABA_INBOX_DESPESAS, 'ID', comprovante.ID, {
    despesa_id_confirmada: comprovante.despesa_id_confirmada || despesaId,
    atualizado_em: new Date()
  }, INBOX_DESPESAS_SCHEMA);
  const comprovanteAtualizado = atualizarEstadoComprovanteFinanceiro_(comprovante.ID);
  return {
    ok: true,
    despesa_id: despesaId,
    despesa,
    pagamento,
    comprovante: comprovanteAtualizado
  };
}

function desfazerAlocacaoComprovanteFinanceiro(comprovanteId, pagamentoId, ambiente) {
  return executarComAmbienteComprovantesFinanceiros_(ambiente, () =>
    executarComLockComprovante_(() => {
      assertCanWrite('Remocao de vinculo de comprovante');
      const pagamento = listarPagamentos(true).find(item =>
        String(item.ID || '').trim() === String(pagamentoId || '').trim() &&
        String(item.comprovante_id || '').trim() === String(comprovanteId || '').trim()
      );
      if (!pagamento) throw new Error('Vinculo do comprovante nao encontrado.');
      const origemTipo = String(pagamento.origem_tipo || '').trim().toUpperCase();
      const origemId = String(pagamento.origem_id || '').trim();
      const removido = removerPagamento_(pagamento.ID, { rollbackEmFalha: true });
      if (!removido) throw new Error('Nao foi possivel remover o vinculo financeiro.');
      if (origemTipo === 'DESPESA_GERAL') {
        const sheet = getSheet(ABA_DESPESAS_GERAIS);
        const despesa = sheet ? rowsToObjects(sheet).find(item => String(item.ID || '').trim() === origemId) : null;
        const geradaPeloComprovante = String(despesa?.client_request_id || '').startsWith('CMP_') &&
          String(despesa?.client_request_id || '').includes(String(comprovanteId || '').slice(-12));
        const outrosPagamentos = listarPagamentos(true).some(item =>
          String(item.origem_tipo || '').trim().toUpperCase() === 'DESPESA_GERAL' &&
          String(item.origem_id || '').trim() === origemId
        );
        if (geradaPeloComprovante && !outrosPagamentos) {
          const despesaRemovida = deletarDespesaGeral(origemId);
          if (!despesaRemovida) {
            updateById(ABA_PAGAMENTOS, 'ID', pagamento.ID, { ativo: true }, PAGAMENTOS_SCHEMA);
            regerarParcelasFinanceirasOrigemComPagamentos(origemTipo, origemId);
            throw new Error('Nao foi possivel remover a despesa simples vinculada. O pagamento foi restaurado.');
          }
        }
      }
      const comprovanteAtualizado = atualizarEstadoComprovanteFinanceiro_(comprovanteId);
      return { ok: true, comprovante: comprovanteAtualizado };
    })
  );
}

function persistirEdicaoComprovanteFinanceiro_(comprovante, payload) {
  const dados = payload || {};
  const comprovanteId = String(comprovante.ID || '').trim();
  const valorAlocadoPagamentos = round2Financeiro(listarPagamentos(true)
    .filter(pagamento => String(pagamento.comprovante_id || '').trim() === comprovanteId)
    .reduce((acc, pagamento) => acc + round2Financeiro(parseNumeroBR(pagamento.valor_pago)), 0));
  const confirmadoLegado = String(comprovante.status || '').trim().toUpperCase() === 'CONFIRMADO' &&
    String(comprovante.despesa_id_confirmada || '').trim() && valorAlocadoPagamentos <= 0.009;
  const valorAlocado = confirmadoLegado
    ? round2Financeiro(parseNumeroBR(comprovante.valor_total))
    : valorAlocadoPagamentos;
  const valorTotal = round2Financeiro(parseNumeroBR(dados.valor_total ?? comprovante.valor_total));
  if (valorTotal <= 0) throw new Error('Valor total do comprovante deve ser maior que zero.');
  if (valorTotal + 0.009 < valorAlocado) throw new Error('Valor total menor que o valor ja alocado.');
  const statusAtual = String(comprovante.status || '').trim().toUpperCase();
  const alteracoes = {
    status: statusAtual === 'ERRO' ? COMPROVANTE_STATUS_PENDENTE : statusAtual,
    descricao: String(dados.descricao ?? comprovante.descricao ?? '').trim().slice(0, 160),
    categoria: String(dados.categoria ?? comprovante.categoria ?? '').trim(),
    fornecedor: String(dados.fornecedor ?? comprovante.fornecedor ?? '').trim().slice(0, 120),
    referencia_transacao: String(dados.referencia_transacao ?? comprovante.referencia_transacao ?? '').trim().slice(0, 160),
    pago_por: String(dados.pago_por ?? comprovante.pago_por ?? '').trim(),
    valor_total: valorTotal,
    data_competencia: String(dados.data_competencia ?? comprovante.data_competencia ?? '').trim(),
    data_vencimento: String(dados.data_vencimento ?? comprovante.data_vencimento ?? '').trim(),
    data_pagamento: String(dados.data_pagamento ?? comprovante.data_pagamento ?? '').trim(),
    forma_pagamento: String(dados.forma_pagamento ?? comprovante.forma_pagamento ?? '').trim(),
    observacao: String(dados.observacao ?? comprovante.observacao ?? '').trim().slice(0, 500),
    erro: '',
    atualizado_em: new Date()
  };
  updateById(ABA_INBOX_DESPESAS, 'ID', comprovante.ID, alteracoes, INBOX_DESPESAS_SCHEMA);
  return obterInboxDespesaPorId_(comprovante.ID);
}

function atualizarEstadoComprovanteFinanceiro_(comprovanteId) {
  const item = obterInboxDespesaPorId_(comprovanteId);
  if (!item) return null;
  const resumo = enriquecerComprovanteFinanceiro_(item, listarPagamentos(true), montarCatalogoFinanceiroComprovantes_());
  const tipos = new Set(resumo.alocacoes.map(alocacao => alocacao.origem_tipo));
  let classificacao = COMPROVANTE_CLASSIFICACAO_NAO_CLASSIFICADO;
  if (tipos.has('COMPRA') && tipos.has('DESPESA_GERAL')) classificacao = COMPROVANTE_CLASSIFICACAO_MISTO;
  else if (tipos.has('COMPRA')) classificacao = COMPROVANTE_CLASSIFICACAO_COMPRA;
  else if (tipos.has('DESPESA_GERAL')) classificacao = COMPROVANTE_CLASSIFICACAO_DESPESA;
  const agora = new Date();
  updateById(ABA_INBOX_DESPESAS, 'ID', item.ID, {
    status: resumo.status_conciliacao,
    classificacao,
    confirmado_em: item.confirmado_em || (resumo.alocacoes.length > 0 ? agora : ''),
    conciliado_em: resumo.status_conciliacao === COMPROVANTE_STATUS_CONCILIADO ? (item.conciliado_em || agora) : '',
    atualizado_em: agora,
    erro: ''
  }, INBOX_DESPESAS_SCHEMA);
  if (resumo.status_conciliacao === COMPROVANTE_STATUS_CONCILIADO &&
      typeof moverArquivoInboxDriveAposConfirmacao_ === 'function') {
    moverArquivoInboxDriveAposConfirmacao_(item);
  }
  return {
    ...resumo,
    ...(obterInboxDespesaPorId_(item.ID) || {}),
    status_conciliacao: resumo.status_conciliacao,
    classificacao,
    valor_alocado: resumo.valor_alocado,
    saldo_nao_alocado: resumo.saldo_nao_alocado
  };
}

function obterComprovanteFinanceiroEnriquecido_(comprovanteId) {
  const item = obterInboxDespesaPorId_(comprovanteId);
  if (!item) throw new Error('Comprovante nao encontrado.');
  return enriquecerComprovanteFinanceiro_(item, listarPagamentos(true), montarCatalogoFinanceiroComprovantes_());
}

function gerarClientRequestIdAlocacaoComprovante_(comprovanteId, origemTipo, origemId, parcelaId, saldo, pendente, valor) {
  const base = [comprovanteId, origemTipo, origemId, parcelaId, round2Financeiro(saldo), round2Financeiro(pendente), round2Financeiro(valor)].join('|');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.newBlob(base).getBytes())
    .map(byte => (`0${(byte & 255).toString(16)}`).slice(-2)).join('').slice(0, 24);
  return `CMP_${String(comprovanteId || '').slice(-12)}_${digest}`.slice(0, 120);
}

function normalizarClientRequestIdComprovante_(valor) {
  return String(valor || '').trim().slice(0, 112);
}

function executarComLockComprovante_(callback) {
  const lock = typeof LockService !== 'undefined' && LockService.getScriptLock ? LockService.getScriptLock() : null;
  if (lock) lock.waitLock(15000);
  try {
    return callback();
  } finally {
    if (lock) {
      try { lock.releaseLock(); } catch (error) { /* sem acao */ }
    }
  }
}

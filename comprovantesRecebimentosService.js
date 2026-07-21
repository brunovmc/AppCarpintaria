const RECEBIMENTO_STATUS_PENDENTE = 'PENDENTE';
const RECEBIMENTO_STATUS_PARCIAL = 'PARCIAL';
const RECEBIMENTO_STATUS_CONCILIADO = 'CONCILIADO';

function obterContextoComprovantesRecebimentos(statusFiltro, ambiente) {
  return executarComAmbienteInboxRecebimentos_(ambiente, () =>
    obterContextoComprovantesRecebimentosAtual_(statusFiltro)
  );
}

function obterContextoComprovantesRecebimentosAtual_(statusFiltro) {
  if (typeof assertCanRead === 'function') assertCanRead('Conciliacao de recebimentos');
  tentarReconciliarEstruturaComprovantesDriveNoAcesso_();
  const catalogo = montarCatalogoRecebimentos_();
  const pagamentos = listarPagamentos(true);
  const todos = listarInboxRecebimentosNoAmbienteAtual_('TODOS')
    .filter(item => String(item.status || '').trim().toUpperCase() !== 'DESCARTADO')
    .map(item => enriquecerComprovanteRecebimento_(item, pagamentos, catalogo));
  const filtro = String(statusFiltro || 'PENDENTES').trim().toUpperCase();
  const comprovantes = todos.filter(item => {
    const status = String(item.status_conciliacao || item.status || '').trim().toUpperCase();
    if (filtro === 'TODOS' || filtro === 'ALL') return true;
    if (filtro === 'CONCILIADOS' || filtro === 'CONCILIADO') {
      return status === RECEBIMENTO_STATUS_CONCILIADO;
    }
    return [RECEBIMENTO_STATUS_PENDENTE, RECEBIMENTO_STATUS_PARCIAL, 'ERRO'].includes(status);
  });
  const destinosMap = {};
  catalogo.pendentes.slice(0, 250).forEach(destino => { destinosMap[destino.chave] = destino; });
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
      pendentes: todos.filter(item => item.status_conciliacao === RECEBIMENTO_STATUS_PENDENTE).length,
      parciais: todos.filter(item => item.status_conciliacao === RECEBIMENTO_STATUS_PARCIAL).length,
      conciliados: todos.filter(item => item.status_conciliacao === RECEBIMENTO_STATUS_CONCILIADO).length,
      erros: todos.filter(item => item.status_conciliacao === 'ERRO').length
    }
  };
}

function montarCatalogoRecebimentos_() {
  const vendas = listarVendas(true);
  const investimentos = typeof listarInvestimentosNoAmbienteAtual_ === 'function'
    ? listarInvestimentosNoAmbienteAtual_(true)
    : [];
  const origens = {};
  vendas.forEach(venda => {
    const id = String(venda.ID || '').trim();
    if (!id) return;
    origens[`${ORIGEM_TIPO_VENDA}|${id}`] = {
      origem_tipo: ORIGEM_TIPO_VENDA,
      origem_id: id,
      origem_classe: 'Venda',
      origem_rotulo: String(venda.item || 'Venda').trim() || 'Venda',
      cliente: String(venda.cliente || '').trim(),
      referencia: String(venda.referencia_venda || '').trim(),
      recebido_por: String(venda.recebido_por || '').trim(),
      forma_pagamento: String(venda.forma_pagamento || '').trim(),
      data_origem: formatarDataYmdFinanceiroSafe(venda.data_venda || venda.criado_em),
      observacao: String(venda.observacao || '').trim()
    };
  });
  investimentos.forEach(investimento => {
    const id = String(investimento.ID || '').trim();
    if (!id) return;
    origens[`${ORIGEM_TIPO_INVESTIMENTO}|${id}`] = {
      origem_tipo: ORIGEM_TIPO_INVESTIMENTO,
      origem_id: id,
      origem_classe: 'Investimento',
      origem_rotulo: String(investimento.descricao || 'Investimento').trim() || 'Investimento',
      cliente: String(investimento.investidor || '').trim(),
      referencia: String(investimento.referencia_investimento || '').trim(),
      recebido_por: String(investimento.recebido_por || '').trim(),
      forma_pagamento: String(investimento.forma_pagamento || '').trim(),
      data_origem: formatarDataYmdFinanceiroSafe(investimento.data_investimento || investimento.criado_em),
      observacao: String(investimento.observacao || '').trim()
    };
  });
  const sheet = getSheet(ABA_PARCELAS_FINANCEIRAS);
  const parcelas = sheet ? rowsToObjects(sheet) : [];
  const todos = parcelas
    .filter(parcela =>
      String(parcela.ativo).toLowerCase() === 'true' &&
      [ORIGEM_TIPO_VENDA, ORIGEM_TIPO_INVESTIMENTO].includes(
        String(parcela.origem_tipo || '').trim().toUpperCase()
      ) &&
      String(parcela.natureza || NATUREZA_RECEBIMENTO).trim().toUpperCase() === NATUREZA_RECEBIMENTO
    )
    .map(parcela => {
      const origemId = String(parcela.origem_id || '').trim();
      const origemTipo = String(parcela.origem_tipo || '').trim().toUpperCase();
      const origem = origens[`${origemTipo}|${origemId}`];
      if (!origem) return null;
      const previsto = round2Financeiro(parseNumeroBR(parcela.valor_previsto));
      const recebido = round2Financeiro(parseNumeroBR(parcela.valor_pago));
      const pendente = round2Financeiro(Math.max(0, previsto - recebido));
      const numero = Math.max(1, Math.floor(parseNumeroBR(parcela.parcela_numero) || 1));
      const total = Math.max(numero, Math.floor(parseNumeroBR(parcela.parcelas_total) || 1));
      const dataPrevista = formatarDataYmdFinanceiroSafe(parcela.data_prevista);
      const identificacao = [origem.cliente, origem.referencia, origem.origem_rotulo]
        .filter(Boolean).join(' · ');
      return {
        ...origem,
        parcela_id: String(parcela.ID || '').trim(),
        parcela_numero: numero,
        parcelas_total: total,
        data_prevista: dataPrevista,
        valor_previsto: previsto,
        valor_recebido: recebido,
        valor_pendente: pendente,
        chave: `${origemTipo}|${origemId}|${String(parcela.ID || '').trim()}`,
        label: `${origem.origem_classe}: ${identificacao || origemId} · parcela ${numero}/${total} · ${formatarMoedaRecebimento_(pendente)}`
      };
    })
    .filter(Boolean);
  const mapaPorParcela = {};
  todos.forEach(item => { mapaPorParcela[item.parcela_id] = item; });
  const pendentes = todos.filter(item => item.valor_pendente > 0.009).sort((a, b) => {
    const da = parseDataFinanceiro(a.data_prevista)?.getTime() || Number.MAX_SAFE_INTEGER;
    const db = parseDataFinanceiro(b.data_prevista)?.getTime() || Number.MAX_SAFE_INTEGER;
    return da - db || b.valor_pendente - a.valor_pendente || a.label.localeCompare(b.label);
  });
  return { todos, pendentes, mapaPorParcela, origens };
}

function enriquecerComprovanteRecebimento_(item, pagamentos, catalogo) {
  const comprovanteId = String(item.ID || '').trim();
  const alocacoes = (Array.isArray(pagamentos) ? pagamentos : [])
    .filter(pagamento =>
      String(pagamento.comprovante_id || '').trim() === comprovanteId &&
      [ORIGEM_TIPO_VENDA, ORIGEM_TIPO_INVESTIMENTO].includes(
        String(pagamento.origem_tipo || '').trim().toUpperCase()
      )
    )
    .map(pagamento => {
      const parcela = catalogo.mapaPorParcela[String(pagamento.parcela_alvo_id || '').trim()] || null;
      const origemId = String(pagamento.origem_id || '').trim();
      const origemTipo = String(pagamento.origem_tipo || '').trim().toUpperCase();
      const origem = catalogo.origens[`${origemTipo}|${origemId}`] || {};
      return {
        pagamento_id: String(pagamento.ID || '').trim(),
        origem_tipo: origemTipo,
        origem_id: origemId,
        origem_rotulo: [origem.origem_classe, origem.cliente, origem.referencia, origem.origem_rotulo]
          .filter(Boolean).join(' · ') || origemId,
        parcela_id: String(pagamento.parcela_alvo_id || '').trim(),
        parcela_numero: Number(parcela?.parcela_numero || 1),
        parcelas_total: Number(parcela?.parcelas_total || 1),
        valor_alocado: round2Financeiro(parseNumeroBR(pagamento.valor_pago)),
        data_recebimento: formatarDataYmdFinanceiroSafe(pagamento.data_pagamento),
        forma_pagamento: String(pagamento.forma_pagamento || '').trim()
      };
    });
  const valorAlocado = round2Financeiro(alocacoes.reduce((acc, itemAlocado) =>
    acc + itemAlocado.valor_alocado, 0));
  const valorTotal = round2Financeiro(parseNumeroBR(item.valor_total));
  const saldo = round2Financeiro(Math.max(0, valorTotal - valorAlocado));
  const statusPersistido = String(item.status || '').trim().toUpperCase();
  let status = RECEBIMENTO_STATUS_PENDENTE;
  if (statusPersistido === 'ERRO' && alocacoes.length === 0) status = 'ERRO';
  else if (valorTotal > 0 && saldo <= 0.009) status = RECEBIMENTO_STATUS_CONCILIADO;
  else if (valorAlocado > 0) status = RECEBIMENTO_STATUS_PARCIAL;
  const sugestoes = saldo > 0.009
    ? catalogo.pendentes
      .map(destino => pontuarDestinoRecebimento_(item, saldo, destino))
      .filter(destino => destino.score >= 18)
      .sort((a, b) => b.score - a.score || a.label.localeCompare(b.label))
      .slice(0, 5)
    : [];
  return {
    ...item,
    status_conciliacao: status,
    valor_alocado: valorAlocado,
    saldo_nao_alocado: saldo,
    alocacoes,
    sugestoes
  };
}

function pontuarDestinoRecebimento_(comprovante, saldo, destino) {
  let score = 0;
  const motivos = [];
  const pendente = round2Financeiro(destino.valor_pendente);
  const diferenca = Math.abs(saldo - pendente);
  if (diferenca <= 0.009) { score += 50; motivos.push('valor exato'); }
  else if (pendente <= saldo + 0.009) { score += 34; motivos.push('parcela cabe no comprovante'); }
  else {
    const proporcao = pendente > 0 ? saldo / pendente : 0;
    score += Math.round(18 + Math.min(15, proporcao * 15));
    motivos.push('possivel recebimento parcial');
  }
  const referenciaComprovante = normalizarTextoRecebimento_(comprovante.referencia_transacao);
  const referenciaDestino = normalizarTextoRecebimento_(destino.referencia || destino.referencia_venda);
  if (referenciaComprovante && referenciaDestino &&
      (referenciaComprovante === referenciaDestino || referenciaComprovante.includes(referenciaDestino) || referenciaDestino.includes(referenciaComprovante))) {
    score += 35;
    motivos.push('mesma referencia');
  }
  const textoComprovante = `${comprovante.pagador_nome || ''} ${comprovante.descricao || ''}`;
  const textoDestino = `${destino.cliente || ''} ${destino.origem_rotulo || ''} ${destino.observacao || ''}`;
  const similaridade = similaridadeTextoRecebimento_(textoComprovante, textoDestino);
  if (similaridade >= 0.15) {
    score += Math.round(Math.min(28, similaridade * 36));
    motivos.push(similaridade >= 0.55 ? 'cliente muito semelhante' : 'nome relacionado');
  }
  const dataRecebimento = parseDataFinanceiro(comprovante.data_recebimento);
  const dataDestino = parseDataFinanceiro(destino.data_prevista || destino.data_origem);
  if (dataRecebimento && dataDestino) {
    const dias = Math.round(Math.abs(dataRecebimento.getTime() - dataDestino.getTime()) / 86400000);
    if (dias === 0) { score += 15; motivos.push('mesma data'); }
    else if (dias <= 3) { score += 12; motivos.push(`datas proximas (${dias}d)`); }
    else if (dias <= 7) { score += 8; motivos.push(`datas proximas (${dias}d)`); }
    else if (dias <= 30) score += 3;
  }
  if (comprovante.recebido_por && destino.recebido_por &&
      normalizarTextoRecebimento_(comprovante.recebido_por) === normalizarTextoRecebimento_(destino.recebido_por)) {
    score += 6;
    motivos.push('mesma conta de recebimento');
  }
  if (comprovante.forma_pagamento && destino.forma_pagamento &&
      normalizarTextoRecebimento_(comprovante.forma_pagamento) === normalizarTextoRecebimento_(destino.forma_pagamento)) {
    score += 4;
    motivos.push('mesma forma');
  }
  return { ...destino, score: Math.min(100, score), motivos };
}

function alocarComprovanteRecebimento(comprovanteId, payload) {
  const ambiente = payload?._db_env || payload?.ambiente || payload?.db_env || '';
  return executarComAmbienteInboxRecebimentos_(ambiente, () =>
    executarComLockRecebimento_(() => alocarComprovanteRecebimentoAtual_(comprovanteId, payload))
  );
}

function alocarComprovanteRecebimentoAtual_(comprovanteId, payload) {
  assertCanWrite('Vinculo de comprovante de recebimento');
  let comprovante = obterInboxRecebimentoPorId_(comprovanteId);
  if (!comprovante) throw new Error('Comprovante de recebimento nao encontrado.');
  if (String(comprovante.status || '').toUpperCase() === 'DESCARTADO') throw new Error('Comprovante descartado.');
  comprovante = persistirEdicaoComprovanteRecebimento_(comprovante, payload);
  const valor = round2Financeiro(parseNumeroBR(payload?.valor_alocado));
  const origemTipo = normalizarOrigemTipoFinanceiro(payload?.origem_tipo || ORIGEM_TIPO_VENDA);
  const origemId = String(payload?.origem_id || '').trim();
  const parcelaId = String(payload?.parcela_id || payload?.parcela_alvo_id || '').trim();
  if (valor <= 0) throw new Error('Informe um valor maior que zero para o vinculo.');
  if (![ORIGEM_TIPO_VENDA, ORIGEM_TIPO_INVESTIMENTO].includes(origemTipo)) {
    throw new Error('Selecione uma venda ou investimento valido.');
  }
  if (!origemId || !parcelaId) throw new Error('Selecione uma origem e uma parcela valida.');
  const clientRequestIdFornecido = normalizarClientRequestIdRecebimento_(payload?.client_request_id);
  if (clientRequestIdFornecido) {
    const existente = listarPagamentos(true).find(item =>
      String(item.client_request_id || '').trim() === clientRequestIdFornecido
    );
    if (existente) {
      if (String(existente.comprovante_id || '').trim() !== String(comprovante.ID || '').trim() ||
          String(existente.origem_tipo || '').trim().toUpperCase() !== origemTipo ||
          String(existente.origem_id || '').trim() !== origemId ||
          String(existente.parcela_alvo_id || '').trim() !== parcelaId) {
        throw new Error('Identificador da operacao ja utilizado por outro vinculo.');
      }
      const pagamentoRecuperado = registrarPagamento(origemTipo, origemId, {
        parcela_alvo_id: parcelaId,
        data_pagamento: existente.data_pagamento,
        valor_pago: existente.valor_pago,
        forma_pagamento: existente.forma_pagamento,
        observacao: existente.observacao,
        comprovante_id: comprovante.ID,
        client_request_id: clientRequestIdFornecido
      });
      return {
        ok: true,
        pagamento: pagamentoRecuperado,
        comprovante: atualizarEstadoComprovanteRecebimento_(comprovante.ID),
        reutilizado: true
      };
    }
  }
  const catalogo = montarCatalogoRecebimentos_();
  const resumo = enriquecerComprovanteRecebimento_(comprovante, listarPagamentos(true), catalogo);
  if (valor > resumo.saldo_nao_alocado + 0.009) throw new Error('Valor maior que o saldo disponivel do comprovante.');
  const destino = catalogo.mapaPorParcela[parcelaId];
  if (!destino || destino.origem_id !== origemId || destino.origem_tipo !== origemTipo) {
    throw new Error('Parcela selecionada nao pertence a origem informada.');
  }
  if (valor > destino.valor_pendente + 0.009) throw new Error('Valor maior que o saldo pendente da parcela.');
  const clientRequestId = clientRequestIdFornecido ||
    gerarClientRequestIdRecebimento_(comprovante.ID, origemId, parcelaId, resumo.saldo_nao_alocado, valor);
  const pagamento = registrarPagamento(origemTipo, origemId, {
    parcela_alvo_id: parcelaId,
    data_pagamento: payload?.data_recebimento || comprovante.data_recebimento,
    valor_pago: valor,
    forma_pagamento: payload?.forma_pagamento || comprovante.forma_pagamento,
    observacao: String(payload?.observacao || `Vinculado ao recebimento ${comprovante.ID}`).trim(),
    comprovante_id: comprovante.ID,
    client_request_id: clientRequestId
  });
  if (!pagamento || !String(pagamento.ID || '').trim()) {
    throw new Error('O recebimento nao foi persistido. Tente novamente.');
  }
  return { ok: true, pagamento, comprovante: atualizarEstadoComprovanteRecebimento_(comprovante.ID) };
}

function salvarRevisaoComprovanteRecebimento(comprovanteId, payload) {
  const ambiente = payload?._db_env || payload?.ambiente || payload?.db_env || '';
  return executarComAmbienteInboxRecebimentos_(ambiente, () =>
    executarComLockRecebimento_(() => {
      assertCanWrite('Revisao de comprovante de recebimento');
      const item = obterInboxRecebimentoPorId_(comprovanteId);
      if (!item) throw new Error('Comprovante de recebimento nao encontrado.');
      persistirEdicaoComprovanteRecebimento_(item, payload);
      return { ok: true, comprovante: atualizarEstadoComprovanteRecebimento_(item.ID) };
    })
  );
}

function desfazerAlocacaoComprovanteRecebimento(comprovanteId, pagamentoId, ambiente) {
  return executarComAmbienteInboxRecebimentos_(ambiente, () =>
    executarComLockRecebimento_(() => {
      assertCanWrite('Remocao de vinculo de recebimento');
      const pagamento = listarPagamentos(true).find(item =>
        String(item.ID || '').trim() === String(pagamentoId || '').trim()
      );
      if (!pagamento ||
          String(pagamento.comprovante_id || '').trim() !== String(comprovanteId || '').trim() ||
          ![ORIGEM_TIPO_VENDA, ORIGEM_TIPO_INVESTIMENTO].includes(
            String(pagamento.origem_tipo || '').trim().toUpperCase()
          )) {
        throw new Error('Vinculo de recebimento nao encontrado.');
      }
      if (!removerPagamento_(pagamento.ID, { rollbackEmFalha: true })) {
        throw new Error('Nao foi possivel remover o vinculo.');
      }
      return { ok: true, comprovante: atualizarEstadoComprovanteRecebimento_(comprovanteId) };
    })
  );
}

function criarVendaEAlocarComprovanteRecebimento(comprovanteId, vendaPayload, payload) {
  const ambiente = payload?._db_env || payload?.ambiente || payload?.db_env || '';
  return executarComAmbienteInboxRecebimentos_(ambiente, () =>
    executarComLockRecebimento_(() => criarVendaEAlocarComprovanteRecebimentoAtual_(
      comprovanteId, vendaPayload, payload
    ))
  );
}

function criarVendaEAlocarComprovanteRecebimentoAtual_(comprovanteId, vendaPayload, payload) {
  assertCanWrite('Criacao de venda a partir de recebimento');
  let comprovante = obterInboxRecebimentoPorId_(comprovanteId);
  if (!comprovante) throw new Error('Comprovante de recebimento nao encontrado.');
  comprovante = persistirEdicaoComprovanteRecebimento_(comprovante, payload);
  const clientRequestIdFornecido = normalizarClientRequestIdRecebimento_(payload?.client_request_id);
  if (clientRequestIdFornecido) {
    const pagamentoExistente = listarPagamentos(true).find(item =>
      String(item.client_request_id || '').trim() === clientRequestIdFornecido
    );
    if (pagamentoExistente) {
      if (String(pagamentoExistente.comprovante_id || '').trim() !== String(comprovante.ID || '').trim() ||
          String(pagamentoExistente.origem_tipo || '').trim().toUpperCase() !== 'VENDA') {
        throw new Error('Identificador da operacao ja utilizado por outro recebimento.');
      }
      const vendaExistente = listarVendas(true).find(item =>
        String(item.ID || '').trim() === String(pagamentoExistente.origem_id || '').trim()
      );
      if (!vendaExistente) throw new Error('A venda da operacao repetida nao esta mais disponivel.');
      return {
        ok: true,
        venda: vendaExistente,
        pagamento: pagamentoExistente,
        comprovante: atualizarEstadoComprovanteRecebimento_(comprovante.ID),
        reutilizado: true
      };
    }
  }
  const resumo = enriquecerComprovanteRecebimento_(comprovante, listarPagamentos(true), montarCatalogoRecebimentos_());
  const totalVendaPayload = round2Financeiro(parseNumeroBR(vendaPayload?.valor_total_venda));
  const solicitado = round2Financeiro(parseNumeroBR(payload?.valor_alocado));
  const valor = round2Financeiro(Math.min(
    solicitado > 0 ? solicitado : resumo.saldo_nao_alocado,
    resumo.saldo_nao_alocado,
    totalVendaPayload
  ));
  if (valor <= 0) throw new Error('Nao ha saldo disponivel para vincular a nova venda.');
  const venda = criarVenda(vendaPayload);
  if (!venda?.ID) throw new Error('A venda nao foi criada.');
  const clientRequestId = clientRequestIdFornecido ||
    gerarClientRequestIdRecebimento_(comprovante.ID, venda.ID, 'NOVA', resumo.saldo_nao_alocado, valor);
  try {
    const pagamento = registrarPagamento('VENDA', venda.ID, {
      distribuir_automaticamente: true,
      data_pagamento: payload?.data_recebimento || comprovante.data_recebimento,
      valor_pago: valor,
      forma_pagamento: payload?.forma_pagamento || comprovante.forma_pagamento,
      observacao: String(payload?.observacao || `Venda criada pelo recebimento ${comprovante.ID}`).trim(),
      comprovante_id: comprovante.ID,
      client_request_id: clientRequestId
    });
    if (!pagamento?.ID) throw new Error('O recebimento da nova venda nao foi persistido.');
    return {
      ok: true,
      venda: listarVendas(true).find(item => item.ID === venda.ID) || venda,
      pagamento,
      comprovante: atualizarEstadoComprovanteRecebimento_(comprovante.ID)
    };
  } catch (error) {
    const pagamentoParcial = listarPagamentos(true).find(item =>
      String(item.client_request_id || '').trim() === clientRequestId
    );
    if (pagamentoParcial) {
      try { removerPagamento_(pagamentoParcial.ID, { rollbackEmFalha: false }); } catch (rollbackError) { /* sem acao */ }
    }
    try {
      updateById(ABA_VENDAS, 'ID', venda.ID, { ativo: false }, VENDAS_SCHEMA);
      limparParcelasFinanceirasOrigem(ORIGEM_TIPO_VENDA, venda.ID);
      limparCacheVendas();
    } catch (rollbackError) {
      console.log(JSON.stringify({
        evento: 'inbox_recebimentos.rollback_venda_falhou',
        venda_id: venda.ID,
        erro: String(rollbackError)
      }));
    }
    throw error;
  }
}

function criarInvestimentoEAlocarComprovanteRecebimento(comprovanteId, investimentoPayload, payload) {
  const ambiente = payload?._db_env || payload?.ambiente || payload?.db_env || '';
  return executarComAmbienteInboxRecebimentos_(ambiente, () =>
    executarComLockRecebimento_(() => criarInvestimentoEAlocarComprovanteRecebimentoAtual_(
      comprovanteId,
      investimentoPayload,
      payload
    ))
  );
}

function criarInvestimentoEAlocarComprovanteRecebimentoAtual_(comprovanteId, investimentoPayload, payload) {
  assertCanWrite('Criacao de investimento a partir de recebimento');
  let comprovante = obterInboxRecebimentoPorId_(comprovanteId);
  if (!comprovante) throw new Error('Comprovante de recebimento nao encontrado.');
  comprovante = persistirEdicaoComprovanteRecebimento_(comprovante, payload);
  const clientRequestId = normalizarClientRequestIdRecebimento_(payload?.client_request_id) ||
    gerarClientRequestIdRecebimento_(comprovante.ID, 'INVESTIMENTO', 'NOVA', comprovante.valor_total, payload?.valor_alocado);

  const pagamentoExistente = listarPagamentos(true).find(item =>
    String(item.client_request_id || '').trim() === clientRequestId
  );
  if (pagamentoExistente) {
    if (String(pagamentoExistente.comprovante_id || '').trim() !== String(comprovante.ID || '').trim() ||
        String(pagamentoExistente.origem_tipo || '').trim().toUpperCase() !== ORIGEM_TIPO_INVESTIMENTO) {
      throw new Error('Identificador da operacao ja utilizado por outro recebimento.');
    }
    const investimentoExistente = listarInvestimentosNoAmbienteAtual_(true).find(item =>
      String(item.ID || '').trim() === String(pagamentoExistente.origem_id || '').trim()
    );
    if (!investimentoExistente) throw new Error('O investimento da operacao repetida nao esta mais disponivel.');
    return {
      ok: true,
      investimento: investimentoExistente,
      pagamento: pagamentoExistente,
      comprovante: atualizarEstadoComprovanteRecebimento_(comprovante.ID),
      reutilizado: true
    };
  }

  const resumo = enriquecerComprovanteRecebimento_(
    comprovante,
    listarPagamentos(true),
    montarCatalogoRecebimentos_()
  );
  const totalInvestimento = round2Financeiro(parseNumeroBR(investimentoPayload?.valor_total_investimento));
  const solicitado = round2Financeiro(parseNumeroBR(payload?.valor_alocado));
  const valor = round2Financeiro(Math.min(
    solicitado > 0 ? solicitado : resumo.saldo_nao_alocado,
    resumo.saldo_nao_alocado,
    totalInvestimento
  ));
  if (valor <= 0) throw new Error('Nao ha saldo disponivel para vincular ao novo investimento.');

  const investimento = criarInvestimento({
    ...(investimentoPayload || {}),
    client_request_id: clientRequestId
  }, { lockJaAdquirido: true });
  if (!investimento?.ID) throw new Error('O investimento nao foi criado.');
  try {
    const pagamento = registrarPagamento(ORIGEM_TIPO_INVESTIMENTO, investimento.ID, {
      distribuir_automaticamente: true,
      data_pagamento: payload?.data_recebimento || comprovante.data_recebimento,
      valor_pago: valor,
      forma_pagamento: payload?.forma_pagamento || comprovante.forma_pagamento,
      observacao: String(payload?.observacao || `Investimento criado pelo recebimento ${comprovante.ID}`).trim(),
      comprovante_id: comprovante.ID,
      client_request_id: clientRequestId
    });
    if (!pagamento?.ID) throw new Error('O recebimento do investimento nao foi persistido.');
    return {
      ok: true,
      investimento: listarInvestimentosNoAmbienteAtual_(true).find(item => item.ID === investimento.ID) || investimento,
      pagamento,
      comprovante: atualizarEstadoComprovanteRecebimento_(comprovante.ID)
    };
  } catch (error) {
    const pagamentoParcial = listarPagamentos(true).find(item =>
      String(item.client_request_id || '').trim() === clientRequestId
    );
    if (pagamentoParcial) {
      try { removerPagamento_(pagamentoParcial.ID, { rollbackEmFalha: false }); } catch (_) { /* sem acao */ }
    }
    try {
      updateById(ABA_INVESTIMENTOS, 'ID', investimento.ID, { ativo: false }, INVESTIMENTOS_SCHEMA);
      limparParcelasFinanceirasOrigem(ORIGEM_TIPO_INVESTIMENTO, investimento.ID);
      limparCacheInvestimentos();
    } catch (_) {
      // A falha original e mais relevante; a rotina de integridade pode concluir o reparo.
    }
    throw error;
  }
}

function persistirEdicaoComprovanteRecebimento_(comprovante, payload) {
  const dados = payload || {};
  const id = String(comprovante.ID || '').trim();
  const valorAlocado = round2Financeiro(listarPagamentos(true)
    .filter(item => String(item.comprovante_id || '').trim() === id)
    .reduce((acc, item) => acc + round2Financeiro(parseNumeroBR(item.valor_pago)), 0));
  const valorTotal = round2Financeiro(parseNumeroBR(dados.valor_total ?? comprovante.valor_total));
  if (valorTotal <= 0) throw new Error('Valor total do comprovante deve ser maior que zero.');
  if (valorTotal + 0.009 < valorAlocado) throw new Error('Valor total menor que o valor ja vinculado.');
  const validacoes = obterValidacoes();
  const recebidoPorRaw = String(dados.recebido_por ?? comprovante.recebido_por ?? '').trim();
  const formaRaw = String(dados.forma_pagamento ?? comprovante.forma_pagamento ?? '').trim();
  const recebidoPor = normalizarValorListaRecebimento_(recebidoPorRaw, validacoes?.pagosPor);
  const formaPagamento = normalizarValorListaRecebimento_(formaRaw, validacoes?.formasPagamento);
  if (recebidoPorRaw && !recebidoPor) throw new Error('Selecione uma opcao valida em Recebido por.');
  if (formaRaw && !formaPagamento) throw new Error('Selecione uma forma de recebimento valida.');
  const dataRecebimentoRaw = String(dados.data_recebimento ?? comprovante.data_recebimento ?? '').trim();
  const dataRecebimento = dataRecebimentoRaw
    ? normalizarDataFinanceiro(dataRecebimentoRaw, false, 'Data de recebimento')
    : '';
  const statusAtual = String(comprovante.status || '').trim().toUpperCase();
  updateById(ABA_INBOX_RECEBIMENTOS, 'ID', id, {
    status: statusAtual === 'ERRO' ? RECEBIMENTO_STATUS_PENDENTE : statusAtual,
    referencia_transacao: String(dados.referencia_transacao ?? comprovante.referencia_transacao ?? '').trim().slice(0, 160),
    pagador_nome: String(dados.pagador_nome ?? comprovante.pagador_nome ?? '').trim().slice(0, 160),
    banco_pagador: String(dados.banco_pagador ?? comprovante.banco_pagador ?? '').trim().slice(0, 120),
    recebido_por: recebidoPor,
    valor_total: valorTotal,
    data_recebimento: dataRecebimento,
    forma_pagamento: formaPagamento,
    descricao: String(dados.descricao ?? comprovante.descricao ?? '').trim().slice(0, 200),
    observacao: String(dados.observacao ?? comprovante.observacao ?? '').trim().slice(0, 500),
    erro: '',
    atualizado_em: new Date()
  }, INBOX_RECEBIMENTOS_SCHEMA);
  return obterInboxRecebimentoPorId_(id);
}

function atualizarEstadoComprovanteRecebimento_(comprovanteId) {
  const item = obterInboxRecebimentoPorId_(comprovanteId);
  if (!item) return null;
  const resumo = enriquecerComprovanteRecebimento_(
    item, listarPagamentos(true), montarCatalogoRecebimentos_()
  );
  const agora = new Date();
  updateById(ABA_INBOX_RECEBIMENTOS, 'ID', item.ID, {
    status: resumo.status_conciliacao,
    confirmado_em: item.confirmado_em || (resumo.alocacoes.length > 0 ? agora : ''),
    conciliado_em: resumo.status_conciliacao === RECEBIMENTO_STATUS_CONCILIADO
      ? (item.conciliado_em || agora)
      : '',
    atualizado_em: agora,
    erro: ''
  }, INBOX_RECEBIMENTOS_SCHEMA);
  if (resumo.status_conciliacao === RECEBIMENTO_STATUS_CONCILIADO &&
      typeof moverArquivoInboxRecebimentoAposConfirmacao_ === 'function') {
    moverArquivoInboxRecebimentoAposConfirmacao_(item);
  }
  return {
    ...resumo,
    ...(obterInboxRecebimentoPorId_(item.ID) || {}),
    status_conciliacao: resumo.status_conciliacao,
    valor_alocado: resumo.valor_alocado,
    saldo_nao_alocado: resumo.saldo_nao_alocado
  };
}

function similaridadeTextoRecebimento_(a, b) {
  const ignorar = new Set(['de', 'da', 'do', 'das', 'dos', 'e', 'pix', 'banco', 'recebimento', 'venda']);
  const tokens = valor => [...new Set(normalizarTextoRecebimento_(valor)
    .replace(/[^a-z0-9]+/g, ' ').split(/\s+/)
    .filter(token => token.length >= 3 && !ignorar.has(token)))];
  const aa = tokens(a);
  const bb = tokens(b);
  if (!aa.length || !bb.length) return 0;
  const setB = new Set(bb);
  const intersecao = aa.filter(token => setB.has(token)).length;
  return intersecao / new Set([...aa, ...bb]).size;
}

function gerarClientRequestIdRecebimento_(comprovanteId, origemId, parcelaId, saldo, valor) {
  const base = [comprovanteId, origemId, parcelaId, round2Financeiro(saldo), round2Financeiro(valor)].join('|');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.newBlob(base).getBytes())
    .map(byte => (`0${(byte & 255).toString(16)}`).slice(-2)).join('').slice(0, 24);
  return `REC_${String(comprovanteId || '').slice(-12)}_${digest}`.slice(0, 120);
}

function normalizarClientRequestIdRecebimento_(valor) {
  return String(valor || '').trim().slice(0, 112);
}

function formatarMoedaRecebimento_(valor) {
  return `R$ ${round2Financeiro(valor).toFixed(2).replace('.', ',')}`;
}

function executarComLockRecebimento_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return callback(); } finally { lock.releaseLock(); }
}

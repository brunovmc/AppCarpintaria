const ABA_DESPESAS_GERAIS = 'DESPESAS_GERAIS';
const ABA_PAGAMENTOS = 'PAGAMENTOS';
const ABA_COMPRAS_FINANCEIRO = 'COMPRAS';

const DESPESAS_GERAIS_CACHE_SCOPE = 'DESPESAS_GERAIS_LISTA_ATIVAS';
const DESPESAS_GERAIS_CACHE_TTL_SEC = 90;

const PAGAMENTOS_CACHE_SCOPE = 'PAGAMENTOS_LISTA_ATIVOS';
const PAGAMENTOS_CACHE_TTL_SEC = 90;

const DASHBOARD_FINANCEIRO_CACHE_SCOPE = 'DASHBOARD_FINANCEIRO';
const DASHBOARD_FINANCEIRO_CACHE_TTL_SEC = 120;

const ORIGEM_TIPO_COMPRA = 'COMPRA';
const ORIGEM_TIPO_DESPESA = 'DESPESA_GERAL';

const DESPESAS_GERAIS_SCHEMA = [
  'ID',
  'descricao',
  'categoria',
  'fornecedor',
  'pago_por',
  'valor_total',
  'ativo',
  'criado_em',
  'data_competencia',
  'data_vencimento',
  'data_pagamento',
  'forma_pagamento',
  'parcelas',
  'fixo',
  'origem_fixo_id',
  'observacao'
];

const PAGAMENTOS_SCHEMA = [
  'ID',
  'origem_tipo',
  'origem_id',
  'data_pagamento',
  'valor_pago',
  'forma_pagamento',
  'observacao',
  'ativo',
  'criado_em'
];

function round2Financeiro(valor) {
  const n = Number(valor || 0);
  if (!isFinite(n)) return 0;
  return Number(n.toFixed(2));
}

function parseBooleanFinanceiro(valor) {
  const s = String(valor ?? '').trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'sim' || s === 'yes';
}

function adicionarMesesComAjusteFinanceiro(dataBase, quantidadeMeses) {
  const base = parseDataFinanceiro(dataBase);
  if (!base || !isFinite(quantidadeMeses)) return null;

  const meses = Number(quantidadeMeses);
  const anoBase = base.getFullYear();
  const mesBase = base.getMonth();
  const diaBase = base.getDate();
  const alvo = new Date(anoBase, mesBase + meses, 1);
  const ultimoDiaMesAlvo = new Date(alvo.getFullYear(), alvo.getMonth() + 1, 0).getDate();
  const diaAjustado = Math.min(diaBase, ultimoDiaMesAlvo);
  return new Date(alvo.getFullYear(), alvo.getMonth(), diaAjustado);
}

function formatarDataYmdFinanceiro(data) {
  return Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatarDataYmdHmFinanceiro(data) {
  return Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function formatarDataYmdFinanceiroSafe(valor) {
  const d = parseDataFinanceiro(valor);
  return d ? formatarDataYmdFinanceiro(d) : '';
}

function formatarDataYmdHmFinanceiroSafe(valor) {
  const d = parseDataFinanceiro(valor);
  return d ? formatarDataYmdHmFinanceiro(d) : '';
}

function parseDataFinanceiro(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor;
  }

  if (typeof valor === 'number' && isFinite(valor)) {
    const d = new Date(valor);
    return isNaN(d.getTime()) ? null : d;
  }

  const s = String(valor || '').trim();
  if (!s) return null;

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const d = new Date(`${s}T00:00:00`);
    return isNaN(d.getTime()) ? null : d;
  }

  if (/^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}$/.test(s)) {
    const d = new Date(s.replace(/\s+/, 'T'));
    return isNaN(d.getTime()) ? null : d;
  }

  const br = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (br) {
    const d = new Date(`${br[3]}-${br[2]}-${br[1]}T00:00:00`);
    return isNaN(d.getTime()) ? null : d;
  }

  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function normalizarDataFinanceiro(valor, obrigatoria, nomeCampo) {
  const campo = String(nomeCampo || 'Data');
  const bruto = String(valor || '').trim();
  if (!bruto) {
    if (obrigatoria) {
      throw new Error(`${campo} e obrigatoria.`);
    }
    return '';
  }

  const d = parseDataFinanceiro(bruto);
  if (!d) {
    throw new Error(`${campo} invalida.`);
  }
  return formatarDataYmdFinanceiro(d);
}

function inicioDoDiaFinanceiro(data) {
  const d = parseDataFinanceiro(data);
  if (!d) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function getMesReferenciaFinanceiro(referenciaYm) {
  const s = String(referenciaYm || '').trim();
  if (/^\d{4}-\d{2}$/.test(s)) {
    return s;
  }
  const hoje = new Date();
  return `${hoje.getFullYear()}-${String(hoje.getMonth() + 1).padStart(2, '0')}`;
}

function getIntervaloMesFinanceiro(referenciaYm) {
  const ref = getMesReferenciaFinanceiro(referenciaYm);
  const ano = Number(ref.slice(0, 4));
  const mes = Number(ref.slice(5, 7));
  const inicio = new Date(ano, mes - 1, 1);
  const fim = new Date(ano, mes, 1);
  return { ref, inicio, fim };
}

function getFormasPagamentoValidasFinanceiro() {
  const validacoes = obterValidacoes();
  const lista = Array.isArray(validacoes?.formasPagamento) ? validacoes.formasPagamento : [];
  const unicos = [];
  lista.forEach(v => {
    const label = String(v || '').trim();
    if (!label) return;
    if (!unicos.some(x => x.toUpperCase() === label.toUpperCase())) {
      unicos.push(label);
    }
  });
  return unicos;
}

function getCategoriasDespesasValidasFinanceiro() {
  const validacoes = obterValidacoes();
  const lista = Array.isArray(validacoes?.categoriasDespesas) ? validacoes.categoriasDespesas : [];
  const unicos = [];
  lista.forEach(v => {
    const label = String(v || '').trim();
    if (!label) return;
    if (!unicos.some(x => x.toUpperCase() === label.toUpperCase())) {
      unicos.push(label);
    }
  });
  return unicos;
}

function getPagosPorValidosFinanceiro() {
  const validacoes = obterValidacoes();
  const lista = Array.isArray(validacoes?.pagosPor) ? validacoes.pagosPor : [];
  const unicos = [];
  lista.forEach(v => {
    const label = String(v || '').trim();
    if (!label) return;
    if (!unicos.some(x => x.toUpperCase() === label.toUpperCase())) {
      unicos.push(label);
    }
  });
  return unicos;
}

function validarFormaPagamentoFinanceiro(formaPagamento, obrigatoria) {
  const valor = String(formaPagamento || '').trim();
  if (!valor) {
    if (obrigatoria) {
      throw new Error('Forma de pagamento e obrigatoria.');
    }
    return '';
  }

  const formasValidas = getFormasPagamentoValidasFinanceiro();
  if (formasValidas.length === 0) {
    return valor;
  }

  const match = formasValidas.find(v => v.toUpperCase() === valor.toUpperCase());
  if (!match) {
    throw new Error('Forma de pagamento invalida.');
  }
  return match;
}

function normalizarFormaPagamentoFinanceiro(formaPagamento, obrigatoria) {
  return validarFormaPagamentoFinanceiro(formaPagamento, obrigatoria);
}

function normalizarTextoSemAcentoFinanceiro(valor) {
  return String(valor || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
}

function formaAceitaParcelasFinanceiro(formaPagamento) {
  const forma = normalizarTextoSemAcentoFinanceiro(formaPagamento);
  return forma === 'CREDITO' || forma === 'PIX PARCELADO' || forma === 'DINHEIRO';
}

function normalizarParcelasFinanceiro(parcelas, formaPagamento) {
  const aceitaParcelas = formaAceitaParcelasFinanceiro(formaPagamento);
  const bruto = String(parcelas ?? '').trim();
  if (!bruto) {
    return 1;
  }

  const numero = Number(bruto.replace(',', '.'));
  const inteiro = Number.isFinite(numero) ? Math.floor(numero) : NaN;
  if (!Number.isFinite(inteiro) || inteiro < 1) {
    throw new Error('Parcelas deve ser um numero inteiro maior ou igual a 1.');
  }

  if (!aceitaParcelas) {
    return 1;
  }
  return inteiro;
}

function validarCategoriaDespesaFinanceiro(categoria, obrigatoria) {
  const valor = String(categoria || '').trim();
  if (!valor) {
    if (obrigatoria) {
      throw new Error('Categoria da despesa e obrigatoria.');
    }
    return '';
  }

  const categoriasValidas = getCategoriasDespesasValidasFinanceiro();
  if (categoriasValidas.length === 0) {
    return valor;
  }

  const match = categoriasValidas.find(v => v.toUpperCase() === valor.toUpperCase());
  if (!match) {
    throw new Error('Categoria da despesa invalida.');
  }
  return match;
}

function validarPagoPorFinanceiro(pagoPor, obrigatorio) {
  const valor = String(pagoPor || '').trim();
  if (!valor) {
    if (obrigatorio) {
      throw new Error('Pago por e obrigatorio.');
    }
    return '';
  }

  const opcoes = getPagosPorValidosFinanceiro();
  if (opcoes.length === 0) {
    return valor;
  }

  const match = opcoes.find(v => v.toUpperCase() === valor.toUpperCase());
  if (!match) {
    throw new Error('Pago por invalido.');
  }
  return match;
}

function aplicarRateioPagoPorFinanceiro(acumulador, pagoPor, valor) {
  const alvo = acumulador || { bruno: 0, zizu: 0 };
  const valorSeguro = round2Financeiro(parseNumeroBR(valor));
  if (valorSeguro <= 0) return alvo;

  const pagoPorNormalizado = normalizarTextoSemAcentoFinanceiro(pagoPor);
  if (pagoPorNormalizado === 'AMBOS') {
    const metade = round2Financeiro(valorSeguro / 2);
    alvo.bruno = round2Financeiro(alvo.bruno + metade);
    alvo.zizu = round2Financeiro(alvo.zizu + metade);
    return alvo;
  }

  if (pagoPorNormalizado === 'BRUNO') {
    alvo.bruno = round2Financeiro(alvo.bruno + valorSeguro);
    return alvo;
  }

  if (pagoPorNormalizado === 'ZIZU') {
    alvo.zizu = round2Financeiro(alvo.zizu + valorSeguro);
    return alvo;
  }

  return alvo;
}

function limparCacheDespesasGerais() {
  return appCacheRemove(DESPESAS_GERAIS_CACHE_SCOPE);
}

function lerCacheDespesasGerais() {
  return appCacheGetJson(DESPESAS_GERAIS_CACHE_SCOPE);
}

function salvarCacheDespesasGerais(lista) {
  return appCachePutJson(
    DESPESAS_GERAIS_CACHE_SCOPE,
    Array.isArray(lista) ? lista : [],
    DESPESAS_GERAIS_CACHE_TTL_SEC
  );
}

function limparCachePagamentos() {
  return appCacheRemove(PAGAMENTOS_CACHE_SCOPE);
}

function lerCachePagamentos() {
  return appCacheGetJson(PAGAMENTOS_CACHE_SCOPE);
}

function salvarCachePagamentos(lista) {
  return appCachePutJson(
    PAGAMENTOS_CACHE_SCOPE,
    Array.isArray(lista) ? lista : [],
    PAGAMENTOS_CACHE_TTL_SEC
  );
}

function limparCacheDashboardFinanceiro() {
  return appCacheRemove(DASHBOARD_FINANCEIRO_CACHE_SCOPE);
}

function lerCacheDashboardFinanceiro() {
  return appCacheGetJson(DASHBOARD_FINANCEIRO_CACHE_SCOPE);
}

function salvarCacheDashboardFinanceiro(payload) {
  return appCachePutJson(
    DASHBOARD_FINANCEIRO_CACHE_SCOPE,
    payload || {},
    DASHBOARD_FINANCEIRO_CACHE_TTL_SEC
  );
}

function normalizarOrigemTipoFinanceiro(origemTipo) {
  const tipo = String(origemTipo || '').trim().toUpperCase();
  if (tipo === ORIGEM_TIPO_COMPRA || tipo === ORIGEM_TIPO_DESPESA) {
    return tipo;
  }
  throw new Error('Origem de pagamento invalida.');
}

function getTotalPrevistoCompraFinanceiro(compra) {
  const quantidade = parseNumeroBR(compra?.quantidade);
  const valorUnit = parseNumeroBR(compra?.valor_unit);
  return round2Financeiro(Math.max(0, quantidade * valorUnit));
}

function getTotalPrevistoDespesaFinanceiro(despesa) {
  return round2Financeiro(Math.max(0, parseNumeroBR(despesa?.valor_total)));
}

function getStatusPagamentoFinanceiro(totalPrevisto, totalPago) {
  const previsto = round2Financeiro(totalPrevisto);
  const pago = round2Financeiro(totalPago);
  if (pago <= 0) return 'PENDENTE';
  if (pago + 0.009 >= previsto) return 'PAGO';
  return 'PARCIAL';
}

function mapearPagamentosPorOrigemFinanceiro(listaPagamentos) {
  const mapa = {};
  (Array.isArray(listaPagamentos) ? listaPagamentos : []).forEach(p => {
    let tipo = '';
    try {
      tipo = normalizarOrigemTipoFinanceiro(p.origem_tipo);
    } catch (error) {
      return;
    }
    const origemId = String(p.origem_id || '').trim();
    if (!origemId) return;

    const chave = `${tipo}|${origemId}`;
    if (!mapa[chave]) {
      mapa[chave] = {
        total_pago: 0,
        quantidade_pagamentos: 0,
        data_ultimo_pagamento: '',
        forma_ultimo_pagamento: ''
      };
    }

    const valor = round2Financeiro(parseNumeroBR(p.valor_pago));
    mapa[chave].total_pago = round2Financeiro(mapa[chave].total_pago + valor);
    mapa[chave].quantidade_pagamentos += 1;

    const dataAtual = parseDataFinanceiro(p.data_pagamento);
    const dataUltima = parseDataFinanceiro(mapa[chave].data_ultimo_pagamento);
    if (dataAtual && (!dataUltima || dataAtual.getTime() >= dataUltima.getTime())) {
      mapa[chave].data_ultimo_pagamento = formatarDataYmdFinanceiro(dataAtual);
      mapa[chave].forma_ultimo_pagamento = String(p.forma_pagamento || '').trim();
    }
  });
  return mapa;
}

function enriquecerListaComPagamentosFinanceiro(listaItens, origemTipo, getTotalPrevistoFn) {
  const tipo = normalizarOrigemTipoFinanceiro(origemTipo);
  const pagamentos = listarPagamentos();
  const mapa = mapearPagamentosPorOrigemFinanceiro(pagamentos);

  return (Array.isArray(listaItens) ? listaItens : []).map(item => {
    const id = String(item?.ID || '').trim();
    const chave = `${tipo}|${id}`;
    const resumo = mapa[chave] || {};

    const totalPrevisto = round2Financeiro(getTotalPrevistoFn(item));
    const totalPago = round2Financeiro(resumo.total_pago || 0);
    const totalPendente = round2Financeiro(Math.max(0, totalPrevisto - totalPago));
    const statusPagamento = getStatusPagamentoFinanceiro(totalPrevisto, totalPago);

    return {
      ...item,
      total_previsto: totalPrevisto,
      total_pago: totalPago,
      total_pendente: totalPendente,
      status_pagamento: statusPagamento,
      quantidade_pagamentos: Number(resumo.quantidade_pagamentos || 0),
      data_ultimo_pagamento: resumo.data_ultimo_pagamento || '',
      forma_ultimo_pagamento: resumo.forma_ultimo_pagamento || ''
    };
  });
}

function enriquecerComprasComResumoPagamento(listaCompras) {
  return enriquecerListaComPagamentosFinanceiro(
    listaCompras,
    ORIGEM_TIPO_COMPRA,
    getTotalPrevistoCompraFinanceiro
  );
}

function enriquecerDespesasComResumoPagamento(listaDespesas) {
  return enriquecerListaComPagamentosFinanceiro(
    listaDespesas,
    ORIGEM_TIPO_DESPESA,
    getTotalPrevistoDespesaFinanceiro
  );
}

function listarPagamentos(forcarRecarregar) {
  if (!forcarRecarregar) {
    const cached = lerCachePagamentos();
    if (Array.isArray(cached)) {
      return cached;
    }
  }

  const sheet = getSheet(ABA_PAGAMENTOS);
  if (!sheet) {
    salvarCachePagamentos([]);
    return [];
  }

  const rows = rowsToObjects(sheet);
  const lista = rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .map(i => ({
      ...i,
      valor_pago: round2Financeiro(parseNumeroBR(i.valor_pago)),
      data_pagamento: i.data_pagamento
        ? formatarDataYmdFinanceiroSafe(i.data_pagamento)
        : '',
      criado_em: i.criado_em
        ? formatarDataYmdHmFinanceiroSafe(i.criado_em)
        : ''
    }))
    .sort((a, b) => {
      const da = parseDataFinanceiro(a.data_pagamento)?.getTime() || 0;
      const db = parseDataFinanceiro(b.data_pagamento)?.getTime() || 0;
      if (da !== db) return db - da;
      const ca = parseDataFinanceiro(a.criado_em)?.getTime() || 0;
      const cb = parseDataFinanceiro(b.criado_em)?.getTime() || 0;
      return cb - ca;
    });

  salvarCachePagamentos(lista);
  return lista;
}

function recarregarCachePagamentos() {
  limparCachePagamentos();
  const dados = listarPagamentos(true);
  return {
    ok: true,
    scope: PAGAMENTOS_CACHE_SCOPE,
    ttl_segundos: PAGAMENTOS_CACHE_TTL_SEC,
    total_itens: Array.isArray(dados) ? dados.length : 0
  };
}

function normalizarPayloadDespesaGeral(payload) {
  const dados = { ...(payload || {}) };
  const descricao = String(dados.descricao || '').trim();
  if (!descricao) {
    throw new Error('Descricao da despesa e obrigatoria.');
  }

  const categoria = validarCategoriaDespesaFinanceiro(dados.categoria, true);
  const valorTotal = round2Financeiro(parseNumeroBR(dados.valor_total));
  if (valorTotal <= 0) {
    throw new Error('Valor total da despesa deve ser maior que zero.');
  }

  const dataCompetencia = normalizarDataFinanceiro(
    dados.data_competencia || new Date(),
    true,
    'Data de competencia'
  );
  const dataVencimento = normalizarDataFinanceiro(dados.data_vencimento, false, 'Data de vencimento');
  const dataPagamento = normalizarDataFinanceiro(dados.data_pagamento, false, 'Data de pagamento');
  const formaPagamento = normalizarFormaPagamentoFinanceiro(dados.forma_pagamento, false);
  const parcelas = normalizarParcelasFinanceiro(dados.parcelas, formaPagamento);
  const fixo = parseBooleanFinanceiro(dados.fixo);

  return {
    descricao,
    categoria,
    fornecedor: String(dados.fornecedor || '').trim(),
    pago_por: validarPagoPorFinanceiro(dados.pago_por, false),
    valor_total: valorTotal,
    data_competencia: dataCompetencia,
    data_vencimento: dataVencimento,
    data_pagamento: dataPagamento,
    forma_pagamento: formaPagamento,
    parcelas,
    fixo,
    observacao: String(dados.observacao || '').trim()
  };
}

function gerarRecorrenciaDespesasFixasFinanceiro() {
  const sheet = getSheet(ABA_DESPESAS_GERAIS);
  if (!sheet) {
    return { ok: true, criados: 0 };
  }

  ensureSchema(sheet, DESPESAS_GERAIS_SCHEMA);

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
  } catch (error) {
    return { ok: false, criados: 0, erro: error.message };
  }

  try {
    const ativasFixas = rowsToObjects(sheet).filter(i => {
      return String(i.ativo).toLowerCase() === 'true' && parseBooleanFinanceiro(i.fixo);
    });
    if (ativasFixas.length === 0) {
      return { ok: true, criados: 0 };
    }

    const hoje = new Date();
    const primeiroDiaMesAtual = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
    const mesAlvoData = primeiroDiaMesAtual;
    const mesAlvoYm = gerarChaveMesFinanceiro(mesAlvoData);
    if (!mesAlvoYm) {
      return { ok: true, criados: 0 };
    }

    const grupos = {};
    ativasFixas.forEach(item => {
      const rootId = String(item.origem_fixo_id || item.ID || '').trim();
      if (!rootId) return;

      const dataCompetencia = parseDataFinanceiro(item.data_competencia)
        || parseDataFinanceiro(item.data_vencimento)
        || parseDataFinanceiro(item.criado_em);
      if (!dataCompetencia) return;

      const ym = gerarChaveMesFinanceiro(dataCompetencia);
      if (!ym) return;

      if (!grupos[rootId]) {
        grupos[rootId] = [];
      }
      grupos[rootId].push({ item, dataCompetencia, ym });
    });

    let criados = 0;
    Object.keys(grupos).forEach(rootId => {
      const grupo = grupos[rootId];
      if (!Array.isArray(grupo) || grupo.length === 0) return;

      grupo.sort((a, b) => a.dataCompetencia.getTime() - b.dataCompetencia.getTime());
      const porYm = {};
      grupo.forEach(g => {
        porYm[g.ym] = g;
      });

      let atual = grupo[grupo.length - 1];
      let guard = 0;
      while (atual && atual.ym < mesAlvoYm && guard < 60) {
        guard += 1;
        const proximaCompetencia = adicionarMesesComAjusteFinanceiro(atual.dataCompetencia, 1);
        const proximoYm = gerarChaveMesFinanceiro(proximaCompetencia);
        if (!proximaCompetencia || !proximoYm) break;

        if (porYm[proximoYm]) {
          atual = porYm[proximoYm];
          continue;
        }

        const vencimentoBase = parseDataFinanceiro(atual.item.data_vencimento);
        const proximoVencimento = vencimentoBase
          ? adicionarMesesComAjusteFinanceiro(vencimentoBase, 1)
          : null;

        const novo = {
          ID: gerarId('DES'),
          descricao: String(atual.item.descricao || '').trim(),
          categoria: String(atual.item.categoria || '').trim(),
          fornecedor: String(atual.item.fornecedor || '').trim(),
          pago_por: String(atual.item.pago_por || '').trim(),
          valor_total: round2Financeiro(parseNumeroBR(atual.item.valor_total)),
          ativo: true,
          criado_em: new Date(),
          data_competencia: formatarDataYmdFinanceiro(proximaCompetencia),
          data_vencimento: proximoVencimento ? formatarDataYmdFinanceiro(proximoVencimento) : '',
          data_pagamento: '',
          forma_pagamento: String(atual.item.forma_pagamento || '').trim(),
          parcelas: normalizarParcelasFinanceiro(atual.item.parcelas, atual.item.forma_pagamento),
          fixo: true,
          origem_fixo_id: rootId,
          observacao: String(atual.item.observacao || '').trim()
        };

        const ok = insert(ABA_DESPESAS_GERAIS, novo, DESPESAS_GERAIS_SCHEMA);
        if (!ok) break;

        criados += 1;
        const novoRef = {
          item: { ...atual.item, ...novo },
          dataCompetencia: proximaCompetencia,
          ym: proximoYm
        };
        porYm[proximoYm] = novoRef;
        atual = novoRef;
      }
    });

    return { ok: true, criados };
  } finally {
    try {
      lock.releaseLock();
    } catch (error) {
      // sem acao
    }
  }
}

function listarDespesasGerais(forcarRecarregar) {
  if (!forcarRecarregar) {
    const cached = lerCacheDespesasGerais();
    if (Array.isArray(cached)) {
      return cached;
    }
  }

  const sheet = getSheet(ABA_DESPESAS_GERAIS);
  if (!sheet) {
    salvarCacheDespesasGerais([]);
    return [];
  }

  gerarRecorrenciaDespesasFixasFinanceiro();
  const rows = rowsToObjects(sheet);
  const base = rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .map(i => ({
      ...i,
      valor_total: round2Financeiro(parseNumeroBR(i.valor_total)),
      fixo: parseBooleanFinanceiro(i.fixo),
      origem_fixo_id: String(i.origem_fixo_id || '').trim(),
      criado_em: i.criado_em
        ? formatarDataYmdHmFinanceiroSafe(i.criado_em)
        : '',
      data_competencia: i.data_competencia
        ? formatarDataYmdFinanceiroSafe(i.data_competencia)
        : '',
      data_vencimento: i.data_vencimento
        ? formatarDataYmdFinanceiroSafe(i.data_vencimento)
        : '',
      data_pagamento: i.data_pagamento
        ? formatarDataYmdFinanceiroSafe(i.data_pagamento)
        : ''
    }));

  const lista = enriquecerDespesasComResumoPagamento(base)
    .sort((a, b) => {
      const ca = parseDataFinanceiro(a.criado_em)?.getTime() || 0;
      const cb = parseDataFinanceiro(b.criado_em)?.getTime() || 0;
      if (ca !== cb) return cb - ca;
      return String(b.ID || '').localeCompare(String(a.ID || ''));
    });

  salvarCacheDespesasGerais(lista);
  return lista;
}

function recarregarCacheDespesasGerais() {
  limparCacheDespesasGerais();
  const dados = listarDespesasGerais(true);
  return {
    ok: true,
    scope: DESPESAS_GERAIS_CACHE_SCOPE,
    ttl_segundos: DESPESAS_GERAIS_CACHE_TTL_SEC,
    total_itens: Array.isArray(dados) ? dados.length : 0
  };
}

function criarDespesaGeral(payload) {
  const dados = normalizarPayloadDespesaGeral(payload);
  const novoId = gerarId('DES');
  const novo = {
    ...dados,
    ID: novoId,
    ativo: true,
    criado_em: new Date(),
    origem_fixo_id: dados.fixo ? novoId : ''
  };

  const ok = insert(ABA_DESPESAS_GERAIS, novo, DESPESAS_GERAIS_SCHEMA);
  if (!ok) return null;
  return listarDespesasGerais(true).find(i => i.ID === novo.ID) || null;
}

function atualizarDespesaGeral(id, payload) {
  const dados = normalizarPayloadDespesaGeral(payload);
  const despesaId = String(id || '').trim();
  if (!despesaId) {
    throw new Error('ID da despesa nao informado.');
  }

  const sheet = getSheet(ABA_DESPESAS_GERAIS);
  if (!sheet) {
    throw new Error('Aba DESPESAS_GERAIS nao encontrada.');
  }

  const atual = rowsToObjects(sheet).find(i => String(i.ID || '').trim() === despesaId);
  if (!atual || String(atual.ativo).toLowerCase() !== 'true') {
    throw new Error('Despesa geral nao encontrada ou inativa.');
  }

  const origemAtual = String(atual.origem_fixo_id || despesaId).trim();
  const origemFixoId = dados.fixo ? (origemAtual || despesaId) : '';

  return updateById(
    ABA_DESPESAS_GERAIS,
    'ID',
    despesaId,
    {
      ...dados,
      origem_fixo_id: origemFixoId
    },
    DESPESAS_GERAIS_SCHEMA
  );
}

function deletarDespesaGeral(id) {
  return updateById(
    ABA_DESPESAS_GERAIS,
    'ID',
    id,
    { ativo: false },
    DESPESAS_GERAIS_SCHEMA
  );
}

function obterOrigemPagamentoFinanceiro(origemTipo, origemId) {
  const tipo = normalizarOrigemTipoFinanceiro(origemTipo);
  const id = String(origemId || '').trim();
  if (!id) {
    throw new Error('Origem de pagamento nao informada.');
  }

  if (tipo === ORIGEM_TIPO_COMPRA) {
    const sheet = getSheet(ABA_COMPRAS_FINANCEIRO);
    if (!sheet) throw new Error('Aba COMPRAS nao encontrada.');
    const compra = rowsToObjects(sheet).find(i => i.ID === id);
    if (!compra || String(compra.ativo).toLowerCase() !== 'true') {
      throw new Error('Compra de origem nao encontrada ou inativa.');
    }
    return {
      tipo,
      id,
      item: compra,
      total_previsto: getTotalPrevistoCompraFinanceiro(compra)
    };
  }

  const sheet = getSheet(ABA_DESPESAS_GERAIS);
  if (!sheet) throw new Error('Aba DESPESAS_GERAIS nao encontrada.');
  const despesa = rowsToObjects(sheet).find(i => i.ID === id);
  if (!despesa || String(despesa.ativo).toLowerCase() !== 'true') {
    throw new Error('Despesa geral de origem nao encontrada ou inativa.');
  }
  return {
    tipo,
    id,
    item: despesa,
    total_previsto: getTotalPrevistoDespesaFinanceiro(despesa)
  };
}

function calcularTotalPagoOrigemFinanceiro(origemTipo, origemId, forcarRecarregar) {
  const tipo = normalizarOrigemTipoFinanceiro(origemTipo);
  const id = String(origemId || '').trim();
  const pagamentos = listarPagamentos(!!forcarRecarregar);
  return round2Financeiro(
    pagamentos
      .filter(p => {
        try {
          return normalizarOrigemTipoFinanceiro(p.origem_tipo) === tipo &&
            String(p.origem_id || '').trim() === id;
        } catch (error) {
          return false;
        }
      })
      .reduce((acc, p) => acc + round2Financeiro(parseNumeroBR(p.valor_pago)), 0)
  );
}

function registrarPagamento(origemTipo, origemId, payload) {
  const origem = obterOrigemPagamentoFinanceiro(origemTipo, origemId);
  const dados = { ...(payload || {}) };

  const valorPago = round2Financeiro(parseNumeroBR(dados.valor_pago));
  if (valorPago <= 0) {
    throw new Error('Valor pago deve ser maior que zero.');
  }

  const dataPagamento = normalizarDataFinanceiro(dados.data_pagamento, true, 'Data de pagamento');
  const formaPagamento = validarFormaPagamentoFinanceiro(dados.forma_pagamento, true);
  const observacao = String(dados.observacao || '').trim();

  const totalPagoAtual = calcularTotalPagoOrigemFinanceiro(origem.tipo, origem.id, true);
  const totalPrevisto = round2Financeiro(origem.total_previsto);
  if (totalPagoAtual + valorPago > totalPrevisto + 0.009) {
    throw new Error('Pagamento maior que o valor pendente.');
  }

  const novo = {
    ID: gerarId('PGT'),
    origem_tipo: origem.tipo,
    origem_id: origem.id,
    data_pagamento: dataPagamento,
    valor_pago: valorPago,
    forma_pagamento: formaPagamento,
    observacao,
    ativo: true,
    criado_em: new Date()
  };

  const ok = insert(ABA_PAGAMENTOS, novo, PAGAMENTOS_SCHEMA);
  if (!ok) return null;

  return {
    ...novo,
    criado_em: formatarDataYmdHmFinanceiro(new Date(novo.criado_em))
  };
}

function registrarPagamentoCompra(compraId, payload) {
  const pagamentoCriado = registrarPagamento(ORIGEM_TIPO_COMPRA, compraId, payload);
  const compraAtualizada = (typeof listarCompras === 'function')
    ? (listarCompras(true).find(i => i.ID === compraId) || null)
    : null;
  return {
    pagamentoCriado,
    compraAtualizada
  };
}

function registrarPagamentoDespesaGeral(despesaId, payload) {
  const pagamentoCriado = registrarPagamento(ORIGEM_TIPO_DESPESA, despesaId, payload);
  const despesaAtualizada = listarDespesasGerais(true).find(i => i.ID === despesaId) || null;
  return {
    pagamentoCriado,
    despesaAtualizada
  };
}

function removerPagamento(id) {
  return updateById(
    ABA_PAGAMENTOS,
    'ID',
    id,
    { ativo: false },
    PAGAMENTOS_SCHEMA
  );
}

function ordenarAgregadoFinanceiro(obj, limite) {
  const arr = Object.keys(obj || {}).map(label => ({
    label,
    valor: round2Financeiro(obj[label])
  }));
  arr.sort((a, b) => b.valor - a.valor || a.label.localeCompare(b.label));
  if (typeof limite === 'number' && limite > 0) {
    return arr.slice(0, limite);
  }
  return arr;
}

function gerarChaveMesFinanceiro(data) {
  const d = parseDataFinanceiro(data);
  if (!d) return '';
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function obterResumoDashboardFinanceiro(referenciaYm, forcarRecarregar) {
  const { ref, inicio, fim } = getIntervaloMesFinanceiro(referenciaYm);

  if (!forcarRecarregar) {
    const cached = lerCacheDashboardFinanceiro();
    if (cached && cached.referencia === ref) {
      return cached;
    }
  }

  const compras = (typeof listarCompras === 'function') ? listarCompras(!!forcarRecarregar) : [];
  const despesas = listarDespesasGerais(!!forcarRecarregar);
  const pagamentos = listarPagamentos(!!forcarRecarregar);
  const estoque = (typeof listarEstoque === 'function') ? listarEstoque(!!forcarRecarregar) : [];

  const comprasPorId = {};
  compras.forEach(i => { comprasPorId[String(i.ID || '').trim()] = i; });
  const despesasPorId = {};
  despesas.forEach(i => { despesasPorId[String(i.ID || '').trim()] = i; });

  const pagamentosValidos = pagamentos.filter(p => {
    let tipo = '';
    try {
      tipo = normalizarOrigemTipoFinanceiro(p.origem_tipo);
    } catch (error) {
      return false;
    }
    const origemId = String(p.origem_id || '').trim();
    if (tipo === ORIGEM_TIPO_COMPRA) {
      return !!comprasPorId[origemId];
    }
    return !!despesasPorId[origemId];
  });

  const porTipo = {};
  const porCategoria = {};
  const porFornecedor = {};
  const porMesPagamento = {};
  let gastoPagoMes = 0;
  const contadoresMes = { bruno: 0, zizu: 0 };
  const origensComPagamentoNoMes = {};

  pagamentosValidos.forEach(p => {
    const valor = round2Financeiro(parseNumeroBR(p.valor_pago));
    const dataPag = parseDataFinanceiro(p.data_pagamento);
    if (!dataPag) return;

    const chaveMes = gerarChaveMesFinanceiro(dataPag);
    porMesPagamento[chaveMes] = round2Financeiro((porMesPagamento[chaveMes] || 0) + valor);

    if (dataPag >= inicio && dataPag < fim) {
      gastoPagoMes = round2Financeiro(gastoPagoMes + valor);

      let tipoOrigem = '';
      try {
        tipoOrigem = normalizarOrigemTipoFinanceiro(p.origem_tipo);
      } catch (error) {
        return;
      }
      let itemOrigem = null;
      if (tipoOrigem === ORIGEM_TIPO_COMPRA) {
        itemOrigem = comprasPorId[String(p.origem_id || '').trim()];
      } else {
        itemOrigem = despesasPorId[String(p.origem_id || '').trim()];
      }

      const tipo = tipoOrigem === ORIGEM_TIPO_COMPRA
        ? String(itemOrigem?.tipo || 'SEM_TIPO').trim() || 'SEM_TIPO'
        : 'DESPESA_GERAL';
      const categoria = String(itemOrigem?.categoria || 'SEM_CATEGORIA').trim() || 'SEM_CATEGORIA';
      const fornecedor = String(itemOrigem?.fornecedor || 'SEM_FORNECEDOR').trim() || 'SEM_FORNECEDOR';

      porTipo[tipo] = round2Financeiro((porTipo[tipo] || 0) + valor);
      porCategoria[categoria] = round2Financeiro((porCategoria[categoria] || 0) + valor);
      porFornecedor[fornecedor] = round2Financeiro((porFornecedor[fornecedor] || 0) + valor);
      aplicarRateioPagoPorFinanceiro(contadoresMes, itemOrigem?.pago_por, valor);
      const chaveOrigem = `${tipoOrigem}|${String(p.origem_id || '').trim()}`;
      origensComPagamentoNoMes[chaveOrigem] = true;
    }
  });

  // Evita dupla contagem: se a origem teve pagamento no mes, conta apenas pagamento.
  // Se nao teve pagamento no mes, conta o lancamento do mes.
  compras.forEach(item => {
    const dataLancamento = parseDataFinanceiro(item.comprado_em);
    if (!dataLancamento || dataLancamento < inicio || dataLancamento >= fim) return;
    const chaveOrigem = `${ORIGEM_TIPO_COMPRA}|${String(item.ID || '').trim()}`;
    if (origensComPagamentoNoMes[chaveOrigem]) return;
    aplicarRateioPagoPorFinanceiro(
      contadoresMes,
      item.pago_por,
      getTotalPrevistoCompraFinanceiro(item)
    );
  });

  despesas.forEach(item => {
    const dataLancamento = parseDataFinanceiro(item.data_competencia);
    if (!dataLancamento || dataLancamento < inicio || dataLancamento >= fim) return;
    const chaveOrigem = `${ORIGEM_TIPO_DESPESA}|${String(item.ID || '').trim()}`;
    if (origensComPagamentoNoMes[chaveOrigem]) return;
    aplicarRateioPagoPorFinanceiro(
      contadoresMes,
      item.pago_por,
      getTotalPrevistoDespesaFinanceiro(item)
    );
  });

  const hoje = inicioDoDiaFinanceiro(new Date());
  const limite7 = new Date(hoje.getTime());
  limite7.setDate(limite7.getDate() + 7);

  let pendenteTotal = 0;
  let vencidoTotal = 0;
  let aVencer7Dias = 0;

  const pendentes = [...compras, ...despesas];
  pendentes.forEach(item => {
    const pendente = round2Financeiro(parseNumeroBR(item.total_pendente));
    if (pendente <= 0) return;

    pendenteTotal = round2Financeiro(pendenteTotal + pendente);
    const vencimento = inicioDoDiaFinanceiro(item.data_vencimento);
    if (!vencimento) return;
    if (vencimento.getTime() < hoje.getTime()) {
      vencidoTotal = round2Financeiro(vencidoTotal + pendente);
      return;
    }
    if (vencimento.getTime() <= limite7.getTime()) {
      aVencer7Dias = round2Financeiro(aVencer7Dias + pendente);
    }
  });

  let valorEstoqueTotal = 0;
  const valorEstoquePorTipo = {};
  estoque.forEach(item => {
    const quantidade = parseNumeroBR(item.quantidade);
    const valorUnit = parseNumeroBR(item.custo_unitario || item.valor_unit);
    const valor = round2Financeiro(Math.max(0, quantidade * valorUnit));
    valorEstoqueTotal = round2Financeiro(valorEstoqueTotal + valor);

    const tipo = String(item.tipo || 'SEM_TIPO').trim() || 'SEM_TIPO';
    valorEstoquePorTipo[tipo] = round2Financeiro((valorEstoquePorTipo[tipo] || 0) + valor);
  });

  const meses = [];
  for (let i = 5; i >= 0; i--) {
    const d = new Date(inicio.getFullYear(), inicio.getMonth() - i, 1);
    const chave = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    meses.push({
      mes: chave,
      valor: round2Financeiro(porMesPagamento[chave] || 0)
    });
  }

  const resultado = {
    referencia: ref,
    cards: {
      gasto_pago_mes: gastoPagoMes,
      pendente_total: pendenteTotal,
      vencido_total: vencidoTotal,
      avencer_7_dias: aVencer7Dias,
      valor_estoque_total: round2Financeiro(valorEstoqueTotal),
      quantidade_itens_estoque: estoque.length,
      contador_bruno_mes: round2Financeiro(contadoresMes.bruno),
      contador_zizu_mes: round2Financeiro(contadoresMes.zizu)
    },
    agregados: {
      gasto_por_tipo: ordenarAgregadoFinanceiro(porTipo, 10),
      gasto_por_categoria: ordenarAgregadoFinanceiro(porCategoria, 10),
      gasto_por_fornecedor: ordenarAgregadoFinanceiro(porFornecedor, 10),
      valor_estoque_por_tipo: ordenarAgregadoFinanceiro(valorEstoquePorTipo, 10)
    },
    evolucao_mensal: meses
  };

  salvarCacheDashboardFinanceiro(resultado);
  return resultado;
}

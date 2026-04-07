const ABA_DESPESAS_GERAIS = 'DESPESAS_GERAIS';
const ABA_PAGAMENTOS = 'PAGAMENTOS';
const ABA_PARCELAS_FINANCEIRAS = 'PARCELAS_FINANCEIRAS';
const ABA_COMPRAS_FINANCEIRO = 'COMPRAS';
const ABA_VENDAS_FINANCEIRO = 'VENDAS';
const ABA_PRODUTOS_FINANCEIRO = 'PRODUTOS';

const DESPESAS_GERAIS_CACHE_SCOPE = 'DESPESAS_GERAIS_LISTA_ATIVAS';
const DESPESAS_GERAIS_CACHE_TTL_SEC = 90;

const PAGAMENTOS_CACHE_SCOPE = 'PAGAMENTOS_LISTA_ATIVOS';
const PAGAMENTOS_CACHE_TTL_SEC = 90;

const DASHBOARD_FINANCEIRO_CACHE_SCOPE = 'DASHBOARD_FINANCEIRO';
const DASHBOARD_FINANCEIRO_CACHE_TTL_SEC = 120;

const ORIGEM_TIPO_COMPRA = 'COMPRA';
const ORIGEM_TIPO_DESPESA = 'DESPESA_GERAL';
const ORIGEM_TIPO_VENDA = 'VENDA';
const ORIGEM_TIPO_ESTOQUE = 'ESTOQUE';
const ORIGEM_TIPO_PRODUCAO = 'PRODUCAO';
const NATUREZA_PAGAMENTO = 'PAGAMENTO';
const NATUREZA_RECEBIMENTO = 'RECEBIMENTO';

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
  'parcelas_detalhe_json',
  'fixo',
  'origem_fixo_id',
  'fixo_encerrado_em',
  'fixo_bloqueios_json',
  'observacao'
];

const PAGAMENTOS_SCHEMA = [
  'ID',
  'origem_tipo',
  'origem_id',
  'parcela_alvo_id',
  'natureza',
  'data_pagamento',
  'valor_pago',
  'forma_pagamento',
  'observacao',
  'ativo',
  'criado_em'
];

const PARCELAS_FINANCEIRAS_SCHEMA = [
  'ID',
  'origem_tipo',
  'origem_id',
  'natureza',
  'parcela_numero',
  'parcelas_total',
  'data_prevista',
  'valor_previsto',
  'valor_pago',
  'status',
  'data_quitacao',
  'pagamento_id',
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

function obterValorUnitarioEstoqueDashboard(item) {
  const tipo = String(item?.tipo || '').trim().toUpperCase();
  const precoVendaBruto = String(item?.preco_venda ?? '').trim();
  const precoVenda = parseNumeroBR(precoVendaBruto);
  const custoBruto = String(item?.custo_unitario ?? '').trim();
  const custoUnitario = parseNumeroBR(custoBruto);

  // Para PRODUTO, valuation de estoque usa preco de venda.
  // Sem fallback para custo: se preco de venda estiver vazio, valor do item no dashboard = 0.
  if (tipo === 'PRODUTO') {
    if (precoVendaBruto === '') return 0;
    return precoVenda;
  }

  if (custoBruto !== '') {
    return custoUnitario;
  }
  return parseNumeroBR(item?.valor_unit);
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

function serializarParcelasDetalheFinanceiro(parcelasDetalhe) {
  const lista = Array.isArray(parcelasDetalhe) ? parcelasDetalhe : [];
  if (lista.length === 0) return '';
  return JSON.stringify(
    lista.map((p, idx) => ({
      numero: Number(p.numero || idx + 1),
      data: String(p.data || '').trim(),
      valor: round2Financeiro(parseNumeroBR(p.valor))
    }))
  );
}

function desserializarParcelasDetalheFinanceiro(parcelasDetalheBruto) {
  if (Array.isArray(parcelasDetalheBruto)) {
    return parcelasDetalheBruto;
  }
  const raw = String(parcelasDetalheBruto || '').trim();
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return [];
  }
}

function gerarParcelasPadraoFinanceiro(totalPrevisto, parcelas, dataInicio) {
  const total = round2Financeiro(parseNumeroBR(totalPrevisto));
  const qtdParcelas = Math.max(1, Math.floor(parseNumeroBR(parcelas) || 1));
  const inicio = parseDataFinanceiro(dataInicio);
  if (!inicio) {
    throw new Error('Data base invalida para gerar parcelas.');
  }

  const valorBase = round2Financeiro(total / qtdParcelas);
  const plano = [];
  for (let i = 1; i <= qtdParcelas; i++) {
    const dataParcela = adicionarMesesComAjusteFinanceiro(inicio, i - 1);
    const valor = i === qtdParcelas
      ? round2Financeiro(total - (valorBase * (qtdParcelas - 1)))
      : valorBase;
    plano.push({
      numero: i,
      data: formatarDataYmdFinanceiro(dataParcela),
      valor
    });
  }
  return plano;
}

function normalizarParcelasDetalhePayloadFinanceiro(parcelasDetalheBruto, parcelas, dataInicio, totalPrevisto) {
  const qtdParcelas = Math.max(1, Math.floor(parseNumeroBR(parcelas) || 1));
  const total = round2Financeiro(parseNumeroBR(totalPrevisto));
  const listaBruta = desserializarParcelasDetalheFinanceiro(parcelasDetalheBruto);

  if (!Array.isArray(listaBruta) || listaBruta.length === 0) {
    return gerarParcelasPadraoFinanceiro(total, qtdParcelas, dataInicio);
  }

  if (listaBruta.length !== qtdParcelas) {
    throw new Error('Quantidade de parcelas detalhadas diferente do campo parcelas.');
  }

  const normalizadas = listaBruta.map((p, idx) => {
    const dataBruta = String(p?.data || p?.data_prevista || '').trim();
    const data = normalizarDataFinanceiro(dataBruta, true, `Data da parcela ${idx + 1}`);
    const valor = round2Financeiro(parseNumeroBR(p?.valor ?? p?.valor_previsto));
    if (valor <= 0) {
      throw new Error(`Valor da parcela ${idx + 1} deve ser maior que zero.`);
    }
    return {
      numero: idx + 1,
      data,
      valor
    };
  });

  const soma = round2Financeiro(normalizadas.reduce((acc, p) => acc + p.valor, 0));
  const diferenca = round2Financeiro(total - soma);
  if (Math.abs(diferenca) > 0.02) {
    throw new Error('Soma das parcelas deve ser igual ao valor total.');
  }
  if (Math.abs(diferenca) > 0 && normalizadas.length > 0) {
    const ultimo = normalizadas[normalizadas.length - 1];
    ultimo.valor = round2Financeiro(ultimo.valor + diferenca);
  }

  return normalizadas;
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
  const alvo = acumulador || { bruno: 0, zizu: 0, investimento: 0 };
  const valorSeguro = round2Financeiro(parseNumeroBR(valor));
  if (valorSeguro <= 0) return alvo;
  alvo.bruno = round2Financeiro(alvo.bruno);
  alvo.zizu = round2Financeiro(alvo.zizu);
  alvo.investimento = round2Financeiro(alvo.investimento);

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

  if (pagoPorNormalizado === 'INVESTIMENTO') {
    alvo.investimento = round2Financeiro(alvo.investimento + valorSeguro);
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
  if (tipo === ORIGEM_TIPO_COMPRA || tipo === ORIGEM_TIPO_DESPESA || tipo === ORIGEM_TIPO_VENDA) {
    return tipo;
  }
  throw new Error('Origem de pagamento invalida.');
}

function normalizarNaturezaFinanceiro(natureza) {
  const n = String(natureza || '').trim().toUpperCase();
  if (n === NATUREZA_PAGAMENTO || n === NATUREZA_RECEBIMENTO) {
    return n;
  }
  throw new Error('Natureza financeira invalida.');
}

function getNaturezaOrigemFinanceiro(origemTipo) {
  const tipo = normalizarOrigemTipoFinanceiro(origemTipo);
  if (tipo === ORIGEM_TIPO_VENDA) return NATUREZA_RECEBIMENTO;
  return NATUREZA_PAGAMENTO;
}

function getTotalPrevistoCompraFinanceiro(compra) {
  const quantidade = parseNumeroBR(compra?.quantidade);
  const valorUnit = parseNumeroBR(compra?.valor_unit);
  return round2Financeiro(Math.max(0, quantidade * valorUnit));
}

function getTotalPrevistoDespesaFinanceiro(despesa) {
  return round2Financeiro(Math.max(0, parseNumeroBR(despesa?.valor_total)));
}

function getTotalPrevistoVendaFinanceiro(venda) {
  return round2Financeiro(Math.max(0, parseNumeroBR(venda?.valor_total_venda)));
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

function mapearParcelasPorOrigemFinanceiro() {
  const mapa = {};
  const sheet = getSheet(ABA_PARCELAS_FINANCEIRAS);
  if (!sheet) return mapa;

  rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(i => {
      let tipo = '';
      try {
        tipo = normalizarOrigemTipoFinanceiro(i.origem_tipo);
      } catch (error) {
        return;
      }

      const origemId = String(i.origem_id || '').trim();
      if (!origemId) return;
      const chave = `${tipo}|${origemId}`;
      if (!Array.isArray(mapa[chave])) mapa[chave] = [];

      const valorPrevisto = round2Financeiro(parseNumeroBR(i.valor_previsto));
      const valorPago = round2Financeiro(parseNumeroBR(i.valor_pago));
      const pendente = round2Financeiro(Math.max(0, valorPrevisto - valorPago));
      const statusRaw = String(i.status || '').trim().toUpperCase();
      const status = statusRaw || (valorPago <= 0 ? 'PENDENTE' : (pendente <= 0.009 ? 'PAGO' : 'PARCIAL'));
      mapa[chave].push({
        ID: String(i.ID || '').trim(),
        origem_tipo: tipo,
        origem_id: origemId,
        natureza: (() => {
          try {
            return normalizarNaturezaFinanceiro(i.natureza || getNaturezaOrigemFinanceiro(tipo));
          } catch (error) {
            return getNaturezaOrigemFinanceiro(tipo);
          }
        })(),
        parcela_numero: Number(i.parcela_numero || 0),
        parcelas_total: Number(i.parcelas_total || 0),
        data_prevista: formatarDataYmdFinanceiroSafe(i.data_prevista),
        valor_previsto: valorPrevisto,
        valor_pago: valorPago,
        valor_pendente: pendente,
        status,
        data_quitacao: formatarDataYmdFinanceiroSafe(i.data_quitacao),
        pagamento_id: String(i.pagamento_id || '').trim()
      });
    });

  Object.keys(mapa).forEach(chave => {
    mapa[chave].sort((a, b) => {
      const pa = Number(a.parcela_numero || 0);
      const pb = Number(b.parcela_numero || 0);
      if (pa !== pb) return pa - pb;
      return String(a.ID || '').localeCompare(String(b.ID || ''));
    });
  });

  return mapa;
}

function enriquecerListaComPagamentosFinanceiro(listaItens, origemTipo, getTotalPrevistoFn) {
  const tipo = normalizarOrigemTipoFinanceiro(origemTipo);
  const pagamentos = listarPagamentos();
  const mapa = mapearPagamentosPorOrigemFinanceiro(pagamentos);
  const mapaParcelas = mapearParcelasPorOrigemFinanceiro();

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
      forma_ultimo_pagamento: resumo.forma_ultimo_pagamento || '',
      parcelas_detalhe: mapaParcelas[chave] || []
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

function enriquecerVendasComResumoPagamento(listaVendas) {
  return enriquecerListaComPagamentosFinanceiro(
    listaVendas,
    ORIGEM_TIPO_VENDA,
    getTotalPrevistoVendaFinanceiro
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
      parcela_alvo_id: String(i.parcela_alvo_id || '').trim(),
      natureza: (() => {
        const bruto = String(i.natureza || '').trim();
        if (bruto) {
          try {
            return normalizarNaturezaFinanceiro(bruto);
          } catch (error) {
            return getNaturezaOrigemFinanceiro(i.origem_tipo);
          }
        }
        try {
          return getNaturezaOrigemFinanceiro(i.origem_tipo);
        } catch (error) {
          return NATUREZA_PAGAMENTO;
        }
      })(),
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

function listarParcelasFinanceirasOrigem(origemTipo, origemId) {
  const tipo = normalizarOrigemTipoFinanceiro(origemTipo);
  const id = String(origemId || '').trim();
  if (!id) return [];

  const sheet = getSheet(ABA_PARCELAS_FINANCEIRAS);
  if (!sheet) return [];

  return rowsToObjects(sheet)
    .filter(i =>
      String(i.ativo).toLowerCase() === 'true' &&
      String(i.origem_tipo || '').trim().toUpperCase() === tipo &&
      String(i.origem_id || '').trim() === id
    )
    .map(i => {
      const valorPrevisto = round2Financeiro(parseNumeroBR(i.valor_previsto));
      const valorPago = round2Financeiro(parseNumeroBR(i.valor_pago));
      const pendente = round2Financeiro(Math.max(0, valorPrevisto - valorPago));
      const statusRaw = String(i.status || '').trim().toUpperCase();
      const status = statusRaw || (valorPago <= 0 ? 'PENDENTE' : (pendente <= 0.009 ? 'PAGO' : 'PARCIAL'));
      return {
        ...i,
        natureza: (() => {
          try {
            return normalizarNaturezaFinanceiro(i.natureza || getNaturezaOrigemFinanceiro(tipo));
          } catch (error) {
            return getNaturezaOrigemFinanceiro(tipo);
          }
        })(),
        parcela_numero: Number(i.parcela_numero || 0),
        parcelas_total: Number(i.parcelas_total || 0),
        valor_previsto: valorPrevisto,
        valor_pago: valorPago,
        valor_pendente: pendente,
        status
      };
    })
    .sort((a, b) => {
      const pa = Number(a.parcela_numero || 0);
      const pb = Number(b.parcela_numero || 0);
      if (pa !== pb) return pa - pb;
      return String(a.ID || '').localeCompare(String(b.ID || ''));
    });
}

function limparParcelasFinanceirasOrigem(origemTipo, origemId) {
  const tipo = normalizarOrigemTipoFinanceiro(origemTipo);
  const id = String(origemId || '').trim();
  if (!id) return { ok: true, removidas: 0 };

  const sheet = getSheet(ABA_PARCELAS_FINANCEIRAS);
  if (!sheet) return { ok: true, removidas: 0 };

  const rows = rowsToObjects(sheet).filter(i =>
    String(i.ativo).toLowerCase() === 'true' &&
    String(i.origem_tipo || '').trim().toUpperCase() === tipo &&
    String(i.origem_id || '').trim() === id
  );
  rows.forEach(r => {
    updateById(
      ABA_PARCELAS_FINANCEIRAS,
      'ID',
      r.ID,
      { ativo: false },
      PARCELAS_FINANCEIRAS_SCHEMA
    );
  });
  return { ok: true, removidas: rows.length };
}

function calcularPlanoParcelasFinanceiro(totalPrevisto, parcelas, dataInicio, parcelasDetalhe) {
  const qtdParcelas = Math.max(1, Math.floor(parseNumeroBR(parcelas) || 1));
  const listaDetalhe = Array.isArray(parcelasDetalhe) ? parcelasDetalhe : [];

  if (listaDetalhe.length > 0) {
    return normalizarParcelasDetalhePayloadFinanceiro(
      listaDetalhe,
      qtdParcelas,
      dataInicio,
      totalPrevisto
    ).map(p => ({
      numero: Number(p.numero || 0),
      data: parseDataFinanceiro(p.data),
      valor: round2Financeiro(p.valor),
      total: qtdParcelas
    }));
  }

  return gerarParcelasPadraoFinanceiro(totalPrevisto, qtdParcelas, dataInicio).map(p => ({
    numero: Number(p.numero || 0),
    data: parseDataFinanceiro(p.data),
    valor: round2Financeiro(p.valor),
    total: qtdParcelas
  }));
}

function gerarParcelasFinanceirasOrigem(origemTipo, origemId) {
  const origem = obterOrigemPagamentoFinanceiro(origemTipo, origemId);
  const item = origem.item || {};
  const parcelas = Math.max(1, Math.floor(parseNumeroBR(item.parcelas) || 1));
  const dataInicio = item.data_pagamento || item.data_venda || item.comprado_em || item.data_competencia || new Date();
  const detalhesPersistidos = desserializarParcelasDetalheFinanceiro(item.parcelas_detalhe_json);

  limparParcelasFinanceirasOrigem(origem.tipo, origem.id);

  let plano = [];
  try {
    plano = calcularPlanoParcelasFinanceiro(
      origem.total_previsto,
      parcelas,
      dataInicio,
      detalhesPersistidos
    );
  } catch (error) {
    plano = calcularPlanoParcelasFinanceiro(
      origem.total_previsto,
      parcelas,
      dataInicio,
      []
    );
  }

  plano.forEach(p => {
    insert(ABA_PARCELAS_FINANCEIRAS, {
      ID: gerarId('PAR'),
      origem_tipo: origem.tipo,
      origem_id: origem.id,
      natureza: origem.natureza,
      parcela_numero: p.numero,
      parcelas_total: p.total,
      data_prevista: formatarDataYmdFinanceiro(p.data),
      valor_previsto: round2Financeiro(p.valor),
      valor_pago: 0,
      status: 'PENDENTE',
      data_quitacao: '',
      pagamento_id: '',
      ativo: true,
      criado_em: new Date()
    }, PARCELAS_FINANCEIRAS_SCHEMA);
  });

  return listarParcelasFinanceirasOrigem(origem.tipo, origem.id);
}

function getPrimeiraParcelaPendenteFinanceiro(parcelas) {
  const lista = Array.isArray(parcelas) ? parcelas : [];
  return lista.find(p => round2Financeiro(parseNumeroBR(p.valor_pendente)) > 0.009) || null;
}

function aplicarPagamentoParcelaAlvoFinanceiro(origemTipo, origemId, parcelaAlvoId, pagamentoId, valorPago, dataPagamento) {
  const tipo = normalizarOrigemTipoFinanceiro(origemTipo);
  const id = String(origemId || '').trim();
  const parcelaId = String(parcelaAlvoId || '').trim();
  if (!id || !parcelaId) {
    throw new Error('Parcela alvo nao informada.');
  }

  const parcelas = listarParcelasFinanceirasOrigem(tipo, id);
  const parcela = parcelas.find(p => String(p.ID || '').trim() === parcelaId);
  if (!parcela) {
    throw new Error('Parcela alvo nao encontrada para esta origem.');
  }

  const primeiraPendente = getPrimeiraParcelaPendenteFinanceiro(parcelas);
  if (!primeiraPendente || String(primeiraPendente.ID || '').trim() !== parcelaId) {
    throw new Error('Pagamento deve seguir a ordem das parcelas pendentes.');
  }

  const pendente = round2Financeiro(parseNumeroBR(parcela.valor_pendente));
  const valor = round2Financeiro(parseNumeroBR(valorPago));
  if (pendente <= 0.009) {
    throw new Error('Parcela selecionada ja esta quitada.');
  }
  if (Math.abs(valor - pendente) > 0.009) {
    throw new Error('Valor pago deve ser igual ao valor pendente da parcela selecionada.');
  }

  const novoPago = round2Financeiro(parseNumeroBR(parcela.valor_pago) + valor);
  const quitada = novoPago + 0.009 >= round2Financeiro(parseNumeroBR(parcela.valor_previsto));
  updateById(
    ABA_PARCELAS_FINANCEIRAS,
    'ID',
    parcelaId,
    {
      valor_pago: novoPago,
      status: quitada ? 'PAGO' : 'PARCIAL',
      data_quitacao: String(dataPagamento || ''),
      pagamento_id: String(pagamentoId || '').trim()
    },
    PARCELAS_FINANCEIRAS_SCHEMA
  );

  return {
    valor_alocado: valor,
    valor_excedente: 0,
    parcela_id: parcelaId
  };
}

function aplicarPagamentoEmParcelasFinanceiro(origemTipo, origemId, pagamentoId, valorPago, dataPagamento) {
  const tipo = normalizarOrigemTipoFinanceiro(origemTipo);
  const id = String(origemId || '').trim();
  if (!id) return { valor_alocado: 0, valor_excedente: round2Financeiro(valorPago) };

  let restante = round2Financeiro(parseNumeroBR(valorPago));
  if (restante <= 0) return { valor_alocado: 0, valor_excedente: 0 };

  const parcelas = listarParcelasFinanceirasOrigem(tipo, id);
  if (parcelas.length === 0) {
    gerarParcelasFinanceirasOrigem(tipo, id);
  }
  const parcelasAtuais = listarParcelasFinanceirasOrigem(tipo, id);

  parcelasAtuais.forEach(parcela => {
    if (restante <= 0.009) return;
    const pendente = round2Financeiro(Math.max(0, parcela.valor_previsto - parcela.valor_pago));
    if (pendente <= 0.009) return;

    const aplicado = round2Financeiro(Math.min(restante, pendente));
    const novoPago = round2Financeiro(parcela.valor_pago + aplicado);
    const quitada = novoPago + 0.009 >= parcela.valor_previsto;

    updateById(
      ABA_PARCELAS_FINANCEIRAS,
      'ID',
      parcela.ID,
      {
        valor_pago: novoPago,
        status: quitada ? 'PAGO' : 'PARCIAL',
        data_quitacao: quitada ? String(dataPagamento || '') : (parcela.data_quitacao || ''),
        pagamento_id: String(pagamentoId || '').trim() || (parcela.pagamento_id || '')
      },
      PARCELAS_FINANCEIRAS_SCHEMA
    );

    restante = round2Financeiro(restante - aplicado);
  });

  return {
    valor_alocado: round2Financeiro(parseNumeroBR(valorPago) - restante),
    valor_excedente: round2Financeiro(restante)
  };
}

function regerarParcelasFinanceirasOrigemComPagamentos(origemTipo, origemId) {
  const tipo = normalizarOrigemTipoFinanceiro(origemTipo);
  const id = String(origemId || '').trim();
  if (!id) throw new Error('Origem nao informada.');

  gerarParcelasFinanceirasOrigem(tipo, id);
  const pagamentos = listarPagamentos(true)
    .filter(p => {
      try {
        return normalizarOrigemTipoFinanceiro(p.origem_tipo) === tipo &&
          String(p.origem_id || '').trim() === id;
      } catch (error) {
        return false;
      }
    })
    .sort((a, b) => {
      const da = parseDataFinanceiro(a.data_pagamento)?.getTime() || 0;
      const db = parseDataFinanceiro(b.data_pagamento)?.getTime() || 0;
      if (da !== db) return da - db;
      return String(a.ID || '').localeCompare(String(b.ID || ''));
    });

  pagamentos.forEach(p => {
    aplicarPagamentoEmParcelasFinanceiro(tipo, id, p.ID, p.valor_pago, p.data_pagamento);
  });

  return {
    ok: true,
    origem_tipo: tipo,
    origem_id: id,
    parcelas: listarParcelasFinanceirasOrigem(tipo, id)
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
  const parcelasDetalhe = normalizarParcelasDetalhePayloadFinanceiro(
    dados.parcelas_detalhe ?? dados.parcelas_detalhe_json,
    parcelas,
    dataPagamento || dataCompetencia || new Date(),
    valorTotal
  );
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
    parcelas_detalhe_json: serializarParcelasDetalheFinanceiro(parcelasDetalhe),
    fixo,
    observacao: String(dados.observacao || '').trim()
  };
}

function getSerieRootIdDespesaFixaFinanceiro(item) {
  return String(item?.origem_fixo_id || item?.ID || '').trim();
}

function obterDataCompetenciaDespesaFinanceiro(item) {
  return parseDataFinanceiro(item?.data_competencia)
    || parseDataFinanceiro(item?.data_vencimento)
    || parseDataFinanceiro(item?.criado_em)
    || null;
}

function normalizarYmFinanceiro(valor) {
  const raw = String(valor || '').trim();
  if (!raw) return '';
  if (/^\d{4}-\d{2}$/.test(raw)) return raw;
  const d = parseDataFinanceiro(raw);
  return d ? gerarChaveMesFinanceiro(d) : '';
}

function desserializarBloqueiosFixosYmFinanceiro(bruto) {
  if (Array.isArray(bruto)) {
    const unicos = {};
    bruto.forEach(item => {
      const ym = normalizarYmFinanceiro(item);
      if (ym) unicos[ym] = true;
    });
    return Object.keys(unicos).sort();
  }

  const raw = String(bruto || '').trim();
  if (!raw) return [];

  let lista = [];
  try {
    const parsed = JSON.parse(raw);
    lista = Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    lista = raw.split(',').map(v => String(v || '').trim()).filter(Boolean);
  }

  const unicos = {};
  lista.forEach(item => {
    const ym = normalizarYmFinanceiro(item);
    if (ym) unicos[ym] = true;
  });
  return Object.keys(unicos).sort();
}

function serializarBloqueiosFixosYmFinanceiro(listaYm) {
  const lista = desserializarBloqueiosFixosYmFinanceiro(listaYm);
  if (lista.length === 0) return '';
  return JSON.stringify(lista);
}

function extrairMetaSerieDespesaFixaFinanceiro(rows, rootId) {
  const alvo = String(rootId || '').trim();
  const meta = {
    encerrado_ym: '',
    bloqueios_map: {}
  };
  if (!alvo) {
    return {
      encerrado_ym: '',
      bloqueios_ym: [],
      bloqueios_map: {}
    };
  }

  (Array.isArray(rows) ? rows : []).forEach(item => {
    if (getSerieRootIdDespesaFixaFinanceiro(item) !== alvo) return;

    const encerradoYm = normalizarYmFinanceiro(item?.fixo_encerrado_em);
    if (encerradoYm && (!meta.encerrado_ym || encerradoYm < meta.encerrado_ym)) {
      meta.encerrado_ym = encerradoYm;
    }

    const bloqueios = desserializarBloqueiosFixosYmFinanceiro(item?.fixo_bloqueios_json);
    bloqueios.forEach(ym => {
      meta.bloqueios_map[ym] = true;
    });
  });

  return {
    encerrado_ym: meta.encerrado_ym,
    bloqueios_ym: Object.keys(meta.bloqueios_map).sort(),
    bloqueios_map: meta.bloqueios_map
  };
}

function encerrarSerieDespesaFixaFinanceiro(sheet, rootId, competenciaCorte) {
  const alvo = String(rootId || '').trim();
  const corteYm = normalizarYmFinanceiro(competenciaCorte);
  if (!sheet || !alvo || !corteYm) {
    return { ok: false, atualizadas: 0, encerrado_ym: '' };
  }

  const rows = rowsToObjects(sheet);
  const metaAtual = extrairMetaSerieDespesaFixaFinanceiro(rows, alvo);
  const encerradoYm = metaAtual.encerrado_ym
    ? (metaAtual.encerrado_ym < corteYm ? metaAtual.encerrado_ym : corteYm)
    : corteYm;
  const bloqueiosJson = serializarBloqueiosFixosYmFinanceiro(metaAtual.bloqueios_ym);

  let atualizadas = 0;
  rows.forEach(item => {
    if (getSerieRootIdDespesaFixaFinanceiro(item) !== alvo) return;
    const id = String(item.ID || '').trim();
    if (!id) return;

    const payload = {};
    if (normalizarYmFinanceiro(item.fixo_encerrado_em) !== encerradoYm) {
      payload.fixo_encerrado_em = encerradoYm;
    }
    if (String(item.fixo_bloqueios_json || '').trim() !== bloqueiosJson) {
      payload.fixo_bloqueios_json = bloqueiosJson;
    }

    const dataComp = obterDataCompetenciaDespesaFinanceiro(item);
    const ymComp = gerarChaveMesFinanceiro(dataComp);
    const origemAtual = String(item.origem_fixo_id || '').trim();
    if (ymComp && ymComp >= encerradoYm && (parseBooleanFinanceiro(item.fixo) || !!origemAtual)) {
      payload.fixo = false;
      payload.origem_fixo_id = '';
    }

    if (Object.keys(payload).length === 0) return;
    const ok = updateById(
      ABA_DESPESAS_GERAIS,
      'ID',
      id,
      payload,
      DESPESAS_GERAIS_SCHEMA
    );
    if (ok) atualizadas += 1;
  });

  return {
    ok: true,
    atualizadas,
    encerrado_ym: encerradoYm,
    bloqueios_ym: metaAtual.bloqueios_ym
  };
}

function bloquearCompetenciaSerieDespesaFixaFinanceiro(sheet, rootId, competenciaBloqueada) {
  const alvo = String(rootId || '').trim();
  const bloqueioYm = normalizarYmFinanceiro(competenciaBloqueada);
  if (!sheet || !alvo || !bloqueioYm) {
    return { ok: false, atualizadas: 0, bloqueios_ym: [] };
  }

  const rows = rowsToObjects(sheet);
  const metaAtual = extrairMetaSerieDespesaFixaFinanceiro(rows, alvo);
  const bloqueiosMap = { ...metaAtual.bloqueios_map };
  bloqueiosMap[bloqueioYm] = true;
  const bloqueiosYm = Object.keys(bloqueiosMap).sort();
  const bloqueiosJson = serializarBloqueiosFixosYmFinanceiro(bloqueiosYm);

  let atualizadas = 0;
  rows.forEach(item => {
    if (getSerieRootIdDespesaFixaFinanceiro(item) !== alvo) return;
    const id = String(item.ID || '').trim();
    if (!id) return;

    const payload = {};
    if (String(item.fixo_bloqueios_json || '').trim() !== bloqueiosJson) {
      payload.fixo_bloqueios_json = bloqueiosJson;
    }
    if (metaAtual.encerrado_ym && normalizarYmFinanceiro(item.fixo_encerrado_em) !== metaAtual.encerrado_ym) {
      payload.fixo_encerrado_em = metaAtual.encerrado_ym;
    }

    if (Object.keys(payload).length === 0) return;
    const ok = updateById(
      ABA_DESPESAS_GERAIS,
      'ID',
      id,
      payload,
      DESPESAS_GERAIS_SCHEMA
    );
    if (ok) atualizadas += 1;
  });

  return {
    ok: true,
    atualizadas,
    bloqueios_ym: bloqueiosYm,
    encerrado_ym: metaAtual.encerrado_ym
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
    const todasDespesas = rowsToObjects(sheet);
    const ativasFixas = todasDespesas.filter(i => {
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
      const rootId = getSerieRootIdDespesaFixaFinanceiro(item);
      if (!rootId) return;

      const dataCompetencia = obterDataCompetenciaDespesaFinanceiro(item);
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

      const metaSerie = extrairMetaSerieDespesaFixaFinanceiro(todasDespesas, rootId);
      const encerradoYm = normalizarYmFinanceiro(metaSerie.encerrado_ym);
      const bloqueiosYmMap = { ...(metaSerie.bloqueios_map || {}) };
      const bloqueiosYmJson = serializarBloqueiosFixosYmFinanceiro(Object.keys(bloqueiosYmMap));

      let atual = grupo[grupo.length - 1];
      let cursorCompetencia = atual.dataCompetencia;
      let cursorYm = atual.ym;
      let cursorVencimento = parseDataFinanceiro(atual.item.data_vencimento);
      let templateItem = atual.item;
      let guard = 0;
      while (cursorCompetencia && cursorYm < mesAlvoYm && guard < 60) {
        guard += 1;
        const proximaCompetencia = adicionarMesesComAjusteFinanceiro(cursorCompetencia, 1);
        const proximoYm = gerarChaveMesFinanceiro(proximaCompetencia);
        if (!proximaCompetencia || !proximoYm) break;

        const proximoVencimento = cursorVencimento
          ? adicionarMesesComAjusteFinanceiro(cursorVencimento, 1)
          : null;

        if (porYm[proximoYm]) {
          atual = porYm[proximoYm];
          cursorCompetencia = atual.dataCompetencia;
          cursorYm = atual.ym;
          templateItem = atual.item;
          cursorVencimento = parseDataFinanceiro(atual.item.data_vencimento) || proximoVencimento;
          continue;
        }

        if (encerradoYm && proximoYm >= encerradoYm) {
          break;
        }

        if (bloqueiosYmMap[proximoYm]) {
          cursorCompetencia = proximaCompetencia;
          cursorYm = proximoYm;
          cursorVencimento = proximoVencimento;
          continue;
        }

        const novo = {
          ID: gerarId('DES'),
          descricao: String(templateItem.descricao || '').trim(),
          categoria: String(templateItem.categoria || '').trim(),
          fornecedor: String(templateItem.fornecedor || '').trim(),
          pago_por: String(templateItem.pago_por || '').trim(),
          valor_total: round2Financeiro(parseNumeroBR(templateItem.valor_total)),
          ativo: true,
          criado_em: new Date(),
          data_competencia: formatarDataYmdFinanceiro(proximaCompetencia),
          data_vencimento: proximoVencimento ? formatarDataYmdFinanceiro(proximoVencimento) : '',
          data_pagamento: '',
          forma_pagamento: String(templateItem.forma_pagamento || '').trim(),
          parcelas: normalizarParcelasFinanceiro(templateItem.parcelas, templateItem.forma_pagamento),
          parcelas_detalhe_json: '',
          fixo: true,
          origem_fixo_id: rootId,
          fixo_encerrado_em: encerradoYm || '',
          fixo_bloqueios_json: bloqueiosYmJson,
          observacao: String(templateItem.observacao || '').trim()
        };

        const ok = insert(ABA_DESPESAS_GERAIS, novo, DESPESAS_GERAIS_SCHEMA);
        if (!ok) break;
        try {
          gerarParcelasFinanceirasOrigem(ORIGEM_TIPO_DESPESA, novo.ID);
        } catch (error) {
          // sem acao: recorrencia nao deve falhar por parcelas
        }

        criados += 1;
        const novoRef = {
          item: { ...templateItem, ...novo },
          dataCompetencia: proximaCompetencia,
          ym: proximoYm
        };
        porYm[proximoYm] = novoRef;
        atual = novoRef;
        templateItem = novoRef.item;
        cursorCompetencia = proximaCompetencia;
        cursorYm = proximoYm;
        cursorVencimento = proximoVencimento;
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
  assertCanWrite('Criacao de despesa geral');
  const dados = normalizarPayloadDespesaGeral(payload);
  const novoId = gerarId('DES');
  const novo = {
    ...dados,
    ID: novoId,
    ativo: true,
    criado_em: new Date(),
    origem_fixo_id: dados.fixo ? novoId : '',
    fixo_encerrado_em: '',
    fixo_bloqueios_json: ''
  };

  const ok = insert(ABA_DESPESAS_GERAIS, novo, DESPESAS_GERAIS_SCHEMA);
  if (!ok) return null;
  gerarParcelasFinanceirasOrigem(ORIGEM_TIPO_DESPESA, novo.ID);
  return listarDespesasGerais(true).find(i => i.ID === novo.ID) || null;
}

function atualizarDespesaGeral(id, payload) {
  assertCanWrite('Atualizacao de despesa geral');
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
  const eraFixo = parseBooleanFinanceiro(atual.fixo);
  const encerrandoSerie = eraFixo && !dados.fixo && !!origemAtual;
  const competenciaEncerramento = dados.data_competencia
    || atual.data_competencia
    || atual.data_vencimento
    || atual.criado_em;

  const ok = updateById(
    ABA_DESPESAS_GERAIS,
    'ID',
    despesaId,
    {
      ...dados,
      origem_fixo_id: origemFixoId
    },
    DESPESAS_GERAIS_SCHEMA
  );
  if (ok) {
    if (encerrandoSerie) {
      try {
        encerrarSerieDespesaFixaFinanceiro(sheet, origemAtual, competenciaEncerramento);
      } catch (error) {
        // sem acao: atualizacao principal deve prevalecer
      }
    }
    regerarParcelasFinanceirasOrigemComPagamentos(ORIGEM_TIPO_DESPESA, despesaId);
  }
  return ok;
}

function deletarDespesaGeral(id) {
  assertCanWrite('Exclusao de despesa geral');
  const despesaId = String(id || '').trim();
  if (!despesaId) return false;

  const sheet = getSheet(ABA_DESPESAS_GERAIS);
  if (sheet) {
    ensureSchema(sheet, DESPESAS_GERAIS_SCHEMA);
    const atual = rowsToObjects(sheet).find(i => String(i.ID || '').trim() === despesaId);
    if (atual && String(atual.ativo).toLowerCase() === 'true' && parseBooleanFinanceiro(atual.fixo)) {
      const rootId = getSerieRootIdDespesaFixaFinanceiro(atual);
      const competenciaYm = gerarChaveMesFinanceiro(obterDataCompetenciaDespesaFinanceiro(atual));
      if (rootId && competenciaYm) {
        try {
          bloquearCompetenciaSerieDespesaFixaFinanceiro(sheet, rootId, competenciaYm);
        } catch (error) {
          // sem acao: exclusao principal deve prevalecer
        }
      }
    }
  }

  const ok = updateById(
    ABA_DESPESAS_GERAIS,
    'ID',
    despesaId,
    { ativo: false },
    DESPESAS_GERAIS_SCHEMA
  );
  if (ok) {
    limparParcelasFinanceirasOrigem(ORIGEM_TIPO_DESPESA, despesaId);
  }
  return ok;
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
      total_previsto: getTotalPrevistoCompraFinanceiro(compra),
      natureza: getNaturezaOrigemFinanceiro(tipo)
    };
  }

  if (tipo === ORIGEM_TIPO_VENDA) {
    const sheet = getSheet(ABA_VENDAS_FINANCEIRO);
    if (!sheet) throw new Error('Aba VENDAS nao encontrada.');
    const venda = rowsToObjects(sheet).find(i => i.ID === id);
    if (!venda || String(venda.ativo).toLowerCase() !== 'true') {
      throw new Error('Venda de origem nao encontrada ou inativa.');
    }
    return {
      tipo,
      id,
      item: venda,
      total_previsto: getTotalPrevistoVendaFinanceiro(venda),
      natureza: getNaturezaOrigemFinanceiro(tipo)
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
    total_previsto: getTotalPrevistoDespesaFinanceiro(despesa),
    natureza: getNaturezaOrigemFinanceiro(tipo)
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
  assertCanWrite('Registro de pagamento');
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

  let parcelasOrigem = listarParcelasFinanceirasOrigem(origem.tipo, origem.id);
  if (parcelasOrigem.length === 0) {
    try {
      gerarParcelasFinanceirasOrigem(origem.tipo, origem.id);
      parcelasOrigem = listarParcelasFinanceirasOrigem(origem.tipo, origem.id);
    } catch (error) {
      parcelasOrigem = [];
    }
  }
  const parcelasOrigemTotal = Math.max(1, Math.floor(parseNumeroBR(origem.item?.parcelas) || 1));
  const parcelado = parcelasOrigemTotal > 1;
  let parcelaAlvoId = String(dados.parcela_alvo_id || '').trim();
  if (parcelado) {
    const primeiraPendente = getPrimeiraParcelaPendenteFinanceiro(parcelasOrigem);
    if (!primeiraPendente) {
      throw new Error('Nao ha parcelas pendentes para registrar pagamento.');
    }
    if (!parcelaAlvoId) {
      throw new Error('Selecione a parcela pendente para registrar o pagamento.');
    }
    if (String(primeiraPendente.ID || '').trim() !== parcelaAlvoId) {
      throw new Error('Pagamento deve seguir a ordem das parcelas pendentes.');
    }
    const valorEsperado = round2Financeiro(parseNumeroBR(primeiraPendente.valor_pendente));
    if (Math.abs(valorPago - valorEsperado) > 0.009) {
      throw new Error('Valor pago deve ser igual ao valor pendente da parcela selecionada.');
    }
  } else {
    parcelaAlvoId = '';
  }

  const novo = {
    ID: gerarId('PGT'),
    origem_tipo: origem.tipo,
    origem_id: origem.id,
    parcela_alvo_id: parcelaAlvoId,
    natureza: origem.natureza,
    data_pagamento: dataPagamento,
    valor_pago: valorPago,
    forma_pagamento: formaPagamento,
    observacao,
    ativo: true,
    criado_em: new Date()
  };

  const ok = insert(ABA_PAGAMENTOS, novo, PAGAMENTOS_SCHEMA);
  if (!ok) return null;

  const alocacaoParcelas = parcelaAlvoId
    ? aplicarPagamentoParcelaAlvoFinanceiro(
      origem.tipo,
      origem.id,
      parcelaAlvoId,
      novo.ID,
      valorPago,
      dataPagamento
    )
    : aplicarPagamentoEmParcelasFinanceiro(
      origem.tipo,
      origem.id,
      novo.ID,
      valorPago,
      dataPagamento
    );

  return {
    ...novo,
    criado_em: formatarDataYmdHmFinanceiro(new Date(novo.criado_em)),
    alocacao_parcelas: alocacaoParcelas
  };
}

function registrarPagamentoCompra(compraId, payload) {
  assertCanWrite('Registro de pagamento de compra');
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
  assertCanWrite('Registro de pagamento de despesa geral');
  const pagamentoCriado = registrarPagamento(ORIGEM_TIPO_DESPESA, despesaId, payload);
  const despesaAtualizada = listarDespesasGerais(true).find(i => i.ID === despesaId) || null;
  return {
    pagamentoCriado,
    despesaAtualizada
  };
}

function removerPagamento(id) {
  assertCanWrite('Exclusao de pagamento');
  const pagamento = listarPagamentos(true).find(p => String(p.ID || '').trim() === String(id || '').trim());
  const ok = updateById(
    ABA_PAGAMENTOS,
    'ID',
    id,
    { ativo: false },
    PAGAMENTOS_SCHEMA
  );
  if (ok && pagamento) {
    try {
      regerarParcelasFinanceirasOrigemComPagamentos(pagamento.origem_tipo, pagamento.origem_id);
    } catch (error) {
      // sem acao
    }
  }
  return ok;
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
  const vendas = (typeof listarVendas === 'function') ? listarVendas(!!forcarRecarregar) : [];
  const pagamentos = listarPagamentos(!!forcarRecarregar);
  const estoque = (typeof listarEstoque === 'function') ? listarEstoque(!!forcarRecarregar) : [];
  const producoes = (typeof listarProducao === 'function') ? listarProducao(!!forcarRecarregar) : [];
  const produtosSheet = getSheet(ABA_PRODUTOS_FINANCEIRO);
  const produtosRows = produtosSheet ? rowsToObjects(produtosSheet) : [];
  const precoVendaProdutoPorId = {};
  produtosRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(p => {
      const produtoId = String(p.produto_id || '').trim();
      if (!produtoId || Object.prototype.hasOwnProperty.call(precoVendaProdutoPorId, produtoId)) return;
      const preco = round2Financeiro(parseNumeroBR(p.preco_venda));
      precoVendaProdutoPorId[produtoId] = preco > 0 ? preco : 0;
    });

  const comprasPorId = {};
  compras.forEach(i => { comprasPorId[String(i.ID || '').trim()] = i; });
  const despesasPorId = {};
  despesas.forEach(i => { despesasPorId[String(i.ID || '').trim()] = i; });
  const vendasPorId = {};
  vendas.forEach(i => { vendasPorId[String(i.ID || '').trim()] = i; });

  const eventosFinanceirosValidos = pagamentos.filter(p => {
    let tipo = '';
    try {
      tipo = normalizarOrigemTipoFinanceiro(p.origem_tipo);
    } catch (error) {
      return false;
    }
    const origemId = String(p.origem_id || '').trim();
    if (!origemId) return false;

    if (tipo === ORIGEM_TIPO_COMPRA) return !!comprasPorId[origemId];
    if (tipo === ORIGEM_TIPO_DESPESA) return !!despesasPorId[origemId];
    if (tipo === ORIGEM_TIPO_VENDA) return !!vendasPorId[origemId];
    return false;
  });

  const porTipo = {};
  const porCategoria = {};
  const porFornecedor = {};
  const porMesPagamento = {};

  let gastoPagoMes = 0;
  let recebidoVendasMes = 0;
  const contadoresMes = { bruno: 0, zizu: 0, investimento: 0 };
  let contadorInvestimentoAcumulado = 0;

  eventosFinanceirosValidos.forEach(p => {
    let tipoOrigem = '';
    try {
      tipoOrigem = normalizarOrigemTipoFinanceiro(p.origem_tipo);
    } catch (error) {
      return;
    }
    const natureza = (() => {
      try {
        return normalizarNaturezaFinanceiro(p.natureza || getNaturezaOrigemFinanceiro(tipoOrigem));
      } catch (error) {
        return '';
      }
    })();
    if (!natureza) return;

    const valor = round2Financeiro(parseNumeroBR(p.valor_pago));
    const dataPag = parseDataFinanceiro(p.data_pagamento);
    if (!dataPag || valor <= 0) return;
    let itemOrigemPagamento = null;
    if (natureza === NATUREZA_PAGAMENTO && (tipoOrigem === ORIGEM_TIPO_COMPRA || tipoOrigem === ORIGEM_TIPO_DESPESA)) {
      itemOrigemPagamento = tipoOrigem === ORIGEM_TIPO_COMPRA
        ? comprasPorId[String(p.origem_id || '').trim()]
        : despesasPorId[String(p.origem_id || '').trim()];
      const rateioAcumulado = calcularRateioLinhaPagoPorFinanceiro(itemOrigemPagamento?.pago_por, valor);
      contadorInvestimentoAcumulado = round2Financeiro(
        contadorInvestimentoAcumulado + round2Financeiro(rateioAcumulado.investimento || 0)
      );
    }

    if (natureza === NATUREZA_PAGAMENTO) {
      const chaveMes = gerarChaveMesFinanceiro(dataPag);
      porMesPagamento[chaveMes] = round2Financeiro((porMesPagamento[chaveMes] || 0) + valor);
    }

    if (dataPag < inicio || dataPag >= fim) return;

    if (natureza === NATUREZA_PAGAMENTO) {
      gastoPagoMes = round2Financeiro(gastoPagoMes + valor);

      const itemOrigem = itemOrigemPagamento;

      const tipo = tipoOrigem === ORIGEM_TIPO_COMPRA
        ? String(itemOrigem?.tipo || 'SEM_TIPO').trim() || 'SEM_TIPO'
        : 'DESPESA_GERAL';
      const categoria = String(itemOrigem?.categoria || 'SEM_CATEGORIA').trim() || 'SEM_CATEGORIA';
      const fornecedor = String(itemOrigem?.fornecedor || 'SEM_FORNECEDOR').trim() || 'SEM_FORNECEDOR';

      porTipo[tipo] = round2Financeiro((porTipo[tipo] || 0) + valor);
      porCategoria[categoria] = round2Financeiro((porCategoria[categoria] || 0) + valor);
      porFornecedor[fornecedor] = round2Financeiro((porFornecedor[fornecedor] || 0) + valor);
      aplicarRateioPagoPorFinanceiro(contadoresMes, itemOrigem?.pago_por, valor);
      return;
    }

    if (natureza === NATUREZA_RECEBIMENTO && tipoOrigem === ORIGEM_TIPO_VENDA) {
      recebidoVendasMes = round2Financeiro(recebidoVendasMes + valor);
    }
  });

  let compromissoPrevistoMes = 0;
  const origensComParcelasPagamento = {};
  let pendenteTotal = 0;
  let pendenteCompromissosMes = 0;
  const origemAtivaResumo = (tipo, origemId) => {
    if (!origemId) return false;
    if (tipo === ORIGEM_TIPO_COMPRA) return !!comprasPorId[origemId];
    if (tipo === ORIGEM_TIPO_DESPESA) return !!despesasPorId[origemId];
    if (tipo === ORIGEM_TIPO_VENDA) return !!vendasPorId[origemId];
    return false;
  };
  const sheetParcelas = getSheet(ABA_PARCELAS_FINANCEIRAS);
  if (sheetParcelas) {
    rowsToObjects(sheetParcelas)
      .filter(i => String(i.ativo).toLowerCase() === 'true')
      .forEach(i => {
        let origemTipo = '';
        let natureza = '';
        try {
          origemTipo = normalizarOrigemTipoFinanceiro(i.origem_tipo);
          natureza = normalizarNaturezaFinanceiro(i.natureza || getNaturezaOrigemFinanceiro(origemTipo));
        } catch (error) {
          return;
        }
        if (natureza !== NATUREZA_PAGAMENTO) return;
        const origemId = String(i.origem_id || '').trim();
        if (!origemAtivaResumo(origemTipo, origemId)) return;
        if (origemTipo && origemId) {
          origensComParcelasPagamento[`${origemTipo}|${origemId}`] = true;
        }
        const dataPrevista = parseDataFinanceiro(i.data_prevista);
        if (!dataPrevista) return;
        const valorPrevisto = round2Financeiro(parseNumeroBR(i.valor_previsto));
        const valorPagoParcela = round2Financeiro(parseNumeroBR(i.valor_pago));
        const valorPendenteParcela = round2Financeiro(Math.max(0, valorPrevisto - valorPagoParcela));
        if (dataPrevista >= inicio && dataPrevista < fim) {
          pendenteTotal = round2Financeiro(pendenteTotal + valorPendenteParcela);
        }
        if (dataPrevista < fim && valorPendenteParcela > 0.009) {
          pendenteCompromissosMes = round2Financeiro(pendenteCompromissosMes + valorPendenteParcela);
        }
      });
  }

  const vendasNoMes = round2Financeiro(
    vendas.reduce((acc, item) => {
      const dataVenda = parseDataFinanceiro(item.data_venda);
      if (!dataVenda || dataVenda < inicio || dataVenda >= fim) return acc;
      return acc + round2Financeiro(parseNumeroBR(item.valor_total_venda));
    }, 0)
  );

  const hoje = inicioDoDiaFinanceiro(new Date());
  const limite7 = new Date(hoje.getTime());
  limite7.setDate(limite7.getDate() + 7);

  let vencidoTotal = 0;
  let aVencer7Dias = 0;

  const origensPendentes = [];
  compras.forEach(item => origensPendentes.push({ origem_tipo: ORIGEM_TIPO_COMPRA, item }));
  despesas.forEach(item => origensPendentes.push({ origem_tipo: ORIGEM_TIPO_DESPESA, item }));

  origensPendentes.forEach(entry => {
    const item = entry.item || {};
    const pendente = round2Financeiro(parseNumeroBR(item.total_pendente));
    if (pendente <= 0) return;
    const origemId = String(item.ID || '').trim();
    const origemComParcelas = !!origensComParcelasPagamento[`${entry.origem_tipo}|${origemId}`];
    const vencimento = inicioDoDiaFinanceiro(item.data_vencimento);

    if (!origemComParcelas) {
      const pendenciaAteMesReferencia = !!vencimento && vencimento < fim;
      if (pendenciaAteMesReferencia) {
        pendenteCompromissosMes = round2Financeiro(pendenteCompromissosMes + pendente);
      }
      const pendenciaDoMesReferencia = !!vencimento && vencimento >= inicio && vencimento < fim;
      if (pendenciaDoMesReferencia) {
        pendenteTotal = round2Financeiro(pendenteTotal + pendente);
      }
    }
    if (!vencimento) return;
    if (vencimento.getTime() < hoje.getTime()) {
      vencidoTotal = round2Financeiro(vencidoTotal + pendente);
      return;
    }
    if (vencimento.getTime() <= limite7.getTime()) {
      aVencer7Dias = round2Financeiro(aVencer7Dias + pendente);
    }
  });

  compromissoPrevistoMes = round2Financeiro(gastoPagoMes + pendenteCompromissosMes);

  let valorEstoqueTotal = 0;
  const valorEstoquePorTipo = {};
  let valorProdutosEstoque = 0;
  estoque.forEach(item => {
    const quantidade = parseNumeroBR(item.quantidade);
    const valorUnit = obterValorUnitarioEstoqueDashboard(item);
    const valor = round2Financeiro(Math.max(0, quantidade * valorUnit));
    valorEstoqueTotal = round2Financeiro(valorEstoqueTotal + valor);

    const tipo = String(item.tipo || 'SEM_TIPO').trim() || 'SEM_TIPO';
    valorEstoquePorTipo[tipo] = round2Financeiro((valorEstoquePorTipo[tipo] || 0) + valor);

    if (normalizarTextoSemAcentoFinanceiro(tipo) === 'PRODUTO') {
      valorProdutosEstoque = round2Financeiro(valorProdutosEstoque + valor);
    }
  });

  const statusContaComoEmProducao = status => {
    const s = normalizarTextoSemAcentoFinanceiro(status);
    return s === 'EM PRODUCAO' || s === 'FINALIZACAO';
  };
  const obterQtdRestanteProducao = op => {
    const restanteRaw = String(op?.qtd_restante ?? '').trim();
    if (restanteRaw !== '') {
      return round2Financeiro(Math.max(0, parseNumeroBR(restanteRaw)));
    }
    const planejada = round2Financeiro(parseNumeroBR(op?.qtd_planejada));
    const produzida = round2Financeiro(parseNumeroBR(op?.qtd_produzida_acumulada));
    return round2Financeiro(Math.max(0, planejada - produzida));
  };

  const valorProdutosProducao = round2Financeiro(
    (Array.isArray(producoes) ? producoes : [])
      .reduce((acc, op) => {
        if (!statusContaComoEmProducao(op?.status)) return acc;
        const produtoId = String(op?.produto_id || '').trim();
        const precoVenda = round2Financeiro(precoVendaProdutoPorId[produtoId] || 0);
        if (precoVenda <= 0) return acc;
        const quantidade = obterQtdRestanteProducao(op);
        if (quantidade <= 0) return acc;
        return acc + round2Financeiro(quantidade * precoVenda);
      }, 0)
  );

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
      compromisso_previsto_mes: compromissoPrevistoMes,
      vendas_no_mes: vendasNoMes,
      recebido_vendas_mes: recebidoVendasMes,
      pendente_total: pendenteTotal,
      vencido_total: vencidoTotal,
      avencer_7_dias: aVencer7Dias,
      valor_estoque_total: round2Financeiro(valorEstoqueTotal),
      valor_produtos_estoque: round2Financeiro(valorProdutosEstoque),
      valor_produtos_producao: round2Financeiro(valorProdutosProducao),
      contador_bruno_mes: round2Financeiro(contadoresMes.bruno),
      contador_zizu_mes: round2Financeiro(contadoresMes.zizu),
      contador_investimento_mes: round2Financeiro(contadoresMes.investimento),
      contador_investimento_acumulado: round2Financeiro(contadorInvestimentoAcumulado)
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

function normalizarCardKeyDashboardFinanceiro(cardKey) {
  const chave = String(cardKey || '').trim().toLowerCase();
  const permitidos = {
    gasto_pago_mes: true,
    compromisso_previsto_mes: true,
    vendas_no_mes: true,
    recebido_vendas_mes: true,
    pendente_total: true,
    vencido_total: true,
    avencer_7_dias: true,
    valor_estoque_total: true,
    valor_produtos_estoque: true,
    valor_produtos_producao: true,
    contador_bruno_mes: true,
    contador_zizu_mes: true,
    contador_investimento_mes: true,
    contador_investimento_acumulado: true
  };
  if (!permitidos[chave]) {
    throw new Error('Card de dashboard invalido.');
  }
  return chave;
}

function getRotuloCardDashboardFinanceiro(cardKey) {
  const mapa = {
    gasto_pago_mes: 'Pago no mes',
    compromisso_previsto_mes: 'Compromissos previstos (mes)',
    vendas_no_mes: 'Vendas no mes',
    recebido_vendas_mes: 'Recebido de vendas (mes)',
    pendente_total: 'Pendente total',
    vencido_total: 'Vencido',
    avencer_7_dias: 'A vencer (7 dias)',
    valor_estoque_total: 'Valor em estoque',
    valor_produtos_estoque: 'Valor de produtos em estoque',
    valor_produtos_producao: 'Valor de produtos em producao',
    contador_bruno_mes: 'BRUNO (pago no mes)',
    contador_zizu_mes: 'ZIZU (pago no mes)',
    contador_investimento_mes: 'INVESTIMENTO (pago no mes)',
    contador_investimento_acumulado: 'INVESTIMENTO (acumulado)'
  };
  return mapa[String(cardKey || '').trim().toLowerCase()] || 'Card';
}

function getAbaOrigemDashboardFinanceiro(origemTipo) {
  const tipo = String(origemTipo || '').trim().toUpperCase();
  if (tipo === ORIGEM_TIPO_COMPRA) return 'compras';
  if (tipo === ORIGEM_TIPO_DESPESA) return 'despesas';
  if (tipo === ORIGEM_TIPO_VENDA) return 'vendas';
  if (tipo === ORIGEM_TIPO_ESTOQUE) return 'estoque';
  if (tipo === ORIGEM_TIPO_PRODUCAO) return 'producao';
  return '';
}

function calcularRateioLinhaPagoPorFinanceiro(pagoPor, valor) {
  const valorSeguro = round2Financeiro(parseNumeroBR(valor));
  const resultado = { bruno: 0, zizu: 0, investimento: 0 };
  if (valorSeguro <= 0) return resultado;
  const pagoPorNormalizado = normalizarTextoSemAcentoFinanceiro(pagoPor);
  if (pagoPorNormalizado === 'AMBOS') {
    const metade = round2Financeiro(valorSeguro / 2);
    resultado.bruno = metade;
    resultado.zizu = metade;
    return resultado;
  }
  if (pagoPorNormalizado === 'BRUNO') {
    resultado.bruno = valorSeguro;
    return resultado;
  }
  if (pagoPorNormalizado === 'ZIZU') {
    resultado.zizu = valorSeguro;
    return resultado;
  }
  if (pagoPorNormalizado === 'INVESTIMENTO') {
    resultado.investimento = valorSeguro;
    return resultado;
  }
  return resultado;
}

function ordenarComposicaoCardDashboardFinanceiro(itens, campoValor) {
  const campo = String(campoValor || 'valor').trim() || 'valor';
  const lista = Array.isArray(itens) ? itens : [];
  lista.sort((a, b) => {
    const da = parseDataFinanceiro(a?.data)?.getTime() || 0;
    const db = parseDataFinanceiro(b?.data)?.getTime() || 0;
    if (da !== db) return db - da;
    const va = round2Financeiro(parseNumeroBR(a?.[campo] || 0));
    const vb = round2Financeiro(parseNumeroBR(b?.[campo] || 0));
    if (va !== vb) return vb - va;
    return String(a?.titulo || '').localeCompare(String(b?.titulo || ''));
  });
  return lista;
}

function obterComposicaoCardDashboardFinanceiro(referenciaYm, cardKey, forcarRecarregar) {
  const chave = normalizarCardKeyDashboardFinanceiro(cardKey);
  const { ref, inicio, fim } = getIntervaloMesFinanceiro(referenciaYm);
  const hoje = inicioDoDiaFinanceiro(new Date());
  const limite7 = new Date(hoje.getTime());
  limite7.setDate(limite7.getDate() + 7);

  const compras = (typeof listarCompras === 'function') ? listarCompras(!!forcarRecarregar) : [];
  const despesas = listarDespesasGerais(!!forcarRecarregar);
  const vendas = (typeof listarVendas === 'function') ? listarVendas(!!forcarRecarregar) : [];
  const pagamentos = listarPagamentos(!!forcarRecarregar);
  const estoque = (typeof listarEstoque === 'function') ? listarEstoque(!!forcarRecarregar) : [];
  const producoes = (typeof listarProducao === 'function') ? listarProducao(!!forcarRecarregar) : [];
  const produtosSheet = getSheet(ABA_PRODUTOS_FINANCEIRO);
  const produtosRows = produtosSheet ? rowsToObjects(produtosSheet) : [];
  const precoVendaProdutoPorId = {};
  produtosRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(p => {
      const produtoId = String(p.produto_id || '').trim();
      if (!produtoId || Object.prototype.hasOwnProperty.call(precoVendaProdutoPorId, produtoId)) return;
      const preco = round2Financeiro(parseNumeroBR(p.preco_venda));
      precoVendaProdutoPorId[produtoId] = preco > 0 ? preco : 0;
    });

  const comprasPorId = {};
  compras.forEach(i => { comprasPorId[String(i.ID || '').trim()] = i; });
  const despesasPorId = {};
  despesas.forEach(i => { despesasPorId[String(i.ID || '').trim()] = i; });
  const vendasPorId = {};
  vendas.forEach(i => { vendasPorId[String(i.ID || '').trim()] = i; });
  const estoquePorId = {};
  estoque.forEach(i => { estoquePorId[String(i.ID || '').trim()] = i; });
  const producoesPorId = {};
  producoes.forEach(i => { producoesPorId[String(i.producao_id || '').trim()] = i; });

  function obterMetaOrigem(origemTipo, origemId) {
    const tipo = String(origemTipo || '').trim().toUpperCase();
    const id = String(origemId || '').trim();
    if (!id) return null;

    if (tipo === ORIGEM_TIPO_COMPRA) {
      const item = comprasPorId[id];
      if (!item) return null;
      return {
        origem_tipo: tipo,
        origem_id: id,
        origem_aba: getAbaOrigemDashboardFinanceiro(tipo),
        titulo: `Compra: ${String(item.item || id).trim()}`,
        detalhe: `${String(item.tipo || 'SEM_TIPO').trim() || 'SEM_TIPO'} | ${String(item.categoria || 'SEM_CATEGORIA').trim() || 'SEM_CATEGORIA'}`,
        pago_por: String(item.pago_por || '').trim(),
        fornecedor: String(item.fornecedor || '').trim()
      };
    }

    if (tipo === ORIGEM_TIPO_DESPESA) {
      const item = despesasPorId[id];
      if (!item) return null;
      return {
        origem_tipo: tipo,
        origem_id: id,
        origem_aba: getAbaOrigemDashboardFinanceiro(tipo),
        titulo: `Despesa: ${String(item.descricao || id).trim()}`,
        detalhe: `${String(item.categoria || 'SEM_CATEGORIA').trim() || 'SEM_CATEGORIA'} | ${String(item.fornecedor || 'SEM_FORNECEDOR').trim() || 'SEM_FORNECEDOR'}`,
        pago_por: String(item.pago_por || '').trim(),
        fornecedor: String(item.fornecedor || '').trim()
      };
    }

    if (tipo === ORIGEM_TIPO_VENDA) {
      const item = vendasPorId[id];
      if (!item) return null;
      return {
        origem_tipo: tipo,
        origem_id: id,
        origem_aba: getAbaOrigemDashboardFinanceiro(tipo),
        titulo: `Venda: ${String(item.item || id).trim()}`,
        detalhe: `Qtd: ${round2Financeiro(parseNumeroBR(item.quantidade))} ${String(item.unidade || '').trim() || ''}`.trim(),
        pago_por: String(item.recebido_por || '').trim(),
        fornecedor: ''
      };
    }

    if (tipo === ORIGEM_TIPO_ESTOQUE) {
      const item = estoquePorId[id];
      if (!item) return null;
      return {
        origem_tipo: tipo,
        origem_id: id,
        origem_aba: getAbaOrigemDashboardFinanceiro(tipo),
        titulo: `Estoque: ${String(item.item || id).trim()}`,
        detalhe: `${String(item.tipo || 'SEM_TIPO').trim() || 'SEM_TIPO'} | ${String(item.categoria || 'SEM_CATEGORIA').trim() || 'SEM_CATEGORIA'}`,
        pago_por: String(item.pago_por || '').trim(),
        fornecedor: String(item.fornecedor || '').trim()
      };
    }

    if (tipo === ORIGEM_TIPO_PRODUCAO) {
      const op = producoesPorId[id];
      if (!op) return null;
      const nomeProduto = String(op.nome_produto || '').trim();
      const nomeOrdem = String(op.nome_ordem || '').trim();
      return {
        origem_tipo: tipo,
        origem_id: id,
        origem_aba: getAbaOrigemDashboardFinanceiro(tipo),
        titulo: `Producao: ${nomeOrdem || nomeProduto || id}`,
        detalhe: `Produto: ${nomeProduto || '-'} | Status: ${String(op.status || '').trim() || '-'}`,
        pago_por: '',
        fornecedor: ''
      };
    }

    return null;
  }

  const pagamentosValidos = pagamentos.filter(p => {
    let tipo = '';
    try {
      tipo = normalizarOrigemTipoFinanceiro(p.origem_tipo);
    } catch (error) {
      return false;
    }
    const origemId = String(p.origem_id || '').trim();
    if (!origemId) return false;
    if (tipo === ORIGEM_TIPO_COMPRA) return !!comprasPorId[origemId];
    if (tipo === ORIGEM_TIPO_DESPESA) return !!despesasPorId[origemId];
    if (tipo === ORIGEM_TIPO_VENDA) return !!vendasPorId[origemId];
    return false;
  });

  const parcelasAtivas = (() => {
    const sheet = getSheet(ABA_PARCELAS_FINANCEIRAS);
    if (!sheet) return [];
    return rowsToObjects(sheet)
      .filter(i => String(i.ativo).toLowerCase() === 'true')
      .map(i => {
        const tipo = String(i.origem_tipo || '').trim().toUpperCase();
        const origemId = String(i.origem_id || '').trim();
        const dataPrevista = formatarDataYmdFinanceiroSafe(i.data_prevista);
        const valorPrevisto = round2Financeiro(parseNumeroBR(i.valor_previsto));
        const valorPago = round2Financeiro(parseNumeroBR(i.valor_pago));
        const valorPendente = round2Financeiro(Math.max(0, valorPrevisto - valorPago));
        let natureza = '';
        try {
          natureza = normalizarNaturezaFinanceiro(i.natureza || getNaturezaOrigemFinanceiro(tipo));
        } catch (error) {
          natureza = '';
        }
        return {
          ID: String(i.ID || '').trim(),
          origem_tipo: tipo,
          origem_id: origemId,
          natureza,
          data_prevista: dataPrevista,
          valor_previsto: valorPrevisto,
          valor_pago: valorPago,
          valor_pendente: valorPendente,
          parcela_numero: Number(i.parcela_numero || 0),
          parcelas_total: Number(i.parcelas_total || 0)
        };
      })
      .filter(i => {
        if (!i.origem_id || i.valor_previsto <= 0) return false;
        if (i.origem_tipo === ORIGEM_TIPO_COMPRA) return !!comprasPorId[i.origem_id];
        if (i.origem_tipo === ORIGEM_TIPO_DESPESA) return !!despesasPorId[i.origem_id];
        if (i.origem_tipo === ORIGEM_TIPO_VENDA) return !!vendasPorId[i.origem_id];
        return false;
      });
  })();

  const itensPagamentosMes = pagamentosValidos
    .filter(p => {
      const dataPag = parseDataFinanceiro(p.data_pagamento);
      if (!dataPag || dataPag < inicio || dataPag >= fim) return false;
      let natureza = '';
      try {
        natureza = normalizarNaturezaFinanceiro(p.natureza || getNaturezaOrigemFinanceiro(p.origem_tipo));
      } catch (error) {
        natureza = '';
      }
      return natureza === NATUREZA_PAGAMENTO;
    })
    .map(p => {
      const tipo = String(p.origem_tipo || '').trim().toUpperCase();
      const origemId = String(p.origem_id || '').trim();
      const meta = obterMetaOrigem(tipo, origemId);
      if (!meta) return null;
      return {
        linha_id: `PGT:${String(p.ID || '').trim()}`,
        tipo_linha: 'pagamento',
        titulo: meta.titulo,
        detalhe: `${meta.detalhe} | Forma: ${String(p.forma_pagamento || '-').trim() || '-'}`,
        data: formatarDataYmdFinanceiroSafe(p.data_pagamento),
        valor: round2Financeiro(parseNumeroBR(p.valor_pago)),
        origem_tipo: meta.origem_tipo,
        origem_id: meta.origem_id,
        origem_aba: meta.origem_aba,
        pagamento_id: String(p.ID || '').trim(),
        pode_editar_origem: true,
        pode_excluir_origem: true,
        pode_remover_pagamento: true,
        eh_contagem: false
      };
    })
    .filter(Boolean);

  let itens = [];
  let campoValor = 'valor';
  const statusContaComoEmProducao = status => {
    const s = normalizarTextoSemAcentoFinanceiro(status);
    return s === 'EM PRODUCAO' || s === 'FINALIZACAO';
  };
  const obterQtdRestanteProducao = op => {
    const restanteRaw = String(op?.qtd_restante ?? '').trim();
    if (restanteRaw !== '') {
      return round2Financeiro(Math.max(0, parseNumeroBR(restanteRaw)));
    }
    const planejada = round2Financeiro(parseNumeroBR(op?.qtd_planejada));
    const produzida = round2Financeiro(parseNumeroBR(op?.qtd_produzida_acumulada));
    return round2Financeiro(Math.max(0, planejada - produzida));
  };

  if (chave === 'gasto_pago_mes') {
    itens = itensPagamentosMes;
  } else if (chave === 'compromisso_previsto_mes') {
    const origensComParcelasPagamento = {};
    parcelasAtivas.forEach(parcela => {
      if (parcela.natureza !== NATUREZA_PAGAMENTO) return;
      const tipo = String(parcela.origem_tipo || '').trim().toUpperCase();
      const origemId = String(parcela.origem_id || '').trim();
      if (!tipo || !origemId) return;
      origensComParcelasPagamento[`${tipo}|${origemId}`] = true;
    });

    const itensPendentesParcelas = parcelasAtivas
      .filter(parcela => {
        if (parcela.natureza !== NATUREZA_PAGAMENTO) return false;
        const dataPrev = parseDataFinanceiro(parcela.data_prevista);
        if (!dataPrev || dataPrev >= fim) return false;
        return round2Financeiro(parcela.valor_pendente) > 0.009;
      })
      .map(parcela => {
        const meta = obterMetaOrigem(parcela.origem_tipo, parcela.origem_id);
        if (!meta) return null;
        return {
          linha_id: `PAR_PEND:${String(parcela.ID || '').trim()}`,
          tipo_linha: 'parcela_pendente',
          titulo: meta.titulo,
          detalhe: `${meta.detalhe} | Parcela ${Number(parcela.parcela_numero || 0)}/${Number(parcela.parcelas_total || 0)} | Pendente`,
          data: String(parcela.data_prevista || '').trim(),
          valor: round2Financeiro(parcela.valor_pendente),
          origem_tipo: meta.origem_tipo,
          origem_id: meta.origem_id,
          origem_aba: meta.origem_aba,
          pagamento_id: '',
          pode_editar_origem: true,
          pode_excluir_origem: true,
          pode_remover_pagamento: false,
          eh_contagem: false
        };
      })
      .filter(Boolean);

    const itensPendentesSemParcelas = [];
    const origensPendentes = [];
    compras.forEach(item => origensPendentes.push({ origem_tipo: ORIGEM_TIPO_COMPRA, item }));
    despesas.forEach(item => origensPendentes.push({ origem_tipo: ORIGEM_TIPO_DESPESA, item }));
    origensPendentes.forEach(entry => {
      const item = entry.item || {};
      const pendente = round2Financeiro(parseNumeroBR(item.total_pendente));
      if (pendente <= 0.009) return;
      const origemId = String(item.ID || '').trim();
      if (!origemId) return;
      if (origensComParcelasPagamento[`${entry.origem_tipo}|${origemId}`]) return;
      const vencimento = inicioDoDiaFinanceiro(item.data_vencimento);
      if (!vencimento || vencimento >= fim) return;
      const meta = obterMetaOrigem(entry.origem_tipo, origemId);
      if (!meta) return;
      itensPendentesSemParcelas.push({
        linha_id: `ORIG_PEND:${entry.origem_tipo}:${origemId}`,
        tipo_linha: 'origem_pendente',
        titulo: meta.titulo,
        detalhe: `${meta.detalhe} | Pendente`,
        data: formatarDataYmdFinanceiroSafe(vencimento),
        valor: pendente,
        origem_tipo: meta.origem_tipo,
        origem_id: meta.origem_id,
        origem_aba: meta.origem_aba,
        pagamento_id: '',
        pode_editar_origem: true,
        pode_excluir_origem: true,
        pode_remover_pagamento: false,
        eh_contagem: false
      });
    });

    itens = [
      ...itensPagamentosMes,
      ...itensPendentesParcelas,
      ...itensPendentesSemParcelas
    ];
  } else if (chave === 'vendas_no_mes') {
    itens = vendas
      .filter(v => {
        const dataVenda = parseDataFinanceiro(v.data_venda);
        return !!dataVenda && dataVenda >= inicio && dataVenda < fim;
      })
      .map(v => {
        const id = String(v.ID || '').trim();
        const meta = obterMetaOrigem(ORIGEM_TIPO_VENDA, id);
        if (!meta) return null;
        return {
          linha_id: `VND:${id}`,
          tipo_linha: 'origem',
          titulo: meta.titulo,
          detalhe: `${meta.detalhe} | Forma: ${String(v.forma_pagamento || '-').trim() || '-'}`,
          data: formatarDataYmdFinanceiroSafe(v.data_venda),
          valor: round2Financeiro(parseNumeroBR(v.valor_total_venda)),
          origem_tipo: meta.origem_tipo,
          origem_id: meta.origem_id,
          origem_aba: meta.origem_aba,
          pagamento_id: '',
          pode_editar_origem: true,
          pode_excluir_origem: true,
          pode_remover_pagamento: false,
          eh_contagem: false
        };
      })
      .filter(Boolean);
  } else if (chave === 'recebido_vendas_mes') {
    itens = pagamentosValidos
      .filter(p => {
        const tipo = String(p.origem_tipo || '').trim().toUpperCase();
        if (tipo !== ORIGEM_TIPO_VENDA) return false;
        const dataPag = parseDataFinanceiro(p.data_pagamento);
        if (!dataPag || dataPag < inicio || dataPag >= fim) return false;
        let natureza = '';
        try {
          natureza = normalizarNaturezaFinanceiro(p.natureza || getNaturezaOrigemFinanceiro(tipo));
        } catch (error) {
          natureza = '';
        }
        return natureza === NATUREZA_RECEBIMENTO;
      })
      .map(p => {
        const meta = obterMetaOrigem(ORIGEM_TIPO_VENDA, p.origem_id);
        if (!meta) return null;
        return {
          linha_id: `PGT:${String(p.ID || '').trim()}`,
          tipo_linha: 'pagamento',
          titulo: meta.titulo,
          detalhe: `${meta.detalhe} | Forma: ${String(p.forma_pagamento || '-').trim() || '-'}`,
          data: formatarDataYmdFinanceiroSafe(p.data_pagamento),
          valor: round2Financeiro(parseNumeroBR(p.valor_pago)),
          origem_tipo: meta.origem_tipo,
          origem_id: meta.origem_id,
          origem_aba: meta.origem_aba,
          pagamento_id: String(p.ID || '').trim(),
          pode_editar_origem: true,
          pode_excluir_origem: true,
          pode_remover_pagamento: true,
          eh_contagem: false
        };
      })
      .filter(Boolean);
  } else if (chave === 'pendente_total' || chave === 'vencido_total' || chave === 'avencer_7_dias') {
    itens = [];
    const origensPendentes = [];
    compras.forEach(item => origensPendentes.push({ origem_tipo: ORIGEM_TIPO_COMPRA, item }));
    despesas.forEach(item => origensPendentes.push({ origem_tipo: ORIGEM_TIPO_DESPESA, item }));
    const origensComParcelasPagamento = {};
    parcelasAtivas.forEach(parcela => {
      if (parcela.natureza !== NATUREZA_PAGAMENTO) return;
      const origemKey = `${String(parcela.origem_tipo || '').trim()}|${String(parcela.origem_id || '').trim()}`;
      if (!origemKey || origemKey === '|') return;
      origensComParcelasPagamento[origemKey] = true;
    });

    if (chave === 'pendente_total') {
      parcelasAtivas.forEach(parcela => {
        if (parcela.natureza !== NATUREZA_PAGAMENTO) return;
        const dataPrev = parseDataFinanceiro(parcela.data_prevista);
        if (!dataPrev || dataPrev < inicio || dataPrev >= fim) return;
        const pendenteParcela = round2Financeiro(parseNumeroBR(parcela.valor_pendente));
        if (pendenteParcela <= 0) return;

        const meta = obterMetaOrigem(parcela.origem_tipo, parcela.origem_id);
        if (!meta) return;

        itens.push({
          linha_id: `PENPAR:${String(parcela.ID || '').trim()}`,
          tipo_linha: 'parcela_pendente',
          titulo: meta.titulo,
          detalhe: `${meta.detalhe} | Parcela ${Number(parcela.parcela_numero || 0)}/${Number(parcela.parcelas_total || 0)}`,
          data: String(parcela.data_prevista || '').trim(),
          valor: pendenteParcela,
          origem_tipo: meta.origem_tipo,
          origem_id: meta.origem_id,
          origem_aba: meta.origem_aba,
          pagamento_id: '',
          pode_editar_origem: true,
          pode_excluir_origem: true,
          pode_remover_pagamento: false,
          eh_contagem: false
        });
      });
    }

    origensPendentes.forEach(entry => {
      const item = entry.item || {};
      const pendente = round2Financeiro(parseNumeroBR(item.total_pendente));
      if (pendente <= 0) return;
      const origemId = String(item.ID || '').trim();
      const origemKey = `${entry.origem_tipo}|${origemId}`;
      const origemTemParcela = !!origensComParcelasPagamento[origemKey];
      const dataVenc = inicioDoDiaFinanceiro(item.data_vencimento);

      if (chave === 'vencido_total') {
        if (!dataVenc || dataVenc.getTime() >= hoje.getTime()) return;
      } else if (chave === 'avencer_7_dias') {
        if (!dataVenc || dataVenc.getTime() < hoje.getTime() || dataVenc.getTime() > limite7.getTime()) return;
      } else if (chave === 'pendente_total') {
        if (origemTemParcela) return;
        if (!dataVenc || dataVenc < inicio || dataVenc >= fim) return;
      }

      const meta = obterMetaOrigem(entry.origem_tipo, item.ID);
      if (!meta) return;
      itens.push({
        linha_id: `PEN:${entry.origem_tipo}:${String(item.ID || '').trim()}`,
        tipo_linha: 'origem',
        titulo: meta.titulo,
        detalhe: `${meta.detalhe} | Status: ${String(item.status_pagamento || 'PENDENTE').trim() || 'PENDENTE'}`,
        data: formatarDataYmdFinanceiroSafe(item.data_vencimento),
        valor: pendente,
        origem_tipo: meta.origem_tipo,
        origem_id: meta.origem_id,
        origem_aba: meta.origem_aba,
        pagamento_id: '',
        pode_editar_origem: true,
        pode_excluir_origem: true,
        pode_remover_pagamento: false,
        eh_contagem: false
      });
    });
  } else if (chave === 'valor_produtos_estoque') {
    itens = estoque
      .filter(item => normalizarTextoSemAcentoFinanceiro(item.tipo) === 'PRODUTO')
      .map(item => {
        const id = String(item.ID || '').trim();
        const meta = obterMetaOrigem(ORIGEM_TIPO_ESTOQUE, id);
        if (!meta) return null;
        const quantidade = round2Financeiro(parseNumeroBR(item.quantidade));
        const valorUnit = round2Financeiro(obterValorUnitarioEstoqueDashboard(item));
        const valor = round2Financeiro(Math.max(0, quantidade * valorUnit));
        if (valor <= 0) return null;
        return {
          linha_id: `EST_PRD:${id}`,
          tipo_linha: 'estoque_produto',
          titulo: meta.titulo,
          detalhe: `${meta.detalhe} | Qtd: ${quantidade} ${String(item.unidade || '').trim() || ''}`.trim(),
          data: formatarDataYmdFinanceiroSafe(item.comprado_em),
          valor,
          origem_tipo: meta.origem_tipo,
          origem_id: meta.origem_id,
          origem_aba: meta.origem_aba,
          pagamento_id: '',
          pode_editar_origem: true,
          pode_excluir_origem: true,
          pode_remover_pagamento: false,
          eh_contagem: false
        };
      })
      .filter(Boolean);
  } else if (chave === 'valor_produtos_producao') {
    itens = producoes
      .filter(op => statusContaComoEmProducao(op?.status))
      .map(op => {
        const opId = String(op?.producao_id || '').trim();
        const produtoId = String(op?.produto_id || '').trim();
        const meta = obterMetaOrigem(ORIGEM_TIPO_PRODUCAO, opId);
        if (!meta) return null;
        const precoVenda = round2Financeiro(precoVendaProdutoPorId[produtoId] || 0);
        if (precoVenda <= 0) return null;
        const quantidade = obterQtdRestanteProducao(op);
        if (quantidade <= 0) return null;
        return {
          linha_id: `OP_VAL:${opId}`,
          tipo_linha: 'producao_produto',
          titulo: meta.titulo,
          detalhe: `${meta.detalhe} | Qtd restante: ${quantidade} ${String(op?.unidade_produto || '').trim() || ''}`.trim(),
          data: formatarDataYmdFinanceiroSafe(op?.data_inicio || op?.criado_em),
          valor: round2Financeiro(quantidade * precoVenda),
          origem_tipo: meta.origem_tipo,
          origem_id: meta.origem_id,
          origem_aba: meta.origem_aba,
          pagamento_id: '',
          pode_editar_origem: false,
          pode_excluir_origem: false,
          pode_remover_pagamento: false,
          eh_contagem: false
        };
      })
      .filter(Boolean);
  } else if (chave === 'valor_estoque_total') {
    itens = estoque.map(item => {
      const id = String(item.ID || '').trim();
      const meta = obterMetaOrigem(ORIGEM_TIPO_ESTOQUE, id);
      if (!meta) return null;
      const quantidade = round2Financeiro(parseNumeroBR(item.quantidade));
      const valorUnit = round2Financeiro(obterValorUnitarioEstoqueDashboard(item));
      return {
        linha_id: `EST:${id}`,
        tipo_linha: 'estoque',
        titulo: meta.titulo,
        detalhe: `${meta.detalhe} | Qtd: ${quantidade} ${String(item.unidade || '').trim() || ''}`.trim(),
        data: formatarDataYmdFinanceiroSafe(item.comprado_em),
        valor: round2Financeiro(Math.max(0, quantidade * valorUnit)),
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
  } else if (
    chave === 'contador_bruno_mes'
    || chave === 'contador_zizu_mes'
    || chave === 'contador_investimento_mes'
    || chave === 'contador_investimento_acumulado'
  ) {
    const filtroMes = chave !== 'contador_investimento_acumulado';
    const alvo = chave === 'contador_bruno_mes'
      ? 'bruno'
      : (chave === 'contador_zizu_mes' ? 'zizu' : 'investimento');
    itens = pagamentosValidos
      .filter(p => {
        const dataPag = parseDataFinanceiro(p.data_pagamento);
        if (!dataPag) return false;
        if (filtroMes && (dataPag < inicio || dataPag >= fim)) return false;
        const tipo = String(p.origem_tipo || '').trim().toUpperCase();
        if (tipo !== ORIGEM_TIPO_COMPRA && tipo !== ORIGEM_TIPO_DESPESA) return false;
        let natureza = '';
        try {
          natureza = normalizarNaturezaFinanceiro(p.natureza || getNaturezaOrigemFinanceiro(tipo));
        } catch (error) {
          natureza = '';
        }
        return natureza === NATUREZA_PAGAMENTO;
      })
      .map(p => {
        const meta = obterMetaOrigem(p.origem_tipo, p.origem_id);
        if (!meta) return null;
        const valorPago = round2Financeiro(parseNumeroBR(p.valor_pago));
        const rateio = calcularRateioLinhaPagoPorFinanceiro(meta.pago_por, valorPago);
        const valorLinha = round2Financeiro(rateio[alvo] || 0);
        if (valorLinha <= 0) return null;
        return {
          linha_id: `PGT:${String(p.ID || '').trim()}:${alvo}`,
          tipo_linha: 'rateio',
          titulo: meta.titulo,
          detalhe: `${meta.detalhe} | Pago por: ${meta.pago_por || '-'} | Valor original: ${valorPago.toFixed(2)}`,
          data: formatarDataYmdFinanceiroSafe(p.data_pagamento),
          valor: valorLinha,
          origem_tipo: meta.origem_tipo,
          origem_id: meta.origem_id,
          origem_aba: meta.origem_aba,
          pagamento_id: String(p.ID || '').trim(),
          pode_editar_origem: true,
          pode_excluir_origem: true,
          pode_remover_pagamento: true,
          eh_contagem: false
        };
      })
      .filter(Boolean);
  }

  ordenarComposicaoCardDashboardFinanceiro(itens, campoValor);

  const total = round2Financeiro(
    itens.reduce((acc, item) => {
      const valorItem = round2Financeiro(parseNumeroBR(item?.[campoValor] || 0));
      return acc + valorItem;
    }, 0)
  );

  return {
    referencia: ref,
    card_key: chave,
    card_label: getRotuloCardDashboardFinanceiro(chave),
    campo_valor: campoValor,
    total,
    total_itens: itens.length,
    itens
  };
}

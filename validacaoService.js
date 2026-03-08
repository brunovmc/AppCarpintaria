/**
 * =========================
 * VALIDACAO SERVICE
 * Aba: VALIDACAO
 * Colunas:
 * A - TIPO
 * B - UNIDADE
 * C - CATEGORIA
 * D - FORNECEDOR
 * E - VALORKWH (uso futuro)
 * F - DESPESAS
 * G - FORMA_PAGAMENTO
 * H - PAGO_POR
 * =========================
 */

/* =========================
   FUNÇÕES EXISTENTES (MANTIDAS)
   ========================= */

function testeValidacao() {
  Logger.log(obterValidacoes());
}

const VALIDACOES_CACHE_TTL_SEC = 21600; // 6 horas (maximo do CacheService)
const VALIDACOES_CACHE_VERSION = 'v1';

function getValidacoesCacheKey() {
  const id = String(typeof DATA_SPREADSHEET_ID === 'string' ? DATA_SPREADSHEET_ID : '').trim();
  return `VALIDACOES_CACHE:${id || 'SEM_ID'}:${VALIDACOES_CACHE_VERSION}`;
}

function lerValidacoesDoCache() {
  try {
    const cache = CacheService.getScriptCache();
    const raw = cache.get(getValidacoesCacheKey());
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch (error) {
    return null;
  }
}

function salvarValidacoesNoCache(payload) {
  try {
    const cache = CacheService.getScriptCache();
    cache.put(
      getValidacoesCacheKey(),
      JSON.stringify(payload || {}),
      VALIDACOES_CACHE_TTL_SEC
    );
  } catch (error) {
    // Sem falha fatal: se cache falhar, segue sem cache.
  }
}

function limparCacheValidacoes() {
  const key = getValidacoesCacheKey();
  try {
    CacheService.getScriptCache().remove(key);
    return { ok: true, key };
  } catch (error) {
    return { ok: false, key, erro: error.message };
  }
}

function recarregarCacheValidacoes() {
  limparCacheValidacoes();
  const dados = obterValidacoes(true);
  return {
    ok: true,
    key: getValidacoesCacheKey(),
    ttl_segundos: VALIDACOES_CACHE_TTL_SEC,
    tipos: Array.isArray(dados?.tipos) ? dados.tipos.length : 0,
    unidades: Array.isArray(dados?.unidades) ? dados.unidades.length : 0,
    categorias: Array.isArray(dados?.categorias) ? dados.categorias.length : 0,
    fornecedores: Array.isArray(dados?.fornecedores) ? dados.fornecedores.length : 0,
    categorias_despesas: Array.isArray(dados?.categoriasDespesas) ? dados.categoriasDespesas.length : 0,
    formas_pagamento: Array.isArray(dados?.formasPagamento) ? dados.formasPagamento.length : 0,
    pagos_por: Array.isArray(dados?.pagosPor) ? dados.pagosPor.length : 0
  };
}

function listarTipos() {
  const sheet = getSheet('VALIDACAO');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const dados = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  return dados
    .flat()
    .filter(v => v && v.toString().trim() !== '');
}

function listarUnidades() {
  const sheet = getSheet('VALIDACAO');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const dados = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

  return [...new Set(
    dados
      .flat()
      .filter(v => v && v.toString().trim() !== '')
  )];
}

function listarUnidadesPorTipo(tipo) {
  if (!tipo) return [];

  const sheet = getSheet('VALIDACAO');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const dados = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  return dados
    .filter(l => l[0] === tipo)
    .map(l => l[1])
    .filter(v => v && v.toString().trim() !== '');
}

/* =========================
   NOVA FUNÇÃO PADRÃO (GLOBAL)
   ========================= */

/**
 * Retorna TODAS as validações estruturadas
 * Esta será a função padrão para o web app
 */
function normalizarHeaderValidacao(valor) {
  return String(valor || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
}

function indiceColunaValidacao(headers, nomeEsperado, fallback) {
  const alvo = normalizarHeaderValidacao(nomeEsperado);
  const idx = headers.findIndex(h => normalizarHeaderValidacao(h) === alvo);
  return idx >= 0 ? idx : fallback;
}

function obterCategoriasPorTipoValidacao() {
  const sheet = getSheet('VALIDACAO_TIPO_CATEGORIA');
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  if (!Array.isArray(data) || data.length < 2) return {};

  const headers = Array.isArray(data[0]) ? data[0] : [];
  const rows = data.slice(1);
  const idxTipo = indiceColunaValidacao(headers, 'TIPO', 0);
  const idxCategoria = indiceColunaValidacao(headers, 'CATEGORIA', 1);

  const mapa = {};
  rows.forEach(row => {
    const tipo = String(row[idxTipo] || '').trim().toUpperCase();
    const categoria = String(row[idxCategoria] || '').trim();
    if (!tipo || !categoria) return;

    if (!mapa[tipo]) {
      mapa[tipo] = [];
    }
    const existe = mapa[tipo].some(c => String(c || '').trim().toUpperCase() === categoria.toUpperCase());
    if (!existe) {
      mapa[tipo].push(categoria);
    }
  });

  if (!Array.isArray(mapa.OUTROS) || mapa.OUTROS.length === 0) {
    mapa.OUTROS = ['OUTROS'];
  } else if (!mapa.OUTROS.some(c => String(c || '').trim().toUpperCase() === 'OUTROS')) {
    mapa.OUTROS.push('OUTROS');
  }

  return mapa;
}

function obterValidacoes(forcarRecarregar) {
  if (!forcarRecarregar) {
    const cached = lerValidacoesDoCache();
    if (cached) {
      return cached;
    }
  }

  const sheet = getSheet('VALIDACAO');
  const categoriasPorTipo = obterCategoriasPorTipoValidacao();
  let resultado;
  if (!sheet) {
    resultado = {
      tipos: [],
      unidades: [],
      categorias: [],
      categoriasDespesas: [],
      fornecedores: [],
      formasPagamento: [],
      pagosPor: [],
      valorKwhPorFornecedor: {},
      categoriasPorTipo
    };
    salvarValidacoesNoCache(resultado);
    return resultado;
  }
  const data = sheet.getDataRange().getValues();
  if (!Array.isArray(data) || data.length < 2) {
    resultado = {
      tipos: [],
      unidades: [],
      categorias: [],
      categoriasDespesas: [],
      fornecedores: [],
      formasPagamento: [],
      pagosPor: [],
      valorKwhPorFornecedor: {},
      categoriasPorTipo
    };
    salvarValidacoesNoCache(resultado);
    return resultado;
  }

  const headers = Array.isArray(data[0]) ? data[0] : [];
  const dados = data.slice(1);

  const idxTipo = indiceColunaValidacao(headers, 'TIPO', 0);
  const idxUnidade = indiceColunaValidacao(headers, 'UNIDADE', 1);
  const idxCategoria = indiceColunaValidacao(headers, 'CATEGORIA', 2);
  const idxFornecedor = indiceColunaValidacao(headers, 'FORNECEDOR', 3);
  const idxValorKwh = indiceColunaValidacao(headers, 'VALORKWH', 4);
  const idxDespesas = indiceColunaValidacao(headers, 'DESPESAS', 5);
  const idxFormaPagamento = indiceColunaValidacao(headers, 'FORMA_PAGAMENTO', headers.length > 6 ? 6 : 5);
  const idxPagoPor = indiceColunaValidacao(headers, 'PAGO_POR', 7);

  const tipos = new Set();
  const unidades = new Set();
  const categorias = new Set();
  const categoriasDespesas = new Set();
  const fornecedores = new Set();
  const formasPagamento = new Set();
  const pagosPor = new Set();
  const valorKwhPorFornecedor = {};

  dados.forEach(linha => {
    const tipo = linha[idxTipo];
    const unidade = linha[idxUnidade];
    const categoria = linha[idxCategoria];
    const fornecedor = linha[idxFornecedor];
    const valorKwh = linha[idxValorKwh];
    const categoriaDespesa = linha[idxDespesas];
    const formaPagamento = linha[idxFormaPagamento];
    const pagoPor = linha[idxPagoPor];

    if (tipo && tipo.toString().trim() !== '') {
      tipos.add(tipo);
    }

    if (unidade && unidade.toString().trim() !== '') {
      unidades.add(unidade);
    }

    if (categoria && categoria.toString().trim() !== '') {
      categorias.add(categoria);
    }

    if (fornecedor && fornecedor.toString().trim() !== '') {
      fornecedores.add(fornecedor);

      if (
        valorKwh !== '' &&
        valorKwh !== null &&
        !isNaN(valorKwh)
      ) {
        valorKwhPorFornecedor[fornecedor] = Number(valorKwh);
      }
    }

    if (formaPagamento && formaPagamento.toString().trim() !== '') {
      formasPagamento.add(formaPagamento);
    }

    if (categoriaDespesa && categoriaDespesa.toString().trim() !== '') {
      categoriasDespesas.add(categoriaDespesa);
    }

    if (pagoPor && pagoPor.toString().trim() !== '') {
      pagosPor.add(pagoPor);
    }
  });

  Object.keys(categoriasPorTipo).forEach(tipo => {
    (categoriasPorTipo[tipo] || []).forEach(cat => {
      if (cat && String(cat).trim() !== '') {
        categorias.add(cat);
      }
    });
  });

  resultado = {
    tipos: [...tipos],
    unidades: [...unidades],
    categorias: [...categorias],
    categoriasDespesas: [...categoriasDespesas],
    fornecedores: [...fornecedores],
    formasPagamento: [...formasPagamento],
    pagosPor: [...pagosPor],
    valorKwhPorFornecedor,
    categoriasPorTipo
  };
  salvarValidacoesNoCache(resultado);
  return resultado;
}

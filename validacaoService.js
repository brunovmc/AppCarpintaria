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
 * =========================
 */

/* =========================
   FUNÇÕES EXISTENTES (MANTIDAS)
   ========================= */

function testeValidacao() {
  Logger.log(obterValidacoes());
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

function obterValidacoes() {
  const sheet = getSheet('VALIDACAO');
  if (!sheet) {
    return {
      tipos: [],
      unidades: [],
      categorias: [],
      fornecedores: [],
      valorKwhPorFornecedor: {}
    };
  }
  const data = sheet.getDataRange().getValues();
  if (!Array.isArray(data) || data.length < 2) {
    return {
      tipos: [],
      unidades: [],
      categorias: [],
      fornecedores: [],
      valorKwhPorFornecedor: {}
    };
  }

  const headers = Array.isArray(data[0]) ? data[0] : [];
  const dados = data.slice(1);

  const idxTipo = indiceColunaValidacao(headers, 'TIPO', 0);
  const idxUnidade = indiceColunaValidacao(headers, 'UNIDADE', 1);
  const idxCategoria = indiceColunaValidacao(headers, 'CATEGORIA', 2);
  const idxFornecedor = indiceColunaValidacao(headers, 'FORNECEDOR', 3);
  const idxValorKwh = indiceColunaValidacao(headers, 'VALORKWH', 4);

  const tipos = new Set();
  const unidades = new Set();
  const categorias = new Set();
  const fornecedores = new Set();
  const valorKwhPorFornecedor = {};

  dados.forEach(linha => {
    const tipo = linha[idxTipo];
    const unidade = linha[idxUnidade];
    const categoria = linha[idxCategoria];
    const fornecedor = linha[idxFornecedor];
    const valorKwh = linha[idxValorKwh];

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
  });

  return {
    tipos: [...tipos],
    unidades: [...unidades],
    categorias: [...categorias],
    fornecedores: [...fornecedores],
    valorKwhPorFornecedor
  };
}

const ABA_COMPRAS = 'COMPRAS';
const COMPRAS_CACHE_SCOPE = 'COMPRAS_LISTA_ATIVAS_PENDENTES';
const COMPRAS_CACHE_TTL_SEC = 90;

const COMPRAS_SCHEMA = [
  'ID',
  'tipo',
  'item',
  'unidade',
  'valor_unit',
  'ativo',
  'criado_em',
  'quantidade',
  'comprimento_cm',
  'largura_cm',
  'espessura_cm',
  'categoria',
  'fornecedor',
  'potencia',
  'voltagem',
  'comprado_em',
  'vida_util_mes',
  'observacao',
  'adicionado_estoque',
  'estoque_id'
];

function categoriaValidaParaTipoCompra(tipo, categoria) {
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

function normalizarPayloadMadeiraCompra(payload) {
  const dados = { ...(payload || {}) };
  const tipo = String(dados.tipo || '').trim().toUpperCase();

  if (!categoriaValidaParaTipoCompra(tipo, dados.categoria)) {
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

function lerCacheListaCompras() {
  return appCacheGetJson(COMPRAS_CACHE_SCOPE);
}

function salvarCacheListaCompras(lista) {
  appCachePutJson(COMPRAS_CACHE_SCOPE, Array.isArray(lista) ? lista : [], COMPRAS_CACHE_TTL_SEC);
}

function limparCacheCompras() {
  return appCacheRemove(COMPRAS_CACHE_SCOPE);
}

function recarregarCacheCompras() {
  limparCacheCompras();
  const dados = listarCompras(true);
  return {
    ok: true,
    scope: COMPRAS_CACHE_SCOPE,
    ttl_segundos: COMPRAS_CACHE_TTL_SEC,
    total_itens: Array.isArray(dados) ? dados.length : 0
  };
}

function listarCompras(forcarRecarregar) {
  if (!forcarRecarregar) {
    const cached = lerCacheListaCompras();
    if (Array.isArray(cached)) {
      return cached;
    }
  }

  const sheet = getDataSpreadsheet().getSheetByName(ABA_COMPRAS);
  if (!sheet) {
    salvarCacheListaCompras([]);
    return [];
  }

  const rows = rowsToObjects(sheet);

  const lista = rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => String(i.adicionado_estoque).toLowerCase() !== 'true')
    .map(i => ({
      ...i,
      criado_em: i.criado_em
        ? Utilities.formatDate(new Date(i.criado_em), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
        : '',
      comprado_em: i.comprado_em
        ? Utilities.formatDate(new Date(i.comprado_em), Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : ''
    }));
  salvarCacheListaCompras(lista);
  return lista;
}

function criarItemCompra(payload) {
  const dados = normalizarPayloadMadeiraCompra(payload);
  const novo = {
    ...dados,
    ID: gerarId('COM'),
    ativo: true,
    criado_em: new Date(),
    adicionado_estoque: false
  };

  return insert(ABA_COMPRAS, novo, COMPRAS_SCHEMA);
}

function atualizarItemCompra(id, payload) {
  const dados = normalizarPayloadMadeiraCompra(payload);
  return updateById(
    ABA_COMPRAS,
    'ID',
    id,
    dados,
    COMPRAS_SCHEMA
  );
}

function deletarItemCompra(id) {
  return updateById(
    ABA_COMPRAS,
    'ID',
    id,
    { ativo: false },
    COMPRAS_SCHEMA
  );
}

function adicionarCompraAoEstoque(compraId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheetCompras = getSheet(ABA_COMPRAS);
    if (!sheetCompras) {
      throw new Error('Aba COMPRAS nao encontrada');
    }

    const compras = rowsToObjects(sheetCompras);
    const compra = compras.find(i => i.ID === compraId);

    if (!compra) {
      throw new Error('Compra nao encontrada');
    }

    const ativo = String(compra.ativo).toLowerCase() === 'true';
    const jaAdicionada = String(compra.adicionado_estoque).toLowerCase() === 'true';

    if (!ativo) {
      throw new Error('Compra inativa');
    }

    if (jaAdicionada) {
      throw new Error('Compra ja adicionada ao estoque');
    }

    const compraNormalizada = normalizarPayloadMadeiraCompra(compra);

    const itemEstoqueCriado = {
      ID: gerarId('EST'),
      tipo: compraNormalizada.tipo || '',
      item: compraNormalizada.item || '',
      unidade: compraNormalizada.unidade || '',
      valor_unit: compraNormalizada.valor_unit || 0,
      ativo: true,
      criado_em: new Date(),
      quantidade: compraNormalizada.quantidade || 0,
      comprimento_cm: compraNormalizada.comprimento_cm || '',
      largura_cm: compraNormalizada.largura_cm || '',
      espessura_cm: compraNormalizada.espessura_cm || '',
      categoria: compraNormalizada.categoria || '',
      fornecedor: compraNormalizada.fornecedor || '',
      potencia: compraNormalizada.potencia || '',
      voltagem: compraNormalizada.voltagem || '',
      comprado_em: compraNormalizada.comprado_em || '',
      vida_util_mes: compraNormalizada.vida_util_mes || '',
      observacao: compraNormalizada.observacao || ''
    };

    insert(ABA_ESTOQUE, itemEstoqueCriado, ESTOQUE_SCHEMA);

    const compraAtualizada = {
      adicionado_estoque: true,
      estoque_id: itemEstoqueCriado.ID
    };

    updateById(
      ABA_COMPRAS,
      'ID',
      compraId,
      compraAtualizada,
      COMPRAS_SCHEMA
    );

    return {
      compraAtualizada: {
        ...compra,
        ...compraAtualizada
      },
      itemEstoqueCriado: {
        ...itemEstoqueCriado,
        criado_em: Utilities.formatDate(new Date(itemEstoqueCriado.criado_em), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
        comprado_em: itemEstoqueCriado.comprado_em
          ? Utilities.formatDate(new Date(itemEstoqueCriado.comprado_em), Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : ''
      }
    };
  } finally {
    lock.releaseLock();
  }
}

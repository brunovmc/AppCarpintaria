const ABA_PRODUTOS = 'PRODUTOS';
const ABA_PRODUTOS_COMPONENTES = 'PRODUTOS_COMPONENTES';
const ABA_PRODUTOS_ETAPAS = 'PRODUTOS_ETAPAS';

const PRODUTOS_SCHEMA = [
  'produto_id',
  'nome_produto',
  'unidade_produto',
  'ativo',
  'criado_em'
];

const PRODUTOS_COMPONENTES_SCHEMA = [
  'id',
  'produto_id',
  'tipo_componente',
  'ref_id',
  'quantidade',
  'unidade',
  'ativo'
];

const PRODUTOS_ETAPAS_SCHEMA = [
  'id',
  'produto_id',
  'nome_etapa',
  'ordem',
  'ativo'
];

function listarProdutos() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS);
  if (!sheet) return [];

  const rows = rowsToObjects(sheet);
  return rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .map(i => ({
      ...i,
      criado_em: i.criado_em
        ? Utilities.formatDate(new Date(i.criado_em), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
        : ''
    }));
}

function criarProduto(payload) {
  const novo = {
    ...payload,
    produto_id: gerarId('PRD'),
    ativo: true,
    criado_em: new Date()
  };

  insert(ABA_PRODUTOS, novo, PRODUTOS_SCHEMA);

  return {
    ...novo,
    criado_em: Utilities.formatDate(novo.criado_em, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
  };
}

function atualizarProduto(produtoId, payload) {
  return updateById(
    ABA_PRODUTOS,
    'produto_id',
    produtoId,
    payload,
    PRODUTOS_SCHEMA
  );
}

function deletarProduto(produtoId) {
  return updateById(
    ABA_PRODUTOS,
    'produto_id',
    produtoId,
    { ativo: false },
    PRODUTOS_SCHEMA
  );
}

function listarComponentesProduto(produtoId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_COMPONENTES);
  if (!sheet) return [];

  const rows = rowsToObjects(sheet);
  return rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.produto_id === produtoId);
}

function criarComponenteProduto(payload) {
  const novo = {
    ...payload,
    id: gerarId('CMP'),
    ativo: true
  };

  insert(ABA_PRODUTOS_COMPONENTES, novo, PRODUTOS_COMPONENTES_SCHEMA);
  return novo;
}

function atualizarComponenteProduto(id, payload) {
  return updateById(
    ABA_PRODUTOS_COMPONENTES,
    'id',
    id,
    payload,
    PRODUTOS_COMPONENTES_SCHEMA
  );
}

function deletarComponenteProduto(id) {
  return updateById(
    ABA_PRODUTOS_COMPONENTES,
    'id',
    id,
    { ativo: false },
    PRODUTOS_COMPONENTES_SCHEMA
  );
}

function listarEtapasProduto(produtoId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_ETAPAS);
  if (!sheet) return [];

  const rows = rowsToObjects(sheet);
  return rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.produto_id === produtoId)
    .sort((a, b) => parseNumeroBR(a.ordem) - parseNumeroBR(b.ordem));
}

function criarEtapaProduto(payload) {
  const novo = {
    ...payload,
    id: gerarId('ETP'),
    ativo: true
  };

  insert(ABA_PRODUTOS_ETAPAS, novo, PRODUTOS_ETAPAS_SCHEMA);
  return novo;
}

function atualizarEtapaProduto(id, payload) {
  return updateById(
    ABA_PRODUTOS_ETAPAS,
    'id',
    id,
    payload,
    PRODUTOS_ETAPAS_SCHEMA
  );
}

function deletarEtapaProduto(id) {
  return updateById(
    ABA_PRODUTOS_ETAPAS,
    'id',
    id,
    { ativo: false },
    PRODUTOS_ETAPAS_SCHEMA
  );
}

function explodirBOM(produtoId, qtd) {
  const sheetComponentes = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_COMPONENTES);
  if (!sheetComponentes) return { itens: [], custoPrevisto: 0 };

  const allComponentes = rowsToObjects(sheetComponentes)
    .filter(i => String(i.ativo).toLowerCase() === 'true');

  const componentesPorProduto = {};
  allComponentes.forEach(c => {
    if (!componentesPorProduto[c.produto_id]) {
      componentesPorProduto[c.produto_id] = [];
    }
    componentesPorProduto[c.produto_id].push(c);
  });

  const sheetEstoque = SpreadsheetApp.getActive().getSheetByName(ABA_ESTOQUE);
  const estoqueRows = sheetEstoque ? rowsToObjects(sheetEstoque) : [];
  const estoqueMap = {};
  estoqueRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(i => {
      estoqueMap[i.ID] = i;
    });

  const resultado = {};
  let custoPrevisto = 0;
  const visitados = {};

  function adicionarEstoque(refId, qtdTotal, unidadeComp) {
    const estoque = estoqueMap[refId] || {};
    const valorUnit = parseNumeroBR(estoque.valor_unit);
    const itemNome = estoque.item || refId || '';
    const unidade = estoque.unidade || unidadeComp || '';

    if (!resultado[refId]) {
      resultado[refId] = {
        estoque_id: refId,
        item: itemNome,
        unidade,
        quantidade: 0,
        valor_unit: valorUnit
      };
    }

    resultado[refId].quantidade += qtdTotal;
    custoPrevisto += qtdTotal * valorUnit;
  }

  function walk(prodId, qtdBase) {
    if (!prodId) return;
    if (visitados[prodId]) {
      throw new Error('Loop detectado na composicao do produto');
    }

    visitados[prodId] = true;
    const componentes = componentesPorProduto[prodId] || [];

    componentes.forEach(c => {
      const qtdComp = parseNumeroBR(c.quantidade) * qtdBase;
      if (c.tipo_componente === 'ESTOQUE') {
        adicionarEstoque(c.ref_id, qtdComp, c.unidade);
        return;
      }

      if (c.tipo_componente === 'PRODUTO') {
        walk(c.ref_id, qtdComp);
      }
    });

    visitados[prodId] = false;
  }

  const qtdBase = parseNumeroBR(qtd);
  if (qtdBase > 0) {
    walk(produtoId, qtdBase);
  }

  return {
    itens: Object.values(resultado),
    custoPrevisto
  };
}

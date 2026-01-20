const ABA_PRODUTOS = 'PRODUTOS';
const ABA_PRODUTOS_COMPONENTES = 'PRODUTOS_COMPONENTES';
const ABA_PRODUTOS_ETAPAS = 'PRODUTOS_ETAPAS';

const PRODUTOS_SCHEMA = [
  'produto_id',
  'nome_produto',
  'unidade_produto',
  'descricao',
  'observacao',
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
  'observacao',
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
    .map(i => {
      let custoPrevisto = 0;
      let custoErro = '';

      try {
        custoPrevisto = calcularCustoProduto(i.produto_id);
      } catch (err) {
        custoPrevisto = 0;
        custoErro = err && err.message ? err.message : 'Erro ao calcular custo';
      }

      return {
        ...i,
        custo_previsto: custoPrevisto,
        custo_erro: custoErro,
        criado_em: i.criado_em
          ? Utilities.formatDate(new Date(i.criado_em), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
          : ''
      };
    });
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

function obterProduto(produtoId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS);
  if (!sheet) return null;

  const rows = rowsToObjects(sheet);
  return rows.find(i => i.produto_id === produtoId) || null;
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

function listarComposicaoProduto(produtoId) {
  const rows = listarComponentesProduto(produtoId);
  return rows.map(i => ({
    id: i.id,
    produto_id: i.produto_id,
    tipo_item: String(i.tipo_componente || '').toUpperCase(),
    item_id: i.ref_id,
    quantidade: parseNumeroBR(i.quantidade),
    unidade: i.unidade || '',
    observacao: i.observacao || ''
  }));
}

function validarCicloComposicao(produtoId, linhas) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_COMPONENTES);
  const rows = sheet ? rowsToObjects(sheet) : [];

  const ativos = rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.produto_id !== produtoId);

  const novas = (Array.isArray(linhas) ? linhas : []).map(l => ({
    produto_id: produtoId,
    tipo_componente: String(l.tipo_item || l.tipo_componente || '').toUpperCase(),
    ref_id: l.item_id || l.ref_id || ''
  }));

  const all = ativos.concat(novas);
  const mapa = {};

  all.forEach(c => {
    if (c.tipo_componente !== 'PRODUTO') return;
    if (!mapa[c.produto_id]) mapa[c.produto_id] = [];
    mapa[c.produto_id].push(c.ref_id);
  });

  const visitados = {};
  function dfs(id) {
    if (visitados[id]) {
      throw new Error('Loop detectado na composicao do produto');
    }
    visitados[id] = true;
    (mapa[id] || []).forEach(dfs);
    visitados[id] = false;
  }

  dfs(produtoId);
  return true;
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

function salvarComposicaoProduto(produtoId, linhas) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(ABA_PRODUTOS_COMPONENTES);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUTOS_COMPONENTES);
  }

  ensureSchema(sheet, PRODUTOS_COMPONENTES_SCHEMA);

  validarCicloComposicao(produtoId, linhas);

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idCol = headers.indexOf('produto_id');

  if (idCol !== -1 && data.length > 1) {
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][idCol] === produtoId) {
        sheet.deleteRow(i + 1);
      }
    }
  }

  const linhasValidas = Array.isArray(linhas) ? linhas : [];
  linhasValidas.forEach(l => {
    const novo = {
      id: gerarId('CMP'),
      produto_id: produtoId,
      tipo_componente: String(l.tipo_item || l.tipo_componente || '').toUpperCase(),
      ref_id: l.item_id || l.ref_id || '',
      quantidade: parseNumeroBR(l.quantidade),
      unidade: l.unidade || '',
      observacao: l.observacao || '',
      ativo: true
    };
    insert(ABA_PRODUTOS_COMPONENTES, novo, PRODUTOS_COMPONENTES_SCHEMA);
  });

  return true;
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

function calcularCustoProduto(produtoId) {
  if (!produtoId) return 0;
  const resp = explodirBOM(produtoId, 1);
  return parseNumeroBR(resp.custoPrevisto);
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

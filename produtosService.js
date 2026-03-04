const ABA_PRODUTOS = 'PRODUTOS';
const ABA_PRODUTOS_COMPONENTES = 'PRODUTOS_COMPONENTES';
const ABA_PRODUTOS_ETAPAS = 'PRODUTOS_ETAPAS';
const ABA_PRODUTOS_RECEITAS = 'PRODUTOS_RECEITAS';
const ABA_PRODUTOS_RECEITAS_ENTRADAS = 'PRODUTOS_RECEITAS_ENTRADAS';
const ABA_PRODUTOS_RECEITAS_SAIDAS = 'PRODUTOS_RECEITAS_SAIDAS';

const PRODUTOS_SCHEMA = [
  'produto_id',
  'nome_produto',
  'unidade_produto',
  'estoque_id',
  'categoria',
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

const PRODUTOS_RECEITAS_SCHEMA = [
  'receita_id',
  'produto_id',
  'nome_receita',
  'descricao',
  'parent_receita_id',
  'ativo',
  'criado_em'
];

const PRODUTOS_RECEITAS_ENTRADAS_SCHEMA = [
  'id',
  'receita_id',
  'tipo_item',
  'nome_item',
  'estoque_ref_id',
  'produto_ref_id',
  'receita_ref_id',
  'categoria',
  'unidade',
  'qtd_pecas',
  'comprimento_cm',
  'largura_cm',
  'espessura_cm',
  'custo_manual',
  'observacao',
  'ativo'
];

const PRODUTOS_RECEITAS_SAIDAS_SCHEMA = [
  'id',
  'receita_id',
  'nome_saida',
  'unidade',
  'quantidade',
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

function adicionarMaterialProduto(produtoId, estoqueId, quantidade) {
  if (!produtoId || !estoqueId) {
    throw new Error('Produto ou estoque invalido');
  }

  const qtd = parseNumeroBR(quantidade);
  if (!qtd || qtd <= 0) {
    throw new Error('Quantidade invalida');
  }

  const sheetEstoque = SpreadsheetApp.getActive().getSheetByName('ESTOQUE');
  if (!sheetEstoque) {
    throw new Error('Aba ESTOQUE nao encontrada');
  }

  const estoqueRows = rowsToObjects(sheetEstoque);
  const estoqueItem = estoqueRows.find(i => i.ID === estoqueId) || {};
  const unidade = estoqueItem.unidade || '';

  const componente = {
    id: gerarId('CMP'),
    produto_id: produtoId,
    tipo_componente: 'ESTOQUE',
    ref_id: estoqueId,
    quantidade: qtd,
    unidade,
    observacao: '',
    ativo: true
  };

  insert(ABA_PRODUTOS_COMPONENTES, componente, PRODUTOS_COMPONENTES_SCHEMA);

  return {
    componente: {
      id: componente.id,
      produto_id: produtoId,
      tipo_item: 'ESTOQUE',
      item_id: estoqueId,
      quantidade: qtd,
      unidade,
      observacao: ''
    },
    custoPrevisto: calcularCustoProduto(produtoId)
  };
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

function listarReceitasProduto(produtoId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (!sheet) return [];

  const receitas = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.produto_id === produtoId);

  const sheetEntradas = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_RECEITAS_ENTRADAS);
  const sheetSaidas = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_RECEITAS_SAIDAS);

  const entradas = sheetEntradas ? rowsToObjects(sheetEntradas) : [];
  const saidas = sheetSaidas ? rowsToObjects(sheetSaidas) : [];

  const entradasAtivas = entradas.filter(i => String(i.ativo).toLowerCase() === 'true');
  const saidasAtivas = saidas.filter(i => String(i.ativo).toLowerCase() === 'true');

  return receitas.map(r => ({
    ...r,
    criado_em: r.criado_em
      ? Utilities.formatDate(new Date(r.criado_em), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
      : '',
    entradas: entradasAtivas.filter(e => e.receita_id === r.receita_id),
    saidas: saidasAtivas.filter(s => s.receita_id === r.receita_id)
  }));
}

function criarReceitaProduto(produtoId, payload) {
  const novo = {
    receita_id: gerarId('REC'),
    produto_id: produtoId,
    nome_receita: payload?.nome_receita || 'Nova receita',
    descricao: payload?.descricao || '',
    parent_receita_id: payload?.parent_receita_id || '',
    ativo: true,
    criado_em: new Date()
  };

  insert(ABA_PRODUTOS_RECEITAS, novo, PRODUTOS_RECEITAS_SCHEMA);

  return {
    ...novo,
    criado_em: Utilities.formatDate(novo.criado_em, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
  };
}

function atualizarReceitaProduto(receitaId, payload) {
  return updateById(
    ABA_PRODUTOS_RECEITAS,
    'receita_id',
    receitaId,
    payload,
    PRODUTOS_RECEITAS_SCHEMA
  );
}

function inativarLinhasReceita(sheetName, receitaId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idCol = headers.indexOf('receita_id');
  const ativoCol = headers.indexOf('ativo');

  if (idCol === -1 || ativoCol === -1) return;

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === receitaId) {
      sheet.getRange(i + 1, ativoCol + 1).setValue(false);
    }
  }
}

function deletarReceitaProduto(receitaId) {
  const ok = updateById(
    ABA_PRODUTOS_RECEITAS,
    'receita_id',
    receitaId,
    { ativo: false },
    PRODUTOS_RECEITAS_SCHEMA
  );

  inativarLinhasReceita(ABA_PRODUTOS_RECEITAS_ENTRADAS, receitaId);
  inativarLinhasReceita(ABA_PRODUTOS_RECEITAS_SAIDAS, receitaId);

  return ok;
}

function limparLinhasReceita(sheetName, receitaId) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idCol = headers.indexOf('receita_id');

  if (idCol !== -1 && data.length > 1) {
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][idCol] === receitaId) {
        sheet.deleteRow(i + 1);
      }
    }
  }
}

function salvarEntradasReceita(receitaId, linhas) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(ABA_PRODUTOS_RECEITAS_ENTRADAS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUTOS_RECEITAS_ENTRADAS);
  }

  ensureSchema(sheet, PRODUTOS_RECEITAS_ENTRADAS_SCHEMA);
  limparLinhasReceita(ABA_PRODUTOS_RECEITAS_ENTRADAS, receitaId);

  const linhasValidas = Array.isArray(linhas) ? linhas : [];
  linhasValidas.forEach(l => {
    const novo = {
      id: gerarId('REN'),
      receita_id: receitaId,
      tipo_item: String(l.tipo_item || '').toUpperCase(),
      nome_item: l.nome_item || '',
      estoque_ref_id: l.estoque_ref_id || '',
      produto_ref_id: l.produto_ref_id || '',
      receita_ref_id: l.receita_ref_id || '',
      categoria: l.categoria || '',
      unidade: l.unidade || '',
      qtd_pecas: parseNumeroBR(l.qtd_pecas),
      comprimento_cm: parseNumeroBR(l.comprimento_cm),
      largura_cm: parseNumeroBR(l.largura_cm),
      espessura_cm: parseNumeroBR(l.espessura_cm),
      custo_manual: parseNumeroBR(l.custo_manual),
      observacao: l.observacao || '',
      ativo: true
    };
    insert(ABA_PRODUTOS_RECEITAS_ENTRADAS, novo, PRODUTOS_RECEITAS_ENTRADAS_SCHEMA);
  });

  return true;
}

function salvarSaidasReceita(receitaId, linhas) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(ABA_PRODUTOS_RECEITAS_SAIDAS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUTOS_RECEITAS_SAIDAS);
  }

  ensureSchema(sheet, PRODUTOS_RECEITAS_SAIDAS_SCHEMA);
  limparLinhasReceita(ABA_PRODUTOS_RECEITAS_SAIDAS, receitaId);

  const linhasValidas = Array.isArray(linhas) ? linhas : [];
  linhasValidas.forEach(l => {
    const novo = {
      id: gerarId('RSA'),
      receita_id: receitaId,
      nome_saida: l.nome_saida || '',
      unidade: l.unidade || '',
      quantidade: parseNumeroBR(l.quantidade),
      ativo: true
    };
    insert(ABA_PRODUTOS_RECEITAS_SAIDAS, novo, PRODUTOS_RECEITAS_SAIDAS_SCHEMA);
  });

  return true;
}

function duplicarReceitaProduto(receitaId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (!sheet) return null;

  const receitas = rowsToObjects(sheet);
  const receita = receitas.find(r => r.receita_id === receitaId);
  if (!receita) return null;

  const sheetEntradas = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_RECEITAS_ENTRADAS);
  const sheetSaidas = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_RECEITAS_SAIDAS);

  const entradas = sheetEntradas ? rowsToObjects(sheetEntradas) : [];
  const saidas = sheetSaidas ? rowsToObjects(sheetSaidas) : [];

  const entradasAtivas = entradas
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.receita_id === receitaId);
  const saidasAtivas = saidas
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.receita_id === receitaId);

  const novo = {
    receita_id: gerarId('REC'),
    produto_id: receita.produto_id,
    nome_receita: `${receita.nome_receita || 'Receita'} (Copia)`,
    descricao: receita.descricao || '',
    parent_receita_id: receita.parent_receita_id || receita.receita_id,
    ativo: true,
    criado_em: new Date()
  };

  insert(ABA_PRODUTOS_RECEITAS, novo, PRODUTOS_RECEITAS_SCHEMA);

  entradasAtivas.forEach(l => {
    const novoItem = {
      id: gerarId('REN'),
      receita_id: novo.receita_id,
      tipo_item: String(l.tipo_item || '').toUpperCase(),
      nome_item: l.nome_item || '',
      estoque_ref_id: l.estoque_ref_id || '',
      produto_ref_id: l.produto_ref_id || '',
      receita_ref_id: l.receita_ref_id || '',
      categoria: l.categoria || '',
      unidade: l.unidade || '',
      qtd_pecas: parseNumeroBR(l.qtd_pecas),
      comprimento_cm: parseNumeroBR(l.comprimento_cm),
      largura_cm: parseNumeroBR(l.largura_cm),
      espessura_cm: parseNumeroBR(l.espessura_cm),
      custo_manual: parseNumeroBR(l.custo_manual),
      observacao: l.observacao || '',
      ativo: true
    };
    insert(ABA_PRODUTOS_RECEITAS_ENTRADAS, novoItem, PRODUTOS_RECEITAS_ENTRADAS_SCHEMA);
  });

  saidasAtivas.forEach(l => {
    const novoItem = {
      id: gerarId('RSA'),
      receita_id: novo.receita_id,
      nome_saida: l.nome_saida || '',
      unidade: l.unidade || '',
      quantidade: parseNumeroBR(l.quantidade),
      ativo: true
    };
    insert(ABA_PRODUTOS_RECEITAS_SAIDAS, novoItem, PRODUTOS_RECEITAS_SAIDAS_SCHEMA);
  });

  return {
    ...novo,
    criado_em: Utilities.formatDate(novo.criado_em, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
  };
}

function salvarReceitaCompleta(receitaId, dados, entradas, saidas) {
  atualizarReceitaProduto(receitaId, dados || {});
  salvarEntradasReceita(receitaId, entradas || []);
  salvarSaidasReceita(receitaId, saidas || []);
  return true;
}

function obterReceitaPadraoProduto(produtoId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (!sheet) return null;

  const receitas = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.produto_id === produtoId)
    .filter(i => i.parent_receita_id);

  return receitas[0] || null;
}

function escolherSaidaPrincipal(saidas, nomeProduto) {
  if (!Array.isArray(saidas) || saidas.length === 0) return null;

  const nome = String(nomeProduto || '').trim().toLowerCase();
  if (nome) {
    const match = saidas.find(s => String(s.nome_saida || '').trim().toLowerCase() === nome);
    if (match) return match;
  }

  let escolhida = saidas[0];
  let maior = parseNumeroBR(escolhida.quantidade);
  for (let i = 1; i < saidas.length; i++) {
    const qtd = parseNumeroBR(saidas[i].quantidade);
    if (qtd > maior) {
      maior = qtd;
      escolhida = saidas[i];
    }
  }
  return escolhida;
}

function explodirReceita(produtoId, receitaId, qtdPlanejada) {
  if (!produtoId || !receitaId) return { itens: [], custoPrevisto: 0 };

  const sheetReceitas = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (!sheetReceitas) return { itens: [], custoPrevisto: 0 };

  const receitas = rowsToObjects(sheetReceitas)
    .filter(i => String(i.ativo).toLowerCase() === 'true');
  const receita = receitas.find(r => r.receita_id === receitaId && r.produto_id === produtoId);
  if (!receita) {
    throw new Error('Receita nao encontrada');
  }
  if (!receita.parent_receita_id) {
    throw new Error('Receita base nao pode ser usada');
  }

  const sheetEntradas = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_RECEITAS_ENTRADAS);
  const sheetSaidas = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_RECEITAS_SAIDAS);
  const entradas = sheetEntradas ? rowsToObjects(sheetEntradas) : [];
  const saidas = sheetSaidas ? rowsToObjects(sheetSaidas) : [];

  const entradasAtivas = entradas
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.receita_id === receitaId);
  const saidasAtivas = saidas
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.receita_id === receitaId);

  const produtosSheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS);
  const produtosRows = produtosSheet ? rowsToObjects(produtosSheet) : [];
  const produtosMap = {};
  produtosRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(p => {
      produtosMap[p.produto_id] = p;
    });

  const produtoRef = produtosMap[produtoId] || {};
  const produtoNome = produtoRef.nome_produto || '';
  const saidaPrincipal = escolherSaidaPrincipal(saidasAtivas, produtoNome);
  if (!saidaPrincipal) {
    throw new Error('Receita sem saida principal');
  }

  const qtdSaida = parseNumeroBR(saidaPrincipal.quantidade);
  if (!qtdSaida || qtdSaida <= 0) {
    throw new Error('Saida principal invalida');
  }

  const qtdPlanejadaNum = parseNumeroBR(qtdPlanejada);
  if (!qtdPlanejadaNum || qtdPlanejadaNum <= 0) {
    return { itens: [], custoPrevisto: 0 };
  }

  const fator = qtdPlanejadaNum / qtdSaida;

  const sheetEstoque = SpreadsheetApp.getActive().getSheetByName(ABA_ESTOQUE);
  const estoqueRows = sheetEstoque ? rowsToObjects(sheetEstoque) : [];
  const estoqueMap = {};
  estoqueRows.forEach(i => {
    estoqueMap[i.ID] = i;
  });

  const resultado = {};
  let custoPrevisto = 0;

  entradasAtivas.forEach(e => {
    const tipo = String(e.tipo_item || '').toUpperCase();
    if (tipo === 'LIVRE') return;

    let estoqueId = '';
    let itemNome = '';
    let unidade = '';
    let valorUnit = 0;
    let quantidade = 0;

    const qtdBase = parseNumeroBR(e.qtd_pecas);
    if (!qtdBase || qtdBase <= 0) {
      throw new Error(`Qtd de pecas invalida: ${e.nome_item || e.estoque_ref_id || e.produto_ref_id || ''}`);
    }

    if (tipo === 'PRODUTO') {
      const prodId = e.produto_ref_id || '';
      if (!prodId) {
        throw new Error('Entrada de produto sem referencia');
      }

      const prod = produtosMap[prodId] || null;
      if (!prod) {
        throw new Error(`Produto nao encontrado: ${prodId}`);
      }

      estoqueId = prod.estoque_id || '';
      if (!estoqueId) {
        throw new Error(`Produto sem estoque vinculado: ${prod.nome_produto || prodId}`);
      }

      const estoqueItem = estoqueMap[estoqueId];
      if (!estoqueItem) {
        throw new Error(`Item de estoque nao encontrado: ${estoqueId}`);
      }
      if (String(estoqueItem.ativo).toLowerCase() !== 'true') {
        throw new Error(`Item de estoque inativo: ${estoqueId}`);
      }
      if (String(estoqueItem.tipo || '').toUpperCase() !== 'PRODUTO') {
        throw new Error(`Item de estoque do produto deve ser do tipo PRODUTO: ${estoqueId}`);
      }

      itemNome = estoqueItem.item || prod.nome_produto || estoqueId;
      unidade = estoqueItem.unidade || '';
      valorUnit = parseNumeroBR(estoqueItem.valor_unit);
      quantidade = qtdBase * fator;
    } else {
      estoqueId = e.estoque_ref_id || '';
      if (!estoqueId) {
        throw new Error(`Item de estoque nao informado: ${e.nome_item || ''}`);
      }

      const estoqueItem = estoqueMap[estoqueId];
      if (!estoqueItem) {
        throw new Error(`Item de estoque nao encontrado: ${estoqueId}`);
      }
      if (String(estoqueItem.ativo).toLowerCase() !== 'true') {
        throw new Error(`Item de estoque inativo: ${estoqueId}`);
      }

      itemNome = estoqueItem.item || e.nome_item || estoqueId;
      unidade = estoqueItem.unidade || e.unidade || '';
      valorUnit = parseNumeroBR(estoqueItem.valor_unit);

      if (tipo === 'MADEIRA') {
        const comp = parseNumeroBR(e.comprimento_cm);
        const larg = parseNumeroBR(e.largura_cm);
        const esp = parseNumeroBR(e.espessura_cm);
        if (!comp || !larg || !esp) {
          throw new Error(`Madeira sem medidas: ${itemNome}`);
        }
        const volumeM3 = (comp * larg * esp) / 1000000;
        quantidade = qtdBase * volumeM3 * fator;
        unidade = 'M3';
      } else {
        quantidade = qtdBase * fator;
      }
    }

    if (!estoqueId || !quantidade || quantidade <= 0) return;

    if (!resultado[estoqueId]) {
      resultado[estoqueId] = {
        estoque_id: estoqueId,
        item: itemNome,
        unidade,
        quantidade: 0,
        valor_unit: valorUnit
      };
    }

    resultado[estoqueId].quantidade += quantidade;
    custoPrevisto += quantidade * valorUnit;
  });

  return {
    itens: Object.values(resultado),
    custoPrevisto
  };
}

function calcularCustoProduto(produtoId) {
  if (!produtoId) return 0;
  const receita = obterReceitaPadraoProduto(produtoId);
  if (!receita) {
    throw new Error('Produto sem receita de producao');
  }
  const resp = explodirReceita(produtoId, receita.receita_id, 1);
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

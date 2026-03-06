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

function getResumoModeloProduto(produtoId) {
  const modelo = obterModeloProduto(produtoId);
  const entradasTotal = Array.isArray(modelo?.entradas) ? modelo.entradas.length : 0;
  const saidasTotal = Array.isArray(modelo?.saidas) ? modelo.saidas.length : 0;
  const etapasTotal = Array.isArray(modelo?.etapas) ? modelo.etapas.length : 0;

  return {
    modelo_entradas_total: entradasTotal,
    modelo_saidas_total: saidasTotal,
    modelo_etapas_total: etapasTotal,
    modelo_resumo: `${entradasTotal} entradas, ${saidasTotal} saidas, ${etapasTotal} etapas`
  };
}

function listarProdutos() {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS);
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

      const resumoModelo = getResumoModeloProduto(i.produto_id);

      return {
        ...i,
        ...resumoModelo,
        custo_previsto: custoPrevisto,
        custo_erro: custoErro,
        criado_em: i.criado_em
          ? Utilities.formatDate(new Date(i.criado_em), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
          : ''
      };
    });
}

function criarProduto(payload) {
  const nome = String(payload?.nome_produto || '').trim();
  if (!nome) {
    throw new Error('Nome do produto obrigatorio');
  }

  const novo = {
    nome_produto: nome,
    unidade_produto: String(payload?.unidade_produto || 'UN').trim() || 'UN',
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
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS);
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
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_COMPONENTES);
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
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_COMPONENTES);
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
  const ss = getDataSpreadsheet();
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

  const sheetEstoque = getDataSpreadsheet().getSheetByName('ESTOQUE');
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
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_ETAPAS);
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

function limparEtapasProduto(produtoId) {
  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(ABA_PRODUTOS_ETAPAS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUTOS_ETAPAS);
  }

  ensureSchema(sheet, PRODUTOS_ETAPAS_SCHEMA);

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idxProduto = headers.indexOf('produto_id');

  if (idxProduto === -1 || data.length <= 1) return;

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idxProduto] || '') === String(produtoId || '')) {
      sheet.deleteRow(i + 1);
    }
  }
}

function salvarEtapasProduto(produtoId, etapas) {
  if (!produtoId) throw new Error('Produto invalido');

  limparEtapasProduto(produtoId);

  const linhas = Array.isArray(etapas) ? etapas : [];
  let ordemSeq = 1;

  linhas.forEach(et => {
    const nomeEtapa = String(et?.nome_etapa || '').trim();
    if (!nomeEtapa) return;

    const ordemInformada = parseNumeroBR(et?.ordem);
    const ordem = ordemInformada > 0 ? ordemInformada : ordemSeq;

    const novo = {
      id: gerarId('ETP'),
      produto_id: produtoId,
      nome_etapa: nomeEtapa,
      ordem,
      ativo: true
    };

    insert(ABA_PRODUTOS_ETAPAS, novo, PRODUTOS_ETAPAS_SCHEMA);
    ordemSeq += 1;
  });

  return true;
}

function listarReceitasProduto(produtoId) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (!sheet) return [];

  const receitas = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.produto_id === produtoId);

  const sheetEntradas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS_ENTRADAS);
  const sheetSaidas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS_SAIDAS);

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
  if (!produtoId) {
    throw new Error('Produto invalido');
  }

  const receitasAtuais = listarReceitasProduto(produtoId);
  if (receitasAtuais.length > 0) {
    return receitasAtuais[0];
  }

  const novo = {
    receita_id: gerarId('REC'),
    produto_id: produtoId,
    nome_receita: payload?.nome_receita || 'Modelo principal',
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
  const sheet = getDataSpreadsheet().getSheetByName(sheetName);
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
  const ss = getDataSpreadsheet();
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
  const ss = getDataSpreadsheet();
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
  const ss = getDataSpreadsheet();
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
  const _ = receitaId;
  throw new Error('Duplicacao desativada: cada produto possui apenas um modelo');
}

function salvarReceitaCompleta(receitaId, dados, entradas, saidas) {
  atualizarReceitaProduto(receitaId, dados || {});
  salvarEntradasReceita(receitaId, entradas || []);
  salvarSaidasReceita(receitaId, saidas || []);
  return true;
}

function inativarReceitasSecundariasProduto(produtoId, receitaPrincipalId) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (!sheet) return true;

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idxProduto = headers.indexOf('produto_id');
  const idxReceita = headers.indexOf('receita_id');
  const idxAtivo = headers.indexOf('ativo');

  if (idxProduto === -1 || idxReceita === -1 || idxAtivo === -1) return true;

  for (let i = 1; i < data.length; i++) {
    const rowProdutoId = String(data[i][idxProduto] || '');
    const rowReceitaId = String(data[i][idxReceita] || '');
    if (rowProdutoId !== String(produtoId || '')) continue;
    if (rowReceitaId === String(receitaPrincipalId || '')) continue;
    sheet.getRange(i + 1, idxAtivo + 1).setValue(false);
  }

  return true;
}

function obterModeloProduto(produtoId) {
  if (!produtoId) {
    return {
      receita_id: '',
      dados: { nome_receita: 'Modelo principal', descricao: '' },
      entradas: [],
      saidas: [],
      etapas: []
    };
  }

  const receitas = listarReceitasProduto(produtoId);
  const receita = Array.isArray(receitas) && receitas.length > 0 ? receitas[0] : null;
  const etapas = listarEtapasProduto(produtoId);

  if (!receita) {
    return {
      receita_id: '',
      dados: { nome_receita: 'Modelo principal', descricao: '' },
      entradas: [],
      saidas: [],
      etapas
    };
  }

  return {
    receita_id: receita.receita_id || '',
    dados: {
      nome_receita: receita.nome_receita || 'Modelo principal',
      descricao: receita.descricao || ''
    },
    entradas: Array.isArray(receita.entradas) ? receita.entradas : [],
    saidas: Array.isArray(receita.saidas) ? receita.saidas : [],
    etapas: Array.isArray(etapas) ? etapas : []
  };
}

function salvarProdutoComModelo(produtoId, payloadProduto, dadosReceita, entradas, saidas, etapas) {
  const dadosProduto = {
    nome_produto: String(payloadProduto?.nome_produto || '').trim(),
    unidade_produto: String(payloadProduto?.unidade_produto || 'UN').trim() || 'UN'
  };

  if (!dadosProduto.nome_produto) {
    throw new Error('Nome do produto obrigatorio');
  }

  let id = String(produtoId || '').trim();
  if (id) {
    atualizarProduto(id, dadosProduto);
  } else {
    const criado = criarProduto(dadosProduto);
    id = criado.produto_id;
  }

  const modeloDados = {
    nome_receita: String(dadosReceita?.nome_receita || 'Modelo principal').trim() || 'Modelo principal',
    descricao: String(dadosReceita?.descricao || '').trim(),
    parent_receita_id: ''
  };

  const receitasAtuais = listarReceitasProduto(id);
  const receitaPrincipal = Array.isArray(receitasAtuais) && receitasAtuais.length > 0
    ? receitasAtuais[0]
    : criarReceitaProduto(id, modeloDados);

  atualizarReceitaProduto(receitaPrincipal.receita_id, modeloDados);
  salvarEntradasReceita(receitaPrincipal.receita_id, entradas || []);
  salvarSaidasReceita(receitaPrincipal.receita_id, saidas || []);
  inativarReceitasSecundariasProduto(id, receitaPrincipal.receita_id);
  salvarEtapasProduto(id, etapas || []);

  const produtoAtual = obterProduto(id) || {};
  const resumoModelo = getResumoModeloProduto(id);

  return {
    ...produtoAtual,
    ...resumoModelo,
    produto_id: id,
    nome_produto: produtoAtual.nome_produto || dadosProduto.nome_produto,
    unidade_produto: produtoAtual.unidade_produto || dadosProduto.unidade_produto,
    criado_em: produtoAtual.criado_em
      ? Utilities.formatDate(new Date(produtoAtual.criado_em), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
      : ''
  };
}

function obterReceitaPadraoProduto(produtoId) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (!sheet) return null;

  const receitas = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.produto_id === produtoId);

  if (receitas.length === 0) return null;
  const semParent = receitas.find(i => !String(i.parent_receita_id || '').trim());
  return semParent || receitas[0] || null;
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

function agruparItensExplosaoPorEstoque(itens) {
  const resultado = {};
  let custoPrevisto = 0;

  (Array.isArray(itens) ? itens : []).forEach(i => {
    if (!i || !i.estoque_id) return;

    const estoqueId = String(i.estoque_id);
    const quantidade = parseNumeroBR(i.quantidade);
    if (!quantidade || quantidade <= 0) return;

    const valorUnit = parseNumeroBR(i.valor_unit);

    if (!resultado[estoqueId]) {
      resultado[estoqueId] = {
        estoque_id: estoqueId,
        item: i.item || estoqueId,
        unidade: i.unidade || '',
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

function explodirReceitaDetalhada(produtoId, receitaId, qtdPlanejada) {
  if (!produtoId || !receitaId) return { itens: [], custoPrevisto: 0 };

  const sheetReceitas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (!sheetReceitas) return { itens: [], custoPrevisto: 0 };

  const receitas = rowsToObjects(sheetReceitas)
    .filter(i => String(i.ativo).toLowerCase() === 'true');
  const receita = receitas.find(r => r.receita_id === receitaId && r.produto_id === produtoId);
  if (!receita) {
    throw new Error('Receita nao encontrada');
  }

  const sheetEntradas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS_ENTRADAS);
  const entradas = sheetEntradas ? rowsToObjects(sheetEntradas) : [];

  const entradasAtivas = entradas
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.receita_id === receitaId);

  const produtosSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS);
  const produtosRows = produtosSheet ? rowsToObjects(produtosSheet) : [];
  const produtosMap = {};
  produtosRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(p => {
      produtosMap[p.produto_id] = p;
    });

  const qtdPlanejadaNum = parseNumeroBR(qtdPlanejada);
  if (!qtdPlanejadaNum || qtdPlanejadaNum <= 0) {
    return { itens: [], custoPrevisto: 0 };
  }

  // A OP representa quantidade de modelos, entao cada linha da receita multiplica por qtdPlanejada.
  const fator = qtdPlanejadaNum;

  const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
  const estoqueRows = sheetEstoque ? rowsToObjects(sheetEstoque) : [];
  const estoqueMap = {};
  estoqueRows.forEach(i => {
    estoqueMap[i.ID] = i;
  });

  const itensDetalhados = [];
  let custoPrevisto = 0;

  entradasAtivas.forEach(e => {
    const tipo = String(e.tipo_item || '').toUpperCase();
    const prodId = e.produto_ref_id || '';

    let estoqueId = e.estoque_ref_id || '';
    let itemNome = e.nome_item || '';
    let unidade = e.unidade || '';
    let valorUnit = parseNumeroBR(e.custo_manual);
    let quantidade = 0;

    const qtdBase = parseNumeroBR(e.qtd_pecas);
    if (!qtdBase || qtdBase <= 0) return;

    if (tipo === 'PRODUTO' && prodId) {
      const prod = produtosMap[prodId] || null;
      if (prod) {
        itemNome = prod.nome_produto || itemNome || prodId;
        unidade = prod.unidade_produto || unidade || 'UN';
      }
    }

    if (tipo === 'MADEIRA') {
      const comp = parseNumeroBR(e.comprimento_cm);
      const larg = parseNumeroBR(e.largura_cm);
      const esp = parseNumeroBR(e.espessura_cm);
      const volumeM3 = (comp > 0 && larg > 0 && esp > 0)
        ? ((comp * larg * esp) / 1000000)
        : 1;
      quantidade = qtdBase * volumeM3 * fator;
      unidade = unidade || 'M3';
    } else {
      quantidade = qtdBase * fator;
    }

    const estoqueItem = estoqueId ? (estoqueMap[estoqueId] || null) : null;
    if (estoqueItem) {
      itemNome = estoqueItem.item || itemNome || estoqueId;
      unidade = estoqueItem.unidade || unidade || '';
      const valorEstoque = parseNumeroBR(estoqueItem.valor_unit);
      if (valorEstoque > 0 || !valorUnit) {
        valorUnit = valorEstoque;
      }
    }

    if (!quantidade || quantidade <= 0) return;
    if (!itemNome) {
      itemNome = prodId || e.nome_item || `Item ${tipo || 'LIVRE'}`;
    }

    const quantidadeNorm = parseNumeroBR(quantidade);
    const valorUnitNorm = parseNumeroBR(valorUnit);
    const total = quantidadeNorm * valorUnitNorm;

    itensDetalhados.push({
      receita_entrada_id: e.id || '',
      receita_id: receitaId,
      estoque_id: estoqueId,
      tipo_item: tipo,
      origem_item: e.nome_item || itemNome,
      item: itemNome,
      unidade,
      quantidade: quantidadeNorm,
      valor_unit: valorUnitNorm,
      total_previsto: total
    });

    custoPrevisto += total;
  });

  return {
    itens: itensDetalhados,
    custoPrevisto
  };
}

function explodirSaidasReceitaDetalhada(produtoId, receitaId, qtdPlanejada) {
  if (!produtoId || !receitaId) return { itens: [] };

  const sheetReceitas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (!sheetReceitas) return { itens: [] };

  const receitas = rowsToObjects(sheetReceitas)
    .filter(i => String(i.ativo).toLowerCase() === 'true');
  const receita = receitas.find(r => r.receita_id === receitaId && r.produto_id === produtoId);
  if (!receita) {
    throw new Error('Receita nao encontrada');
  }

  const sheetSaidas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS_SAIDAS);
  const saidas = sheetSaidas ? rowsToObjects(sheetSaidas) : [];
  const saidasAtivas = saidas
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.receita_id === receitaId);

  const produtosSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS);
  const produto = produtosSheet
    ? rowsToObjects(produtosSheet).find(i => i.produto_id === produtoId && String(i.ativo).toLowerCase() === 'true')
    : null;
  const nomeProduto = produto ? (produto.nome_produto || '') : '';

  const qtdPlanejadaNum = parseNumeroBR(qtdPlanejada);
  if (!qtdPlanejadaNum || qtdPlanejadaNum <= 0) {
    return { itens: [] };
  }

  // A OP representa quantidade de modelos, entao cada saida multiplica por qtdPlanejada.
  const fator = qtdPlanejadaNum;

  const itens = [];
  saidasAtivas.forEach(s => {
    const qtdBase = parseNumeroBR(s.quantidade);
    if (!qtdBase || qtdBase <= 0) return;

    const nomeSaida = String(s.nome_saida || '').trim() || nomeProduto || 'Saida';
    const quantidade = parseNumeroBR(qtdBase * fator);
    if (!quantidade || quantidade <= 0) return;

    itens.push({
      receita_saida_id: s.id || '',
      receita_id: receitaId,
      nome_saida: nomeSaida,
      unidade: s.unidade || '',
      quantidade
    });
  });

  return { itens };
}

function explodirReceita(produtoId, receitaId, qtdPlanejada) {
  const detalhada = explodirReceitaDetalhada(produtoId, receitaId, qtdPlanejada);
  const agrupada = agruparItensExplosaoPorEstoque(detalhada.itens || []);

  return {
    itens: agrupada.itens || [],
    custoPrevisto: parseNumeroBR(agrupada.custoPrevisto)
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
  const sheetComponentes = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_COMPONENTES);
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

  const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
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

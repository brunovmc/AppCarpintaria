const ABA_PRODUCAO = 'PRODUCAO';
const ABA_PRODUCAO_ETAPAS = 'PRODUCAO_ETAPAS';
const ABA_PRODUCAO_CONSUMO = 'PRODUCAO_CONSUMO';
const ABA_PRODUCAO_MATERIAIS = 'PRODUCAO_MATERIAIS';
const ABA_PRODUCAO_MATERIAIS_PREVISTOS = 'PRODUCAO_MATERIAIS_PREVISTOS';

const PRODUCAO_SCHEMA = [
  'producao_id',
  'produto_id',
  'receita_id',
  'nome_ordem',
  'qtd_planejada',
  'status',
  'data_inicio',
  'data_prevista_termino',
  'data_conclusao',
  'observacao',
  'estoque_atualizado',
  'data_estoque_atualizado',
  'ativo',
  'criado_em'
];

const PRODUCAO_ETAPAS_SCHEMA = [
  'id',
  'producao_id',
  'nome_etapa',
  'feito',
  'ordem',
  'data_feito',
  'ativo'
];

const PRODUCAO_CONSUMO_SCHEMA = [
  'id',
  'producao_id',
  'estoque_id',
  'quantidade_consumida',
  'valor_unit_snapshot',
  'total_snapshot',
  'criado_em',
  'ativo'
];

const PRODUCAO_MATERIAIS_SCHEMA = [
  'id',
  'producao_id',
  'estoque_id',
  'quantidade',
  'unidade',
  'criado_em',
  'ativo'
];

const PRODUCAO_MATERIAIS_PREVISTOS_SCHEMA = [
  'id',
  'producao_id',
  'estoque_id',
  'quantidade',
  'unidade',
  'item_snapshot',
  'valor_unit_snapshot',
  'criado_em',
  'ativo'
];

function formatDateSafe(value, pattern) {
  if (!value) return '';
  const dt = value instanceof Date ? value : new Date(value);
  if (!(dt instanceof Date) || isNaN(dt)) return '';
  const fmt = pattern || 'yyyy-MM-dd';
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), fmt);
}

function listarProducao() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUCAO);
  if (!sheet) return [];

  const produtosSheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS);
  const produtos = produtosSheet ? rowsToObjects(produtosSheet) : [];
  const produtosMap = {};
  produtos
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(p => {
      produtosMap[p.produto_id] = p;
    });

  const receitasSheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_RECEITAS);
  const receitasRows = receitasSheet ? rowsToObjects(receitasSheet) : [];
  const receitasMap = {};
  receitasRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(r => {
      receitasMap[r.receita_id] = r;
    });

  const etapasSheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUCAO_ETAPAS);
  const etapasRows = etapasSheet ? rowsToObjects(etapasSheet) : [];
  const etapasMap = {};
  etapasRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(e => {
      const etapa = {
        ...e,
        feito: String(e.feito).toLowerCase() === 'true',
        data_feito: formatDateSafe(e.data_feito, 'yyyy-MM-dd')
      };
      if (!etapasMap[etapa.producao_id]) etapasMap[etapa.producao_id] = [];
      etapasMap[etapa.producao_id].push(etapa);
    });

  Object.keys(etapasMap).forEach(id => {
    etapasMap[id].sort((a, b) => parseNumeroBR(a.ordem) - parseNumeroBR(b.ordem));
  });

  const rows = rowsToObjects(sheet);
  return rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .map(i => {
      const prod = produtosMap[i.produto_id] || {};
      const receita = receitasMap[i.receita_id] || {};
      return {
        ...i,
        nome_produto: prod.nome_produto || '',
        unidade_produto: prod.unidade_produto || '',
        receita_nome: receita.nome_receita || '',
        qtd_planejada: parseNumeroBR(i.qtd_planejada),
        estoque_atualizado: String(i.estoque_atualizado).toLowerCase() === 'true',
        criado_em: i.criado_em
          ? formatDateSafe(i.criado_em, 'yyyy-MM-dd HH:mm')
          : '',
        data_inicio: i.data_inicio
          ? formatDateSafe(i.data_inicio, 'yyyy-MM-dd')
          : '',
        data_prevista_termino: i.data_prevista_termino
          ? formatDateSafe(i.data_prevista_termino, 'yyyy-MM-dd')
          : '',
        data_conclusao: i.data_conclusao
          ? formatDateSafe(i.data_conclusao, 'yyyy-MM-dd')
          : '',
        data_estoque_atualizado: i.data_estoque_atualizado
          ? formatDateSafe(i.data_estoque_atualizado, 'yyyy-MM-dd')
          : '',
        etapas: etapasMap[i.producao_id] || []
      };
    });
}

function criarProducao(payload) {
  const novo = {
    ...payload,
    producao_id: gerarId('OP'),
    status: payload.status || 'Em planejamento',
    qtd_planejada: parseNumeroBR(payload.qtd_planejada),
    estoque_atualizado: false,
    data_estoque_atualizado: '',
    ativo: true,
    criado_em: new Date()
  };

  const base = gerarMateriaisPrevistosReceita(
    novo.produto_id,
    novo.receita_id,
    novo.qtd_planejada
  );

  insert(ABA_PRODUCAO, novo, PRODUCAO_SCHEMA);
  salvarMateriaisPrevistosSnapshot(novo.producao_id, base.itens || []);

  let nomeProduto = '';
  let unidadeProduto = '';
  const produtosSheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS);
  if (produtosSheet) {
    const produtos = rowsToObjects(produtosSheet);
    const prod = produtos.find(p => p.produto_id === novo.produto_id);
    if (prod) {
      nomeProduto = prod.nome_produto || '';
      unidadeProduto = prod.unidade_produto || '';
    }
  }

  let receitaNome = '';
  const receitasSheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (receitasSheet) {
    const receitas = rowsToObjects(receitasSheet);
    const rec = receitas.find(r => r.receita_id === novo.receita_id);
    if (rec) {
      receitaNome = rec.nome_receita || '';
    }
  }

  const etapasTemplate = listarEtapasProduto(novo.produto_id) || [];
  const etapasCriadas = etapasTemplate.map(etapa => {
    const nova = {
      id: gerarId('EPR'),
      producao_id: novo.producao_id,
      nome_etapa: etapa.nome_etapa || '',
      feito: false,
      ordem: etapa.ordem || '',
      data_feito: '',
      ativo: true
    };
    insert(ABA_PRODUCAO_ETAPAS, nova, PRODUCAO_ETAPAS_SCHEMA);
    return nova;
  });

  return {
    ...novo,
    nome_produto: nomeProduto,
    unidade_produto: unidadeProduto,
    receita_nome: receitaNome,
    custo_previsto: base.custoPrevisto || 0,
    criado_em: Utilities.formatDate(novo.criado_em, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    data_inicio: novo.data_inicio || '',
    data_prevista_termino: novo.data_prevista_termino || '',
    data_conclusao: novo.data_conclusao || '',
    etapas: etapasCriadas
  };
}

function atualizarProducao(id, payload) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUCAO);
  if (!sheet) return false;

  const rows = rowsToObjects(sheet);
  const atual = rows.find(i => i.producao_id === id);
  if (!atual) return false;

  const dadosAtualizados = {
    ...atual,
    ...payload
  };

  const mudouReceita = Object.prototype.hasOwnProperty.call(payload, 'receita_id') &&
    payload.receita_id !== atual.receita_id;
  const mudouProduto = Object.prototype.hasOwnProperty.call(payload, 'produto_id') &&
    payload.produto_id !== atual.produto_id;
  const mudouQtd = Object.prototype.hasOwnProperty.call(payload, 'qtd_planejada') &&
    parseNumeroBR(payload.qtd_planejada) !== parseNumeroBR(atual.qtd_planejada);

  if (mudouReceita || mudouProduto || mudouQtd) {
    const base = gerarMateriaisPrevistosReceita(
      dadosAtualizados.produto_id,
      dadosAtualizados.receita_id,
      parseNumeroBR(dadosAtualizados.qtd_planejada)
    );
    salvarMateriaisPrevistosSnapshot(id, base.itens || []);
  }

  return updateById(
    ABA_PRODUCAO,
    'producao_id',
    id,
    payload,
    PRODUCAO_SCHEMA
  );
}

function deletarProducao(id) {
  return updateById(
    ABA_PRODUCAO,
    'producao_id',
    id,
    { ativo: false },
    PRODUCAO_SCHEMA
  );
}

function listarEtapasProducao(producaoId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUCAO_ETAPAS);
  if (!sheet) return [];

  const rows = rowsToObjects(sheet);
  return rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.producao_id === producaoId)
    .map(i => ({
      ...i,
      feito: String(i.feito).toLowerCase() === 'true',
      data_feito: formatDateSafe(i.data_feito, 'yyyy-MM-dd')
    }))
    .sort((a, b) => parseNumeroBR(a.ordem) - parseNumeroBR(b.ordem));
}

function listarMateriaisExtrasProducao(producaoId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUCAO_MATERIAIS);
  if (!sheet) return [];

  const rows = rowsToObjects(sheet);
  return rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.producao_id === producaoId)
    .map(i => ({
      estoque_id: i.estoque_id,
      quantidade: parseNumeroBR(i.quantidade),
      unidade: i.unidade || ''
    }));
}

function listarMateriaisPrevistosSnapshot(producaoId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUCAO_MATERIAIS_PREVISTOS);
  if (!sheet) return [];

  const rows = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.producao_id === producaoId);

  if (!rows || rows.length === 0) return [];

  const sheetEstoque = SpreadsheetApp.getActive().getSheetByName(ABA_ESTOQUE);
  const estoqueRows = sheetEstoque ? rowsToObjects(sheetEstoque) : [];
  const estoqueMap = {};
  estoqueRows.forEach(i => {
    estoqueMap[i.ID] = i;
  });

  return rows.map(r => {
    const estoqueItem = estoqueMap[r.estoque_id] || {};
    return {
      estoque_id: r.estoque_id,
      item: r.item_snapshot || estoqueItem.item || r.estoque_id,
      unidade: r.unidade || estoqueItem.unidade || '',
      quantidade: parseNumeroBR(r.quantidade),
      valor_unit: parseNumeroBR(r.valor_unit_snapshot || estoqueItem.valor_unit)
    };
  });
}

function limparMateriaisPrevistosSnapshot(producaoId) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(ABA_PRODUCAO_MATERIAIS_PREVISTOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUCAO_MATERIAIS_PREVISTOS);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idCol = headers.indexOf('producao_id');

  if (idCol !== -1 && data.length > 1) {
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][idCol] === producaoId) {
        sheet.deleteRow(i + 1);
      }
    }
  }
}

function salvarMateriaisPrevistosSnapshot(producaoId, itens) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(ABA_PRODUCAO_MATERIAIS_PREVISTOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUCAO_MATERIAIS_PREVISTOS);
  }

  ensureSchema(sheet, PRODUCAO_MATERIAIS_PREVISTOS_SCHEMA);
  limparMateriaisPrevistosSnapshot(producaoId);

  const linhasValidas = Array.isArray(itens) ? itens : [];
  linhasValidas.forEach(l => {
    if (!l.estoque_id) return;
    const quantidade = parseNumeroBR(l.quantidade);
    if (!quantidade || quantidade <= 0) return;

    const novo = {
      id: gerarId('PMP'),
      producao_id: producaoId,
      estoque_id: l.estoque_id,
      quantidade,
      unidade: l.unidade || '',
      item_snapshot: l.item || '',
      valor_unit_snapshot: parseNumeroBR(l.valor_unit),
      criado_em: new Date(),
      ativo: true
    };
    insert(ABA_PRODUCAO_MATERIAIS_PREVISTOS, novo, PRODUCAO_MATERIAIS_PREVISTOS_SCHEMA);
  });

  return true;
}

function gerarMateriaisPrevistosReceita(produtoId, receitaId, qtdPlanejada) {
  if (!produtoId || !receitaId) {
    throw new Error('Receita nao informada');
  }

  const resp = explodirReceita(produtoId, receitaId, qtdPlanejada);
  return resp || { itens: [], custoPrevisto: 0 };
}

function adicionarMaterialExtraProducao(producaoId, estoqueId, quantidade) {
  if (!producaoId || !estoqueId) {
    throw new Error('Producao ou estoque invalido');
  }

  const qtd = parseNumeroBR(quantidade);
  if (!qtd || qtd <= 0) {
    throw new Error('Quantidade invalida');
  }

  const sheetEstoque = SpreadsheetApp.getActive().getSheetByName(ABA_ESTOQUE);
  if (!sheetEstoque) {
    throw new Error('Aba ESTOQUE nao encontrada');
  }

  const estoqueRows = rowsToObjects(sheetEstoque);
  const estoqueItem = estoqueRows.find(i => i.ID === estoqueId) || {};
  const unidade = estoqueItem.unidade || '';

  const novo = {
    id: gerarId('PM'),
    producao_id: producaoId,
    estoque_id: estoqueId,
    quantidade: qtd,
    unidade,
    criado_em: new Date(),
    ativo: true
  };

  insert(ABA_PRODUCAO_MATERIAIS, novo, PRODUCAO_MATERIAIS_SCHEMA);

  return {
    ...novo,
    criado_em: formatDateSafe(novo.criado_em, 'yyyy-MM-dd HH:mm')
  };
}

function aplicarBaixaMateriaisExtras(producaoId, itensValidos) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUCAO_MATERIAIS);
  if (!sheet) return;

  const rows = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.producao_id === producaoId);

  const extrasPorEstoque = {};
  rows.forEach(r => {
    if (!extrasPorEstoque[r.estoque_id]) {
      extrasPorEstoque[r.estoque_id] = [];
    }
    extrasPorEstoque[r.estoque_id].push(r);
  });

  (Array.isArray(itensValidos) ? itensValidos : []).forEach(item => {
    let restante = parseNumeroBR(item.quantidade);
    if (restante <= 0) return;

    const lista = extrasPorEstoque[item.estoque_id] || [];
    for (let i = 0; i < lista.length && restante > 0; i++) {
      const extra = lista[i];
      const qtdAtual = parseNumeroBR(extra.quantidade);
      if (!extra.id || qtdAtual <= 0) continue;

      if (qtdAtual <= restante) {
        restante -= qtdAtual;
        updateById(
          ABA_PRODUCAO_MATERIAIS,
          'id',
          extra.id,
          { quantidade: 0, ativo: false },
          PRODUCAO_MATERIAIS_SCHEMA
        );
        continue;
      }

      const novoQtd = qtdAtual - restante;
      restante = 0;
      updateById(
        ABA_PRODUCAO_MATERIAIS,
        'id',
        extra.id,
        { quantidade: novoQtd },
        PRODUCAO_MATERIAIS_SCHEMA
      );
    }
  });
}

function atualizarEtapasProducao(producaoId, etapas) {
  if (!Array.isArray(etapas)) return false;

  etapas.forEach(et => {
    const dataFeito = et.data_feito
      ? new Date(et.data_feito)
      : '';
    const dataFeitoValida = dataFeito instanceof Date && !isNaN(dataFeito)
      ? dataFeito
      : '';

    if (et.id) {
      updateById(
        ABA_PRODUCAO_ETAPAS,
        'id',
        et.id,
        {
          nome_etapa: et.nome_etapa,
          feito: et.feito,
          ordem: et.ordem,
          data_feito: dataFeitoValida
        },
        PRODUCAO_ETAPAS_SCHEMA
      );
      return;
    }

    const novo = {
      id: gerarId('EPR'),
      producao_id: producaoId,
      nome_etapa: et.nome_etapa || '',
      feito: !!et.feito,
      ordem: et.ordem || '',
      data_feito: dataFeitoValida,
      ativo: true
    };
    insert(ABA_PRODUCAO_ETAPAS, novo, PRODUCAO_ETAPAS_SCHEMA);
  });

  return true;
}

function deletarEtapaProducao(id) {
  return updateById(
    ABA_PRODUCAO_ETAPAS,
    'id',
    id,
    { ativo: false },
    PRODUCAO_ETAPAS_SCHEMA
  );
}

function obterMateriaisPrevistosProducao(producaoId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUCAO);
  if (!sheet) return { itens: [], custoPrevisto: 0 };

  const rows = rowsToObjects(sheet);
  const ordem = rows.find(i => i.producao_id === producaoId);
  if (!ordem) return { itens: [], custoPrevisto: 0 };
  if (!ordem.receita_id) return { itens: [], custoPrevisto: 0 };

  let baseItens = listarMateriaisPrevistosSnapshot(producaoId);
  let custoPrevisto = 0;

  if (!baseItens || baseItens.length === 0) {
    const base = gerarMateriaisPrevistosReceita(
      ordem.produto_id,
      ordem.receita_id,
      ordem.qtd_planejada
    );
    baseItens = base.itens || [];
    custoPrevisto = parseNumeroBR(base.custoPrevisto);
    salvarMateriaisPrevistosSnapshot(producaoId, baseItens);
  } else {
    custoPrevisto = baseItens.reduce((acc, i) => {
      return acc + (parseNumeroBR(i.quantidade) * parseNumeroBR(i.valor_unit));
    }, 0);
  }

  const extras = listarMateriaisExtrasProducao(producaoId);

  if (!extras || extras.length === 0) {
    return {
      itens: baseItens,
      custoPrevisto
    };
  }

  const sheetEstoque = SpreadsheetApp.getActive().getSheetByName(ABA_ESTOQUE);
  const estoqueRows = sheetEstoque ? rowsToObjects(sheetEstoque) : [];
  const estoqueMap = {};
  estoqueRows.forEach(i => {
    estoqueMap[i.ID] = i;
  });

  const itensMap = {};
  (baseItens || []).forEach(i => {
    itensMap[i.estoque_id] = { ...i };
  });

  extras.forEach(e => {
    const estoqueItem = estoqueMap[e.estoque_id] || {};
    const valorUnit = parseNumeroBR(estoqueItem.valor_unit);
    const unidade = estoqueItem.unidade || e.unidade || '';
    const nomeItem = estoqueItem.item || e.estoque_id || '';
    const qtd = parseNumeroBR(e.quantidade);

    if (!itensMap[e.estoque_id]) {
      itensMap[e.estoque_id] = {
        estoque_id: e.estoque_id,
        item: nomeItem,
        unidade,
        quantidade: 0,
        valor_unit: valorUnit
      };
    }

    itensMap[e.estoque_id].quantidade += qtd;
    custoPrevisto += qtd * valorUnit;
  });

  return {
    itens: Object.values(itensMap),
    custoPrevisto
  };
}

function consumirEstoque(producaoId, itensParaBaixar) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (!producaoId) {
      throw new Error('Producao invalida');
    }

    const sheetProducao = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUCAO);
    if (!sheetProducao) {
      throw new Error('Aba PRODUCAO nao encontrada');
    }

    const producaoRows = rowsToObjects(sheetProducao);
    const ordem = producaoRows.find(i => i.producao_id === producaoId);
    if (!ordem) {
      throw new Error('Producao nao encontrada');
    }

    if (!ordem.receita_id) {
      throw new Error('Receita nao informada');
    }

    const status = String(ordem.status || '');
    if (status !== 'Concluido' && status !== 'Concluida') {
      throw new Error('Atualizacao de estoque liberada apenas quando status for Concluido');
    }

    if (String(ordem.estoque_atualizado).toLowerCase() === 'true') {
      throw new Error('Estoque ja atualizado para esta producao');
    }

    const itens = Array.isArray(itensParaBaixar) ? itensParaBaixar : [];
    const itensMap = {};
    itens.forEach(i => {
      const estoqueId = i && i.estoque_id ? String(i.estoque_id) : '';
      const qtd = parseNumeroBR(i ? i.quantidade : 0);
      if (!estoqueId || qtd <= 0) return;
      itensMap[estoqueId] = (itensMap[estoqueId] || 0) + qtd;
    });

    const itensValidos = Object.keys(itensMap).map(estoqueId => ({
      estoque_id: estoqueId,
      quantidade: itensMap[estoqueId]
    }));

    if (itensValidos.length === 0) {
      throw new Error('Nenhum item para baixar');
    }

    const sheetEstoque = SpreadsheetApp.getActive().getSheetByName(ABA_ESTOQUE);
    if (!sheetEstoque) {
      throw new Error('Aba ESTOQUE nao encontrada');
    }

    const estoqueRows = rowsToObjects(sheetEstoque);
    const estoqueMap = {};
    estoqueRows.forEach(i => {
      estoqueMap[i.ID] = i;
    });

    const produtosSheet = SpreadsheetApp.getActive().getSheetByName(ABA_PRODUTOS);
    const produtos = produtosSheet ? rowsToObjects(produtosSheet) : [];
    const produto = produtos.find(p => p.produto_id === ordem.produto_id) || null;
    if (!produto) {
      throw new Error('Produto nao encontrado');
    }

    const estoqueProdutoId = produto.estoque_id || '';
    if (!estoqueProdutoId) {
      throw new Error('Produto sem estoque vinculado');
    }

    const estoqueProduto = estoqueMap[estoqueProdutoId];
    if (!estoqueProduto) {
      throw new Error('Item de estoque do produto nao encontrado');
    }

    const ativoProduto = String(estoqueProduto.ativo).toLowerCase() === 'true';
    if (!ativoProduto) {
      throw new Error('Item de estoque do produto inativo');
    }
    if (String(estoqueProduto.tipo || '').toUpperCase() !== 'PRODUTO') {
      throw new Error('Item de estoque do produto deve ser do tipo PRODUTO');
    }

    const qtdProduzida = parseNumeroBR(ordem.qtd_planejada);
    if (!qtdProduzida || qtdProduzida <= 0) {
      throw new Error('Quantidade planejada invalida');
    }

    // Obriga informar consumo para todos os materiais previstos da ordem.
    const previstosResp = obterMateriaisPrevistosProducao(producaoId);
    const previstos = Array.isArray(previstosResp && previstosResp.itens)
      ? previstosResp.itens
      : [];
    const previstosMap = {};
    previstos.forEach(i => {
      const estoqueId = i && i.estoque_id ? String(i.estoque_id) : '';
      const qtd = parseNumeroBR(i ? i.quantidade : 0);
      if (!estoqueId || qtd <= 0) return;
      previstosMap[estoqueId] = (previstosMap[estoqueId] || 0) + qtd;
    });

    const faltantes = Object.keys(previstosMap).filter(estoqueId => !itensMap[estoqueId]);
    if (faltantes.length > 0) {
      const nomes = faltantes.map(id => {
        const it = estoqueMap[id];
        return it ? (it.item || id) : id;
      });
      throw new Error(`Consumo incompleto. Informe todos os materiais previstos: ${nomes.join(', ')}`);
    }

    const naoPrevistos = Object.keys(itensMap).filter(estoqueId => !previstosMap[estoqueId]);
    if (naoPrevistos.length > 0) {
      const nomes = naoPrevistos.map(id => {
        const it = estoqueMap[id];
        return it ? (it.item || id) : id;
      });
      throw new Error(`Itens nao previstos informados para baixa: ${nomes.join(', ')}`);
    }

    itensValidos.forEach(i => {
      const estoqueItem = estoqueMap[i.estoque_id];
      if (!estoqueItem) {
        throw new Error(`Item de estoque nao encontrado: ${i.estoque_id}`);
      }
      const ativo = String(estoqueItem.ativo).toLowerCase() === 'true';
      if (!ativo) {
        throw new Error(`Item de estoque inativo: ${i.estoque_id}`);
      }
      const saldo = parseNumeroBR(estoqueItem.quantidade);
      if (saldo < i.quantidade) {
        throw new Error(`Saldo insuficiente para ${estoqueItem.item || i.estoque_id}`);
      }
    });

    const consumoRegistrado = [];
    const estoqueAtualizados = [];

    itensValidos.forEach(i => {
      const estoqueItem = estoqueMap[i.estoque_id];
      const saldo = parseNumeroBR(estoqueItem.quantidade);
      const novoSaldo = saldo - i.quantidade;

      updateById(
        ABA_ESTOQUE,
        'ID',
        i.estoque_id,
        { quantidade: novoSaldo },
        ESTOQUE_SCHEMA
      );

      estoqueMap[i.estoque_id].quantidade = novoSaldo;

      estoqueAtualizados.push({
        ID: i.estoque_id,
        quantidade: novoSaldo
      });

      const valorUnit = parseNumeroBR(estoqueItem.valor_unit);
      const total = i.quantidade * valorUnit;

      const consumo = {
        id: gerarId('PC'),
        producao_id: producaoId,
        estoque_id: i.estoque_id,
        quantidade_consumida: i.quantidade,
        valor_unit_snapshot: valorUnit,
        total_snapshot: total,
        criado_em: new Date(),
        ativo: true
      };

      insert(ABA_PRODUCAO_CONSUMO, consumo, PRODUCAO_CONSUMO_SCHEMA);
      consumoRegistrado.push(consumo);
    });

    aplicarBaixaMateriaisExtras(producaoId, itensValidos);

    const saldoProduto = parseNumeroBR(estoqueProduto.quantidade);
    const novoSaldoProduto = saldoProduto + qtdProduzida;

    updateById(
      ABA_ESTOQUE,
      'ID',
      estoqueProdutoId,
      { quantidade: novoSaldoProduto },
      ESTOQUE_SCHEMA
    );

    estoqueAtualizados.push({
      ID: estoqueProdutoId,
      quantidade: novoSaldoProduto
    });

    const dataAtualizacao = new Date();
    updateById(
      ABA_PRODUCAO,
      'producao_id',
      producaoId,
      {
        estoque_atualizado: true,
        data_estoque_atualizado: dataAtualizacao
      },
      PRODUCAO_SCHEMA
    );

    return {
      estoqueAtualizados,
      consumoRegistrado,
      producaoAtualizada: {
        producao_id: producaoId,
        estoque_atualizado: true,
        data_estoque_atualizado: formatDateSafe(dataAtualizacao, 'yyyy-MM-dd')
      }
    };
  } finally {
    lock.releaseLock();
  }
}

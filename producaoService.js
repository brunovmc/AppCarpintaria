const ABA_PRODUCAO = 'PRODUCAO';
const ABA_PRODUCAO_ETAPAS = 'PRODUCAO_ETAPAS';
const ABA_PRODUCAO_CONSUMO = 'PRODUCAO_CONSUMO';

const PRODUCAO_SCHEMA = [
  'producao_id',
  'produto_id',
  'nome_ordem',
  'qtd_planejada',
  'status',
  'data_inicio',
  'data_prevista_termino',
  'data_conclusao',
  'observacao',
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
      return {
        ...i,
        nome_produto: prod.nome_produto || '',
        unidade_produto: prod.unidade_produto || '',
        qtd_planejada: parseNumeroBR(i.qtd_planejada),
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
    ativo: true,
    criado_em: new Date()
  };

  insert(ABA_PRODUCAO, novo, PRODUCAO_SCHEMA);

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
    criado_em: Utilities.formatDate(novo.criado_em, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    data_inicio: novo.data_inicio || '',
    data_prevista_termino: novo.data_prevista_termino || '',
    data_conclusao: novo.data_conclusao || '',
    etapas: etapasCriadas
  };
}

function atualizarProducao(id, payload) {
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

  return explodirBOM(ordem.produto_id, ordem.qtd_planejada);
}

function consumirEstoque(producaoId, itensParaBaixar) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (!producaoId) {
      throw new Error('Producao invalida');
    }

    const itens = Array.isArray(itensParaBaixar) ? itensParaBaixar : [];
    const itensValidos = itens
      .map(i => ({
        estoque_id: i.estoque_id,
        quantidade: parseNumeroBR(i.quantidade)
      }))
      .filter(i => i.estoque_id && i.quantidade > 0);

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

    return {
      estoqueAtualizados,
      consumoRegistrado
    };
  } finally {
    lock.releaseLock();
  }
}

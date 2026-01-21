const ABA_COMPRAS = 'COMPRAS';

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

function listarCompras() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ABA_COMPRAS);
  if (!sheet) return [];

  const rows = rowsToObjects(sheet);

  return rows
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
}

function criarItemCompra(payload) {
  const novo = {
    ...payload,
    ID: gerarId('COM'),
    ativo: true,
    criado_em: new Date(),
    adicionado_estoque: false
  };

  return insert(ABA_COMPRAS, novo, COMPRAS_SCHEMA);
}

function atualizarItemCompra(id, payload) {
  return updateById(
    ABA_COMPRAS,
    'ID',
    id,
    payload,
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

    const itemEstoqueCriado = {
      ID: gerarId('EST'),
      tipo: compra.tipo || '',
      item: compra.item || '',
      unidade: compra.unidade || '',
      valor_unit: compra.valor_unit || 0,
      ativo: true,
      criado_em: new Date(),
      quantidade: compra.quantidade || 0,
      comprimento_cm: compra.comprimento_cm || '',
      largura_cm: compra.largura_cm || '',
      espessura_cm: compra.espessura_cm || '',
      categoria: compra.categoria || '',
      fornecedor: compra.fornecedor || '',
      potencia: compra.potencia || '',
      voltagem: compra.voltagem || '',
      comprado_em: compra.comprado_em || '',
      vida_util_mes: compra.vida_util_mes || '',
      observacao: compra.observacao || ''
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

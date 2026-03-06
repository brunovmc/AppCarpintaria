const ABA_ESTOQUE = 'ESTOQUE';

const ESTOQUE_SCHEMA = [
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
  'observacao'
];

function normalizarPayloadMadeiraEstoque(payload) {
  const dados = { ...(payload || {}) };
  const tipo = String(dados.tipo || '').trim().toUpperCase();
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


function listarEstoque() {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
  if (!sheet) return [];

  const rows = rowsToObjects(sheet);

  return rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
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

function testeListarEstoqueDireto() {
  const itens = listarEstoque();
  Logger.log(itens);
}

function testeDebugEstoque() {
  const sheet = getSheet('ESTOQUE');
  const todos = rowsToObjects(sheet);
  Logger.log('TODOS:');
  Logger.log(todos);

  const filtrados = listarEstoque();
  Logger.log('FILTRADOS:');
  Logger.log(filtrados);
}



function criarItemEstoque(payload) {
  const dados = normalizarPayloadMadeiraEstoque(payload);
  const novo = {
    ...dados,
    ID: gerarId('EST'),
    ativo: true,
    criado_em: new Date()
  };

  return insert(ABA_ESTOQUE, novo, ESTOQUE_SCHEMA);
}


function testeCriarItem(){
  criarItemEstoque({
  tipo: 'MADEIRA',
  item: 'TESTE',
  unidade: 'M3',
  valor_unit: 9999
});
}

function obterItemEstoque(id) {
  const sheet = getSheet(ABA_ESTOQUE);
  if (!sheet) return null;

  const item = rowsToObjects(sheet).find(i => i.ID === id);
  if (!item) return null;

  return {
    ID: item.ID,
    tipo: item.tipo || '',
    item: item.item || '',
    categoria: item.categoria || '',
    unidade: item.unidade || '',
    quantidade: item.quantidade || '',
    comprimento_cm: item.comprimento_cm || '',
    largura_cm: item.largura_cm || '',
    espessura_cm: item.espessura_cm || '',
    valor_unit: item.valor_unit || '',
    fornecedor: item.fornecedor || '',
    observacao: item.observacao || '',

    potencia: item.potencia || '',
    voltagem: item.voltagem || '',
    vida_util_mes: item.vida_util_mes || '',

    criado_em: item.criado_em
      ? Utilities.formatDate(
          new Date(item.criado_em),
          Session.getScriptTimeZone(),
          'yyyy-MM-dd HH:mm'
        )
      : '',

    comprado_em: item.comprado_em
      ? Utilities.formatDate(
          new Date(item.comprado_em),
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        )
      : ''
  };
}



function atualizarItemEstoque(id, payload) {
  const dados = normalizarPayloadMadeiraEstoque(payload);
  return updateById(
    ABA_ESTOQUE,
    'ID',
    id,
    dados,
    ESTOQUE_SCHEMA
  );
}



function deletarItemEstoque(id) {
  return updateById(
    ABA_ESTOQUE,
    'ID',
    id,
    { ativo: false },
    ESTOQUE_SCHEMA
  );
}

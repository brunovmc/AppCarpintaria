const DATA_SPREADSHEET_ID = "1j363esbdvygUO1s72bCz5u3Vhn2AnhiIHlmO4ZApowg";

function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("CarpintariaZizu")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function include(nome) {
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}

function listarEstoqueAPI() {
  return listarEstoque();
}

function criarItemEstoqueAPI(payload) {
  return criarItemEstoque(payload);
}

function atualizarItemEstoqueAPI(id, payload) {
  return atualizarItemEstoque(id, payload);
}

function deletarItemEstoqueAPI(id) {
  return deletarItemEstoque(id);
}

function sincronizarSchemaFinanceiro() {
  const schemaCompras = [
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
    'data_vencimento',
    'forma_pagamento_padrao',
    'vida_util_mes',
    'observacao',
    'adicionado_estoque',
    'estoque_id',
    'origem_compra_id'
  ];

  const schemaDespesasGerais = [
    'ID',
    'descricao',
    'categoria',
    'fornecedor',
    'valor_total',
    'ativo',
    'criado_em',
    'data_competencia',
    'data_vencimento',
    'forma_pagamento_padrao',
    'observacao'
  ];

  const schemaPagamentos = [
    'ID',
    'origem_tipo',
    'origem_id',
    'data_pagamento',
    'valor_pago',
    'forma_pagamento',
    'observacao',
    'ativo',
    'criado_em'
  ];

  const schemaValidacao = [
    'TIPO',
    'UNIDADE',
    'CATEGORIA',
    'FORNECEDOR',
    'VALORKWH',
    'FORMA_PAGAMENTO'
  ];

  const ss = getDataSpreadsheet();

  function garantirAbaComSchema(nomeAba, schema) {
    let sheet = ss.getSheetByName(nomeAba);
    if (!sheet) {
      sheet = ss.insertSheet(nomeAba);
    }
    ensureSchema(sheet, schema);

    const totalColunas = sheet.getLastColumn();
    const headers = totalColunas > 0
      ? sheet.getRange(1, 1, 1, totalColunas).getValues()[0]
      : [];

    return {
      aba: nomeAba,
      colunas: totalColunas,
      headers
    };
  }

  const resumo = [
    garantirAbaComSchema('COMPRAS', schemaCompras),
    garantirAbaComSchema('DESPESAS_GERAIS', schemaDespesasGerais),
    garantirAbaComSchema('PAGAMENTOS', schemaPagamentos),
    garantirAbaComSchema('VALIDACAO', schemaValidacao)
  ];

  return {
    ok: true,
    atualizado_em: new Date(),
    abas: resumo
  };
}

function sincronizarSchemaCompras() {
  return sincronizarSchemaFinanceiro();
}

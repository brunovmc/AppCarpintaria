function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('CarpintariaZizu')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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

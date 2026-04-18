// Ambiente de dados
const DATA_SPREADSHEET_ID_PROD = "1j363esbdvygUO1s72bCz5u3Vhn2AnhiIHlmO4ZApowg";
const DATA_SPREADSHEET_ID_DEV = "1wqW2WPZvLWPr72Xsd_8ZAhAgbjvMw8A7b5Riapz6P4s";

// Fallback legado (evita quebrar referencias antigas)
const DATA_SPREADSHEET_ID = DATA_SPREADSHEET_ID_PROD;

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

function atualizarCachesManualmente() {
  const atualizadoEm = new Date();
  const referenciaAtual = Utilities.formatDate(
    atualizadoEm,
    Session.getScriptTimeZone(),
    "yyyy-MM",
  );

  function executar(nome, fn) {
    if (typeof fn !== "function") {
      return { ok: false, erro: `Funcao indisponivel: ${nome}` };
    }
    try {
      return fn();
    } catch (error) {
      return { ok: false, erro: error.message || String(error) };
    }
  }

  const resultado = {
    ok: true,
    atualizado_em: atualizadoEm,
    referencia_dashboard: referenciaAtual,
    caches: {
      validacoes: executar(
        "recarregarCacheValidacoes",
        recarregarCacheValidacoes,
      ),
      estoque: executar("recarregarCacheEstoque", recarregarCacheEstoque),
      compras: executar("recarregarCacheCompras", recarregarCacheCompras),
      vendas: executar("recarregarCacheVendas", recarregarCacheVendas),
      produtos: executar("recarregarCacheProdutos", recarregarCacheProdutos),
      producao: executar("recarregarCacheProducao", recarregarCacheProducao),
      despesas_gerais: executar(
        "recarregarCacheDespesasGerais",
        recarregarCacheDespesasGerais,
      ),
      pagamentos: executar(
        "recarregarCachePagamentos",
        recarregarCachePagamentos,
      ),
      usuarios: executar(
        "recarregarCacheUsuariosAcesso",
        recarregarCacheUsuariosAcesso,
      ),
    },
  };

  const limparDashboard = executar(
    "limparCacheDashboardFinanceiro",
    limparCacheDashboardFinanceiro,
  );
  const dashboard = executar("obterResumoDashboardFinanceiro", () =>
    obterResumoDashboardFinanceiro(referenciaAtual, true),
  );

  resultado.caches.dashboard = {
    limpeza: limparDashboard,
    recarregado: {
      ok: !dashboard?.erro,
      referencia: dashboard?.referencia || referenciaAtual,
    },
  };

  resultado.ok = Object.values(resultado.caches).every((item) => {
    if (!item || typeof item !== "object") return false;
    if (item.limpeza && item.recarregado) {
      const okLimpeza = item.limpeza?.ok !== false;
      const okRecarregado = item.recarregado?.ok !== false;
      return okLimpeza && okRecarregado;
    }
    return item.ok !== false;
  });

  return resultado;
}

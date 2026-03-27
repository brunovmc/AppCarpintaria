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

function prepararAbasFinanceiroVendas() {
  const ss = getDataSpreadsheet();

  function removerColunasPorHeader(sheet, headersParaRemover) {
    if (!sheet || !Array.isArray(headersParaRemover) || headersParaRemover.length === 0) {
      return { removidas: [] };
    }
    const lastCol = sheet.getLastColumn();
    if (lastCol <= 0) return { removidas: [] };

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const alvoUpper = headersParaRemover
      .map((h) => String(h || "").trim().toUpperCase())
      .filter((h) => h);
    const colunas = [];

    headers.forEach((h, idx) => {
      const label = String(h || "").trim().toUpperCase();
      if (alvoUpper.includes(label)) {
        colunas.push({ col: idx + 1, header: String(h || "").trim() });
      }
    });

    colunas
      .sort((a, b) => b.col - a.col)
      .forEach((item) => sheet.deleteColumn(item.col));

    return { removidas: colunas.map((item) => item.header) };
  }

  function garantirAbaComSchema(nomeAba, schema, opcoes) {
    if (!Array.isArray(schema) || schema.length === 0) {
      return { ok: false, aba: nomeAba, erro: "Schema invalido." };
    }
    const opts = opcoes || {};
    let sheet = ss.getSheetByName(nomeAba);
    if (!sheet) {
      sheet = ss.insertSheet(nomeAba);
    }

    const remocao = removerColunasPorHeader(sheet, opts.remover_headers || []);
    ensureSchema(sheet, schema);

    const headersAtuais = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map((h) => String(h || "").trim());

    return {
      ok: true,
      aba: nomeAba,
      colunas: headersAtuais.length,
      headers_atuais: headersAtuais,
      headers_esperados: schema,
      headers_removidos: remocao.removidas,
    };
  }

  const resultado = {
    ok: true,
    executado_em: new Date(),
    abas: {
      compras: garantirAbaComSchema("COMPRAS", COMPRAS_SCHEMA),
      despesas_gerais: garantirAbaComSchema("DESPESAS_GERAIS", DESPESAS_GERAIS_SCHEMA),
      vendas: garantirAbaComSchema("VENDAS", VENDAS_SCHEMA, {
        remover_headers: ["plano_recebimento", "data_entrega_prevista"],
      }),
      pagamentos: garantirAbaComSchema("PAGAMENTOS", PAGAMENTOS_SCHEMA),
      parcelas_financeiras: garantirAbaComSchema(
        "PARCELAS_FINANCEIRAS",
        PARCELAS_FINANCEIRAS_SCHEMA,
      ),
    },
  };

  resultado.ok = Object.values(resultado.abas).every((item) => item?.ok !== false);
  return resultado;
}

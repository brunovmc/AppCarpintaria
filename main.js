// Ambiente de dados
const DATA_SPREADSHEET_ID_PROD = "1j363esbdvygUO1s72bCz5u3Vhn2AnhiIHlmO4ZApowg";
const DATA_SPREADSHEET_ID_DEV = "1wqW2WPZvLWPr72Xsd_8ZAhAgbjvMw8A7b5Riapz6P4s";

// Fallback legado (evita quebrar referencias antigas)
const DATA_SPREADSHEET_ID = DATA_SPREADSHEET_ID_PROD;

function doGet() {
  tentarReconciliarEstruturaComprovantesDriveNoAcesso_();
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

/**
 * CONFIGURACAO TEMPORARIA — executar manualmente uma unica vez no editor do GAS.
 * Cria a estrutura de investimentos em PROD e DEV sem remover ou reordenar
 * colunas existentes. Depois da execucao validada, esta funcao pode ser removida.
 */
function configurarEstruturaInvestimentos() {
  const ambientes = [DB_ENV_PROD, DB_ENV_DEV];
  const resultados = [];
  ambientes.forEach(ambiente => {
    const resultado = executarComAmbienteBancoDados_(ambiente, () => {
      const ss = getDataSpreadsheet({ skipAccessCheck: true, targetEnv: ambiente });
      let investimentos = ss.getSheetByName(ABA_INVESTIMENTOS);
      if (!investimentos) investimentos = ss.insertSheet(ABA_INVESTIMENTOS);
      ensureSchema(investimentos, INVESTIMENTOS_SCHEMA);

      let validacao = ss.getSheetByName('VALIDACAO');
      if (!validacao) validacao = ss.insertSheet('VALIDACAO');
      ensureSchema(validacao, ['INVESTIDOR', 'TIPO_INVESTIMENTO']);
      SpreadsheetApp.flush();

      limparCacheInvestimentos();
      limparCacheValidacoes();
      limparCacheDashboardFinanceiro();

      return {
        ambiente,
        planilha_id: ss.getId(),
        aba_investimentos: investimentos.getName(),
        cabecalhos_investimentos: investimentos.getRange(
          1, 1, 1, investimentos.getLastColumn()
        ).getValues()[0],
        cabecalhos_validacao: validacao.getRange(
          1, 1, 1, validacao.getLastColumn()
        ).getValues()[0]
      };
    });
    resultados.push(resultado);
  });
  return { ok: true, ambientes: resultados };
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
      investimentos: executar(
        "recarregarCacheInvestimentos",
        recarregarCacheInvestimentos,
      ),
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
  const dashboardExecutivo = executar("obterDashboardExecutivoFinanceiro", () =>
    obterDashboardExecutivoFinanceiro(referenciaAtual, true),
  );

  resultado.caches.dashboard = {
    limpeza: limparDashboard,
    recarregado: {
      ok: !dashboard?.erro,
      referencia: dashboard?.referencia || referenciaAtual,
    },
    executivo: {
      ok: !dashboardExecutivo?.erro,
      referencia: dashboardExecutivo?.referencia || referenciaAtual,
      versao: dashboardExecutivo?.version || "",
    },
  };

  resultado.ok = Object.values(resultado.caches).every((item) => {
    if (!item || typeof item !== "object") return false;
    if (item.limpeza && item.recarregado) {
      const okLimpeza = item.limpeza?.ok !== false;
      const okRecarregado = item.recarregado?.ok !== false;
      const okExecutivo = item.executivo?.ok !== false;
      return okLimpeza && okRecarregado && okExecutivo;
    }
    return item.ok !== false;
  });

  return resultado;
}

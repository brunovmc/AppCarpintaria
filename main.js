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

function prepararAbasFinanceiroVendas() {
  assertCanWrite("Preparacao de abas financeiras");
  const ss = getDataSpreadsheet();

  function removerColunasPorHeader(sheet, headersParaRemover) {
    if (
      !sheet ||
      !Array.isArray(headersParaRemover) ||
      headersParaRemover.length === 0
    ) {
      return { removidas: [] };
    }
    const lastCol = sheet.getLastColumn();
    if (lastCol <= 0) return { removidas: [] };

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const alvoUpper = headersParaRemover
      .map((h) =>
        String(h || "")
          .trim()
          .toUpperCase(),
      )
      .filter((h) => h);
    const colunas = [];

    headers.forEach((h, idx) => {
      const label = String(h || "")
        .trim()
        .toUpperCase();
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
      despesas_gerais: garantirAbaComSchema(
        "DESPESAS_GERAIS",
        DESPESAS_GERAIS_SCHEMA,
      ),
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

  resultado.ok = Object.values(resultado.abas).every(
    (item) => item?.ok !== false,
  );
  return resultado;
}

const VALIDACAO_SCHEMA_PADRAO = [
  "TIPO",
  "COR",
  "UNIDADE",
  "CATEGORIA",
  "FORNECEDOR",
  "VALORKWH",
  "DESPESAS",
  "FORMA_PAGAMENTO",
  "PAGO_POR",
];

const VALIDACAO_TIPO_CATEGORIA_SCHEMA_PADRAO = [
  "TIPO",
  "CATEGORIA",
  "VENDAVEL",
];

const USUARIOS_SCHEMA_PADRAO = ["email", "role", "ativo"];

function garantirAbaComSchemaMain_(ss, nomeAba, schema, opcoes) {
  if (!ss) {
    return { ok: false, aba: nomeAba, erro: "Planilha invalida." };
  }
  if (!Array.isArray(schema) || schema.length === 0) {
    return { ok: false, aba: nomeAba, erro: "Schema invalido." };
  }

  const opts = opcoes || {};
  let sheet = ss.getSheetByName(nomeAba);
  if (!sheet) {
    sheet = ss.insertSheet(nomeAba);
  }

  const headersParaRemover = Array.isArray(opts.remover_headers)
    ? opts.remover_headers
    : [];
  if (headersParaRemover.length > 0 && sheet.getLastColumn() > 0) {
    const headersAtuais = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const alvoUpper = headersParaRemover
      .map((h) => String(h || "").trim().toUpperCase())
      .filter((h) => h);
    const colunas = [];

    headersAtuais.forEach((header, idx) => {
      const label = String(header || "").trim().toUpperCase();
      if (alvoUpper.includes(label)) {
        colunas.push(idx + 1);
      }
    });

    colunas
      .sort((a, b) => b - a)
      .forEach((col) => sheet.deleteColumn(col));
  }

  ensureSchema(sheet, schema);

  const headers = sheet.getLastColumn() > 0
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map((h) => String(h || "").trim())
    : [];

  return {
    ok: true,
    aba: nomeAba,
    colunas: headers.length,
    headers_atuais: headers,
    headers_esperados: schema,
  };
}

function listarDefinicoesAbasSistemaMain_() {
  return [
    { nome: "VALIDACAO", schema: VALIDACAO_SCHEMA_PADRAO },
    {
      nome: "VALIDACAO_TIPO_CATEGORIA",
      schema: VALIDACAO_TIPO_CATEGORIA_SCHEMA_PADRAO,
    },
    {
      nome: typeof ABA_USUARIOS_ACESSO === "string" ? ABA_USUARIOS_ACESSO : "USUARIOS",
      schema: USUARIOS_SCHEMA_PADRAO,
    },
    { nome: ABA_ESTOQUE, schema: ESTOQUE_SCHEMA },
    { nome: ABA_COMPRAS, schema: COMPRAS_SCHEMA },
    { nome: ABA_VENDAS, schema: VENDAS_SCHEMA, remover_headers: ["plano_recebimento", "data_entrega_prevista"] },
    { nome: ABA_DESPESAS_GERAIS, schema: DESPESAS_GERAIS_SCHEMA },
    { nome: ABA_PAGAMENTOS, schema: PAGAMENTOS_SCHEMA },
    { nome: ABA_PARCELAS_FINANCEIRAS, schema: PARCELAS_FINANCEIRAS_SCHEMA },
    { nome: ABA_PRODUTOS, schema: PRODUTOS_SCHEMA },
    { nome: ABA_PRODUTOS_COMPONENTES, schema: PRODUTOS_COMPONENTES_SCHEMA },
    { nome: ABA_PRODUTOS_ETAPAS, schema: PRODUTOS_ETAPAS_SCHEMA },
    { nome: ABA_PRODUTOS_RECEITAS, schema: PRODUTOS_RECEITAS_SCHEMA },
    {
      nome: ABA_PRODUTOS_RECEITAS_ENTRADAS,
      schema: PRODUTOS_RECEITAS_ENTRADAS_SCHEMA,
    },
    {
      nome: ABA_PRODUTOS_RECEITAS_SAIDAS,
      schema: PRODUTOS_RECEITAS_SAIDAS_SCHEMA,
    },
    { nome: ABA_PRODUCAO, schema: PRODUCAO_SCHEMA },
    { nome: ABA_PRODUCAO_ETAPAS, schema: PRODUCAO_ETAPAS_SCHEMA },
    { nome: ABA_PRODUCAO_CONSUMO, schema: PRODUCAO_CONSUMO_SCHEMA },
    { nome: ABA_PRODUCAO_MATERIAIS, schema: PRODUCAO_MATERIAIS_SCHEMA },
    {
      nome: ABA_PRODUCAO_MATERIAIS_PREVISTOS,
      schema: PRODUCAO_MATERIAIS_PREVISTOS_SCHEMA,
    },
    { nome: ABA_PRODUCAO_VINCULOS, schema: PRODUCAO_VINCULOS_SCHEMA },
    {
      nome: ABA_PRODUCAO_NECESSIDADES_ENTRADA,
      schema: PRODUCAO_NECESSIDADES_ENTRADA_SCHEMA,
    },
    {
      nome: ABA_PRODUCAO_RESERVAS_ENTRADA,
      schema: PRODUCAO_RESERVAS_ENTRADA_SCHEMA,
    },
    { nome: ABA_PRODUCAO_DESTINOS, schema: PRODUCAO_DESTINOS_SCHEMA },
    { nome: ABA_PRODUCAO_SAIDAS_LOTES, schema: PRODUCAO_SAIDAS_LOTES_SCHEMA },
  ];
}

function prepararEstruturaPlanilhaDadosPorAmbienteMain_(targetEnv) {
  const env = String(targetEnv || "").trim().toLowerCase() === "dev" ? "dev" : "prod";
  const ss = getDataSpreadsheet({ skipAccessCheck: true, targetEnv: env });
  const definicoes = listarDefinicoesAbasSistemaMain_();
  const abas = {};

  definicoes.forEach((def) => {
    abas[def.nome] = garantirAbaComSchemaMain_(
      ss,
      def.nome,
      def.schema,
      def,
    );
  });

  return {
    ok: Object.values(abas).every((item) => item?.ok !== false),
    ambiente: env,
    spreadsheet_id: ss.getId(),
    spreadsheet_nome: ss.getName(),
    total_abas: definicoes.length,
    abas,
  };
}

function prepararEstruturaPlanilhasProdDev() {
  assertCanWrite("Preparacao da estrutura das planilhas PROD e DEV");

  const ambientes = ["prod"];
  if (String(DATA_SPREADSHEET_ID_DEV || "").trim()) {
    ambientes.push("dev");
  }

  const resultado = {
    ok: true,
    executado_em: new Date(),
    ambientes: {},
  };

  ambientes.forEach((env) => {
    resultado.ambientes[env] = prepararEstruturaPlanilhaDadosPorAmbienteMain_(env);
  });

  resultado.ok = Object.values(resultado.ambientes).every(
    (item) => item?.ok !== false,
  );
  return resultado;
}

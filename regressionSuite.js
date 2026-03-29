const REGRESSION_SPRINT3_VERSION = 'v1';

function runRegressionSuiteSprint3() {
  const inicio = Date.now();
  const casos = [
    ['sheetRepo:updateById', regressionSheetRepoUpdateById],
    ['cache:app+validacoes+producao', regressionCacheCore],
    ['financeiro:parcelas', regressionFinanceiroParcelas],
    ['producao:agregacao+custo', regressionProducaoCalculos]
  ];

  const resultados = casos.map(([nome, fn]) => executarCasoRegressao(nome, fn));
  const ok = resultados.every(r => r.ok !== false);

  return {
    ok,
    suite: 'sprint3',
    version: REGRESSION_SPRINT3_VERSION,
    executado_em: new Date(),
    duracao_ms: Date.now() - inicio,
    resultados
  };
}

function executarSuiteRegressaoSprint3() {
  return runRegressionSuiteSprint3();
}

function executarSuiteRegressaoSprint3ComLog() {
  const resultado = executarSuiteRegressaoSprint3();
  const json = JSON.stringify(resultado, null, 2);
  try {
    console.log(json);
  } catch (error) {
    // sem acao
  }
  try {
    Logger.log(json);
  } catch (error) {
    // sem acao
  }
  return resultado;
}

function executarCasoRegressao(nome, fn) {
  const inicio = Date.now();
  try {
    const detalhes = fn() || {};
    return {
      nome,
      ok: true,
      duracao_ms: Date.now() - inicio,
      detalhes
    };
  } catch (error) {
    return {
      nome,
      ok: false,
      duracao_ms: Date.now() - inicio,
      erro: String(error?.message || error || 'Falha desconhecida')
    };
  }
}

function assertRegressao(condicao, mensagem) {
  if (!condicao) {
    throw new Error(mensagem || 'Falha de assercao');
  }
}

function assertAproxRegressao(atual, esperado, tolerancia, mensagem) {
  const tol = Number.isFinite(Number(tolerancia)) ? Number(tolerancia) : 0.000001;
  const a = Number(atual);
  const e = Number(esperado);
  if (!Number.isFinite(a) || !Number.isFinite(e) || Math.abs(a - e) > tol) {
    throw new Error(mensagem || `Valor fora da tolerancia. atual=${a} esperado=${e} tol=${tol}`);
  }
}

function regressionSheetRepoUpdateById() {
  if (typeof assertCanWrite === 'function') {
    assertCanWrite('Suite regressao sprint3');
  }

  const ss = getDataSpreadsheet({ skipAccessCheck: true });
  const nomeAba = `__REGTEST_S3_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
  const schema = ['ID', 'NOME', 'VALOR'];
  const payloadInicial = { ID: 'ABC-001', NOME: 'Item inicial', VALOR: 10 };
  let sheet = null;

  try {
    sheet = ss.insertSheet(nomeAba);
    ensureSchema(sheet, schema);
    insert(nomeAba, payloadInicial, schema);

    const okTrim = updateById(nomeAba, 'ID', ' ABC-001 ', { NOME: 'Atualizado trim' }, schema);
    assertRegressao(okTrim === true, 'updateById deveria atualizar ID com trim');

    const okNumero = updateById(nomeAba, 'ID', 'ABC-001', { VALOR: 20 }, schema);
    assertRegressao(okNumero === true, 'updateById deveria atualizar o registro existente');

    const okAusente = updateById(nomeAba, 'ID', 'NAO-EXISTE', { VALOR: 99 }, schema);
    assertRegressao(okAusente === false, 'updateById deveria retornar false para ID ausente');

    const rows = rowsToObjects(sheet);
    assertRegressao(rows.length === 1, 'Aba de teste deveria conter 1 linha');
    assertRegressao(String(rows[0].NOME || '') === 'Atualizado trim', 'Nome nao atualizado conforme esperado');
    assertRegressao(Number(rows[0].VALOR) === 20, 'Valor nao atualizado conforme esperado');

    return {
      aba_teste: nomeAba,
      linhas: rows.length
    };
  } finally {
    if (sheet) {
      ss.deleteSheet(sheet);
    }
  }
}

function regressionCacheCore() {
  const scopeTeste = `__REGTEST_CACHE_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
  const payloadAppCache = { ok: true, numero: 42, itens: [1, 2, 3] };

  appCacheRemove(scopeTeste);
  const put = appCachePutJson(scopeTeste, payloadAppCache, 60);
  assertRegressao(put?.ok !== false, 'Falha ao salvar app cache');

  const lido = appCacheGetJson(scopeTeste);
  assertRegressao(!!lido && lido.numero === 42, 'Leitura de app cache inconsistente');

  appCacheRemove(scopeTeste);
  const aposRemocao = appCacheGetJson(scopeTeste);
  assertRegressao(aposRemocao === null, 'Cache deveria estar limpo apos remocao');

  const payloadValidacoes = {
    tipos: ['TIPO_TESTE'],
    categorias: ['CAT_TESTE'],
    categoriasPorTipo: { TIPO_TESTE: ['CAT_TESTE'] }
  };
  limparCacheValidacoes();
  salvarValidacoesNoCache(payloadValidacoes);
  const validacoesCache = lerValidacoesDoCache();
  assertRegressao(
    Array.isArray(validacoesCache?.tipos) && validacoesCache.tipos[0] === 'TIPO_TESTE',
    'Cache de validacoes nao refletiu payload esperado'
  );
  limparCacheValidacoes();

  limparCacheProducao();
  salvarCacheProducao([{ producao_id: 'P-1', ativo: 'true' }]);
  const cacheProducao = lerCacheProducao();
  assertRegressao(Array.isArray(cacheProducao) && cacheProducao.length === 1, 'Cache de producao inconsistente');
  limparCacheProducao();

  return {
    scope_app_cache: scopeTeste,
    validacoes_cache_ok: true,
    producao_cache_ok: true
  };
}

function regressionFinanceiroParcelas() {
  const parcelasCredito = normalizarParcelasFinanceiro('3', 'Credito');
  assertRegressao(parcelasCredito === 3, 'Credito deveria aceitar 3 parcelas');

  const parcelasPix = normalizarParcelasFinanceiro('5', 'PIX');
  assertRegressao(parcelasPix === 1, 'PIX nao parcelado deveria normalizar para 1 parcela');

  const planoPadrao = normalizarParcelasDetalhePayloadFinanceiro('', 3, '2026-01-15', 300);
  assertRegressao(Array.isArray(planoPadrao) && planoPadrao.length === 3, 'Plano padrao deveria ter 3 parcelas');
  const somaPadrao = planoPadrao.reduce((acc, p) => acc + Number(p.valor || 0), 0);
  assertAproxRegressao(somaPadrao, 300, 0.02, 'Soma do plano padrao deveria fechar em 300');

  const detalheManual = [
    { data: '2026-01-15', valor: 100 },
    { data: '2026-02-15', valor: 100 },
    { data: '2026-03-15', valor: 100 }
  ];
  const planoManual = normalizarParcelasDetalhePayloadFinanceiro(detalheManual, 3, '2026-01-15', 300);
  const somaManual = planoManual.reduce((acc, p) => acc + Number(p.valor || 0), 0);
  assertAproxRegressao(somaManual, 300, 0.02, 'Soma do plano manual deveria fechar em 300');

  return {
    parcelas_credito: parcelasCredito,
    parcelas_pix: parcelasPix,
    plano_padrao_parcelas: planoPadrao.length
  };
}

function regressionProducaoCalculos() {
  const saidas = [
    { nome_saida: 'Porta', tipo_item: 'PRODUTO', categoria: 'PECA', unidade: 'UN', quantidade: 2 },
    { nome_saida: 'porta', tipo_item: 'produto', categoria: 'peca', unidade: 'un', quantidade: '3' },
    { nome_saida: 'Sobra', tipo_item: 'OUTROS', categoria: '', unidade: 'KG', quantidade: '1.5' }
  ];

  const agrupadas = agruparSaidasReceitaParaEstoque(saidas);
  assertRegressao(Array.isArray(agrupadas) && agrupadas.length >= 2, 'Agrupamento de saidas deveria gerar pelo menos 2 grupos');

  const porta = agrupadas.find(item => normalizarTextoProducao(item?.nome_saida) === 'PORTA');
  assertRegressao(!!porta, 'Grupo Porta nao encontrado');
  assertAproxRegressao(porta.quantidade, 5, 0.000001, 'Quantidade agregada da Porta deveria ser 5');

  const custoUnit = calcularCustoUnitarioSaidas(50, [{ quantidade: 2 }, { quantidade: 3 }]);
  assertAproxRegressao(custoUnit, 10, 0.000001, 'Custo unitario esperado era 10');

  const custoProdutoSemCusto = obterCustoUnitarioEstoqueItem({ tipo: 'PRODUTO', custo_unitario: '' });
  assertAproxRegressao(custoProdutoSemCusto, 0, 0.000001, 'Produto sem custo explicito deveria retornar 0');

  const custoMadeiraFallback = obterCustoUnitarioEstoqueItem({ tipo: 'MADEIRA', custo_unitario: '', valor_unit: 12.5 });
  assertAproxRegressao(custoMadeiraFallback, 12.5, 0.000001, 'Fallback de valor_unit nao aplicado');

  return {
    grupos_saida: agrupadas.length,
    custo_unitario: custoUnit
  };
}

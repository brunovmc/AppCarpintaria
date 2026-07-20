const REGRESSION_SPRINT3_VERSION = 'v3';
const REGRESSION_PURGE_MARKER = '__REGTEST';

function listarCasosRegressaoSprint3_() {
  return [
    ['sheetRepo:updateById', regressionSheetRepoUpdateById],
    ['ambiente:override-execucao', regressionAmbienteOverrideExecucao],
    ['acesso:bootstrap-contexto', regressionBootstrapContextoInicial],
    ['cache:app+validacoes+producao', regressionCacheCore],
    ['financeiro:parcelas', regressionFinanceiroParcelas],
    ['producao:agregacao+custo', regressionProducaoCalculos],
    ['producao:reserva-quantidade', regressionProducaoReservaQuantidade],
    ['producao:madeira-reserva-exclusiva', regressionProducaoMadeiraReservaExclusiva],
    ['producao:consumo-soft-delete', regressionProducaoConsumoSoftDelete],
    ['producao:editar-op-libera-reservas', regressionProducaoAtualizacaoRegeraNecessidades],
    ['vendas:saldo-disponivel-com-reserva', regressionVendaSaldoReservado],
    ['documentos-compra:rateio+quantidade', regressionDocumentoCompraCalculos],
    ['comprovantes:matching+schema', regressionComprovantesFinanceiros],
    ['recebimentos:matching+schema', regressionComprovantesRecebimentos],
    ['comprovantes-drive:estrutura', regressionEstruturaComprovantesDrive]
  ];
}

function listarGruposRegressaoSprint3_() {
  return {
    parte1: [
      'sheetRepo:updateById',
      'ambiente:override-execucao',
      'acesso:bootstrap-contexto',
      'cache:app+validacoes+producao',
      'financeiro:parcelas',
      'producao:agregacao+custo'
    ],
    parte2: [
      'producao:reserva-quantidade',
      'producao:madeira-reserva-exclusiva'
    ],
    parte3: [
      'producao:consumo-soft-delete',
      'producao:editar-op-libera-reservas'
    ],
    parte4: [
      'vendas:saldo-disponivel-com-reserva'
    ],
    parte5: [
      'documentos-compra:rateio+quantidade'
    ],
    parte6: [
      'comprovantes:matching+schema',
      'recebimentos:matching+schema',
      'comprovantes-drive:estrutura'
    ]
  };
}

function selecionarCasosRegressaoSprint3_(nomesCasos) {
  const todos = listarCasosRegressaoSprint3_();
  if (!Array.isArray(nomesCasos) || nomesCasos.length === 0) {
    return todos;
  }

  const mapa = {};
  todos.forEach(([nome, fn]) => {
    mapa[nome] = fn;
  });

  return nomesCasos.map((nome) => {
    const chave = String(nome || '').trim();
    const fn = mapa[chave];
    if (typeof fn !== 'function') {
      throw new Error(`Caso de regressao nao encontrado: ${chave}`);
    }
    return [chave, fn];
  });
}

function runRegressionSuiteSprint3(opcoes) {
  const inicio = Date.now();
  const opts = opcoes || {};
  const casos = selecionarCasosRegressaoSprint3_(opts.casos);

  const resultados = casos.map(([nome, fn]) => executarCasoRegressao(nome, fn));
  const ok = resultados.every(r => r.ok !== false);

  return {
    ok,
    suite: String(opts.suite || 'sprint3'),
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

function executarSuiteRegressaoSprint3Parte(parte) {
  const grupos = listarGruposRegressaoSprint3_();
  const chave = String(parte || '').trim().toLowerCase();
  const casos = grupos[chave];
  if (!Array.isArray(casos) || casos.length === 0) {
    throw new Error(`Parte da suite nao encontrada: ${String(parte || '')}`);
  }
  return runRegressionSuiteSprint3({
    suite: `sprint3:${chave}`,
    casos
  });
}

function executarSuiteRegressaoSprint3ParteComLog(parte) {
  const resultado = executarSuiteRegressaoSprint3Parte(parte);
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

function executarRegressaoComLogNoAmbiente_(ambiente, executor) {
  const envAlvo = String(ambiente || '').trim().toLowerCase() === 'dev' ? 'dev' : 'prod';
  const contextoAnterior = (typeof obterContextoBancoDados === 'function')
    ? (obterContextoBancoDados() || {})
    : {};
  const ambienteAnterior = String(contextoAnterior?.selected_env || contextoAnterior?.effective_env || 'prod')
    .trim()
    .toLowerCase() === 'dev'
      ? 'dev'
      : 'prod';

  let ambienteRestaurado = ambienteAnterior;

  try {
    if (typeof definirContextoBancoDados === 'function') {
      const contextoDefinido = definirContextoBancoDados(envAlvo) || {};
      const ambienteEfetivo = String(
        contextoDefinido?.selected_env || contextoDefinido?.effective_env || 'prod'
      ).trim().toLowerCase() === 'dev'
        ? 'dev'
        : 'prod';
      if (ambienteEfetivo !== envAlvo) {
        throw new Error(`Nao foi possivel executar a suite no ambiente ${envAlvo.toUpperCase()}.`);
      }
    }

    const resultado = (typeof executor === 'function')
      ? (executor() || {})
      : executarSuiteRegressaoSprint3ComLog();
    return {
      ...resultado,
      ambiente_execucao: envAlvo,
      ambiente_anterior: ambienteAnterior
    };
  } finally {
    if (typeof definirContextoBancoDados === 'function') {
      const contextoRestaurado = definirContextoBancoDados(ambienteAnterior) || {};
      ambienteRestaurado = String(
        contextoRestaurado?.selected_env || contextoRestaurado?.effective_env || ambienteAnterior
      ).trim().toLowerCase() === 'dev'
        ? 'dev'
        : 'prod';
    }
    try {
      const mensagem = `Regression suite context restored to ${ambienteRestaurado.toUpperCase()}.`;
      console.log(mensagem);
      Logger.log(mensagem);
    } catch (error) {
      // sem acao
    }
  }
}

function executarSuiteRegressaoSprint3ComLogNoAmbiente_(ambiente) {
  return executarRegressaoComLogNoAmbiente_(ambiente, executarSuiteRegressaoSprint3ComLog);
}

function executarSuiteRegressaoSprint3ComLogDev() {
  return executarSuiteRegressaoSprint3ComLogNoAmbiente_('dev');
}

function executarSuiteRegressaoSprint3ComLogProd() {
  return executarSuiteRegressaoSprint3ComLogNoAmbiente_('prod');
}

function executarSuiteRegressaoSprint3Parte1ComLogDev() {
  return executarRegressaoComLogNoAmbiente_('dev', () => executarSuiteRegressaoSprint3ParteComLog('parte1'));
}

function executarSuiteRegressaoSprint3Parte2ComLogDev() {
  return executarRegressaoComLogNoAmbiente_('dev', () => executarSuiteRegressaoSprint3ParteComLog('parte2'));
}

function executarSuiteRegressaoSprint3Parte3ComLogDev() {
  return executarRegressaoComLogNoAmbiente_('dev', () => executarSuiteRegressaoSprint3ParteComLog('parte3'));
}

function executarSuiteRegressaoSprint3Parte4ComLogDev() {
  return executarRegressaoComLogNoAmbiente_('dev', () => executarSuiteRegressaoSprint3ParteComLog('parte4'));
}

function executarSuiteRegressaoSprint3Parte1ComLogProd() {
  return executarRegressaoComLogNoAmbiente_('prod', () => executarSuiteRegressaoSprint3ParteComLog('parte1'));
}

function executarSuiteRegressaoSprint3Parte2ComLogProd() {
  return executarRegressaoComLogNoAmbiente_('prod', () => executarSuiteRegressaoSprint3ParteComLog('parte2'));
}

function executarSuiteRegressaoSprint3Parte3ComLogProd() {
  return executarRegressaoComLogNoAmbiente_('prod', () => executarSuiteRegressaoSprint3ParteComLog('parte3'));
}

function executarSuiteRegressaoSprint3Parte4ComLogProd() {
  return executarRegressaoComLogNoAmbiente_('prod', () => executarSuiteRegressaoSprint3ParteComLog('parte4'));
}

function normalizarCampoPurgeRegressao_(campo) {
  return String(campo || '').trim();
}

function normalizarValorPurgeRegressao_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).trim();
}

function contemMarcadorPurgeRegressao_(valor) {
  const texto = normalizarValorPurgeRegressao_(valor).toUpperCase();
  return !!texto && texto.indexOf(REGRESSION_PURGE_MARKER) !== -1;
}

function listarAliasesCampoPurgeRegressao_(campo) {
  const chave = normalizarCampoPurgeRegressao_(campo);
  switch (chave) {
    case 'ID':
    case 'estoque_id':
      return ['ID', 'estoque_id'];
    case 'produto_id':
    case 'produto_ref_id':
      return ['produto_id', 'produto_ref_id'];
    case 'receita_id':
    case 'receita_ref_id':
    case 'parent_receita_id':
      return ['receita_id', 'receita_ref_id', 'parent_receita_id'];
    case 'producao_id':
    case 'op_id':
    case 'origem_id':
      return ['producao_id', 'op_id', 'origem_id'];
    default:
      return [chave];
  }
}

function adicionarValorRelacionadoPurgeRegressao_(mapa, campo, valor) {
  const valorNorm = normalizarValorPurgeRegressao_(valor);
  if (!valorNorm) return false;

  let alterou = false;
  listarAliasesCampoPurgeRegressao_(campo).forEach((alias) => {
    const chave = normalizarCampoPurgeRegressao_(alias);
    if (!chave) return;
    if (!mapa[chave]) {
      mapa[chave] = {};
    }
    if (!mapa[chave][valorNorm]) {
      mapa[chave][valorNorm] = true;
      alterou = true;
    }
  });
  return alterou;
}

function coletarRelacionamentosLinhaPurgeRegressao_(headers, rowValues, relacionados) {
  let alterou = false;
  (Array.isArray(headers) ? headers : []).forEach((header, idx) => {
    const campo = normalizarCampoPurgeRegressao_(header);
    if (!campo) return;
    if (campo === 'id' || campo === 'ativo' || campo === 'criado_em' || campo === 'data') return;
    if (campo === 'ID' || campo.endsWith('_id')) {
      alterou = adicionarValorRelacionadoPurgeRegressao_(relacionados, campo, rowValues[idx]) || alterou;
    }
  });
  return alterou;
}

function linhaTemRelacionamentoPurgeRegressao_(headers, rowValues, relacionados) {
  return (Array.isArray(headers) ? headers : []).some((header, idx) => {
    const campo = normalizarCampoPurgeRegressao_(header);
    if (!campo) return false;
    const valor = normalizarValorPurgeRegressao_(rowValues[idx]);
    if (!valor) return false;
    const conjunto = relacionados[campo];
    return !!(conjunto && conjunto[valor]);
  });
}

function agruparLinhasContiguasPurgeRegressao_(linhas) {
  const ordenadas = Array.from(new Set((Array.isArray(linhas) ? linhas : []).filter((n) => Number(n) > 1)))
    .map((n) => Number(n))
    .sort((a, b) => a - b);

  if (ordenadas.length === 0) return [];

  const grupos = [];
  let inicio = ordenadas[0];
  let anterior = ordenadas[0];

  for (let i = 1; i < ordenadas.length; i++) {
    const atual = ordenadas[i];
    if (atual === anterior + 1) {
      anterior = atual;
      continue;
    }
    grupos.push({ inicio, quantidade: anterior - inicio + 1 });
    inicio = atual;
    anterior = atual;
  }
  grupos.push({ inicio, quantidade: anterior - inicio + 1 });
  return grupos;
}

function limparCachesPosPurgeRegressao_() {
  const funcoes = [
    'limparCacheValidacoes',
    'limparCacheEstoque',
    'limparCacheCompras',
    'limparCacheVendas',
    'limparCacheProdutos',
    'limparCacheProducao',
    'limparCacheDespesasGerais',
    'limparCachePagamentos',
    'limparCacheUsuariosAcesso',
    'limparCacheDashboardFinanceiro'
  ];

  const caches = {};
  funcoes.forEach((nome) => {
    if (typeof globalThis?.[nome] === 'function') {
      try {
        globalThis[nome]();
        caches[nome] = { ok: true };
      } catch (error) {
        caches[nome] = { ok: false, erro: String(error?.message || error || 'Falha ao limpar cache') };
      }
    }
  });
  return caches;
}

function purgarResiduosRegtestAmbienteAtual_() {
  assertCanWrite('Purge de residuos __REGTEST_');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getDataSpreadsheet({ skipAccessCheck: true });
    const sheets = ss.getSheets();
    const relacionados = {};
    const deletarPorAba = {};
    const analise = [];
    const abasTemporarias = [];

    sheets.forEach((sheet) => {
      const nomeAba = sheet.getName();
      if (contemMarcadorPurgeRegressao_(nomeAba)) {
        abasTemporarias.push(nomeAba);
      }

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow <= 1 || lastCol <= 0) {
        analise.push({
          sheet,
          nomeAba,
          headers: [],
          linhas: []
        });
        return;
      }

      const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      const headers = values[0] || [];
      const linhas = [];

      for (let i = 1; i < values.length; i++) {
        const rowValues = values[i];
        const rowNumber = i + 1;
        const temMarcador = rowValues.some((cell) => contemMarcadorPurgeRegressao_(cell));
        linhas.push({
          rowNumber,
          values: rowValues,
          temMarcador,
          deletar: false
        });
        if (temMarcador) {
          coletarRelacionamentosLinhaPurgeRegressao_(headers, rowValues, relacionados);
        }
      }

      analise.push({
        sheet,
        nomeAba,
        headers,
        linhas
      });
    });

    let alterou = true;
    while (alterou) {
      alterou = false;
      analise.forEach((item) => {
        item.linhas.forEach((linha) => {
          if (linha.deletar) return;
          if (!linha.temMarcador && !linhaTemRelacionamentoPurgeRegressao_(item.headers, linha.values, relacionados)) {
            return;
          }
          linha.deletar = true;
          deletarPorAba[item.nomeAba] = deletarPorAba[item.nomeAba] || [];
          deletarPorAba[item.nomeAba].push(linha.rowNumber);
          alterou = coletarRelacionamentosLinhaPurgeRegressao_(item.headers, linha.values, relacionados) || alterou;
        });
      });
    }

    const resumoAbas = {};
    Object.keys(deletarPorAba).forEach((nomeAba) => {
      const linhas = deletarPorAba[nomeAba];
      const item = analise.find((entry) => entry.nomeAba === nomeAba);
      if (!item || !Array.isArray(linhas) || linhas.length === 0) return;

      const grupos = agruparLinhasContiguasPurgeRegressao_(linhas);
      for (let i = grupos.length - 1; i >= 0; i--) {
        const grupo = grupos[i];
        item.sheet.deleteRows(grupo.inicio, grupo.quantidade);
      }

      resumoAbas[nomeAba] = {
        linhas_removidas: Array.from(new Set(linhas)).length
      };
    });

    const abasRemovidas = [];
    abasTemporarias.forEach((nomeAba) => {
      const sheet = ss.getSheetByName(nomeAba);
      if (!sheet) return;
      if (ss.getSheets().length <= 1) return;
      ss.deleteSheet(sheet);
      abasRemovidas.push(nomeAba);
    });

    SpreadsheetApp.flush();

    return {
      ok: true,
      ambiente: (typeof obterContextoBancoDados === 'function'
        ? (obterContextoBancoDados()?.effective_env || obterContextoBancoDados()?.selected_env || '')
        : ''),
      marker: REGRESSION_PURGE_MARKER,
      spreadsheet_id: ss.getId(),
      spreadsheet_nome: ss.getName(),
      abas_analisadas: sheets.length,
      abas_com_linhas_removidas: Object.keys(resumoAbas).length,
      abas_removidas: abasRemovidas,
      total_linhas_removidas: Object.values(resumoAbas).reduce((acc, item) => acc + Number(item?.linhas_removidas || 0), 0),
      detalhes_abas: resumoAbas,
      caches_limpos: limparCachesPosPurgeRegressao_()
    };
  } finally {
    lock.releaseLock();
  }
}

function purgarResiduosRegtestAtualComLog() {
  const resultado = purgarResiduosRegtestAmbienteAtual_();
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

function purgarResiduosRegtestDevComLog() {
  return executarRegressaoComLogNoAmbiente_('dev', purgarResiduosRegtestAtualComLog);
}

function purgarResiduosRegtestProdComLog() {
  return executarRegressaoComLogNoAmbiente_('prod', purgarResiduosRegtestAtualComLog);
}

function purgarResiduosRegtestDevEProdComLog() {
  const inicio = Date.now();
  const dev = purgarResiduosRegtestDevComLog();
  const prod = purgarResiduosRegtestProdComLog();
  const resultado = {
    ok: dev?.ok !== false && prod?.ok !== false,
    operacao: 'purge_regtest_dev_prod',
    marker: REGRESSION_PURGE_MARKER,
    executado_em: new Date(),
    duracao_ms: Date.now() - inicio,
    ambientes: {
      dev,
      prod
    }
  };
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

function regressionAmbienteOverrideExecucao() {
  const ambienteOriginal = DB_ENV_EXECUTION_OVERRIDE_;
  try {
    DB_ENV_EXECUTION_OVERRIDE_ = DB_ENV_PROD;
    const observado = executarComAmbienteBancoDados_(DB_ENV_DEV, ambiente => ({
      argumento: ambiente,
      override: DB_ENV_EXECUTION_OVERRIDE_
    }));
    assertRegressao(
      observado.argumento === DB_ENV_DEV && observado.override === DB_ENV_DEV,
      'Override deve usar o ambiente solicitado durante a execucao.'
    );
    assertRegressao(
      DB_ENV_EXECUTION_OVERRIDE_ === DB_ENV_PROD,
      'Override deve restaurar o ambiente anterior apos sucesso.'
    );

    const aninhado = executarComAmbienteBancoDados_(DB_ENV_DEV, () => {
      const externo = DB_ENV_EXECUTION_OVERRIDE_;
      const interno = executarComAmbienteBancoDados_(DB_ENV_PROD, () => DB_ENV_EXECUTION_OVERRIDE_);
      return { externo, interno, restaurado: DB_ENV_EXECUTION_OVERRIDE_ };
    });
    assertRegressao(
      aninhado.externo === DB_ENV_DEV && aninhado.interno === DB_ENV_PROD && aninhado.restaurado === DB_ENV_DEV,
      'Override aninhado deve restaurar o ambiente externo.'
    );

    let lancou = false;
    try {
      executarComAmbienteBancoDados_(DB_ENV_DEV, () => {
        throw new Error('falha esperada');
      });
    } catch (error) {
      lancou = String(error?.message || '') === 'falha esperada';
    }
    assertRegressao(lancou, 'Override deve propagar o erro original.');
    assertRegressao(
      DB_ENV_EXECUTION_OVERRIDE_ === DB_ENV_PROD,
      'Override deve restaurar o ambiente anterior apos erro.'
    );
    return { ok: true };
  } finally {
    DB_ENV_EXECUTION_OVERRIDE_ = ambienteOriginal;
  }
}

function regressionBootstrapContextoInicial() {
  const ambienteOriginal = getUserDbEnvironment_();
  const acessoNegado = {
    email: 'negado@regressao.local',
    role: 'viewer',
    ativo: false,
    can_read: false,
    can_write: false,
    read_only: true,
    motivo: 'Fixture de acesso negado.'
  };
  const acessoAdmin = {
    email: 'admin@regressao.local',
    role: 'admin',
    ativo: true,
    can_read: true,
    can_write: true,
    read_only: false,
    motivo: ''
  };
  const acessoViewer = {
    email: 'viewer@regressao.local',
    role: 'viewer',
    ativo: true,
    can_read: true,
    can_write: false,
    read_only: true,
    motivo: 'Perfil somente leitura.'
  };

  try {
    const ambienteAntesNegado = getUserDbEnvironment_();
    const negado = obterContextoInicialAplicacaoComAcesso_(DB_ENV_DEV, acessoNegado);
    assertRegressao(negado?.acesso === acessoNegado, 'Bootstrap deve preservar o contexto de acesso negado.');
    assertRegressao(negado?.banco_dados === null, 'Acesso negado nao deve receber contexto de banco.');
    assertRegressao(
      getUserDbEnvironment_() === ambienteAntesNegado,
      'Acesso negado nao deve alterar a preferencia de ambiente.'
    );

    const adminProd = obterContextoInicialAplicacaoComAcesso_(DB_ENV_PROD, acessoAdmin);
    assertRegressao(adminProd?.banco_dados?.can_toggle === true, 'Admin deve poder alternar o banco.');
    assertRegressao(
      adminProd?.banco_dados?.effective_env === DB_ENV_PROD,
      'Bootstrap admin PROD deve aplicar o ambiente PROD.'
    );

    const adminDev = obterContextoInicialAplicacaoComAcesso_(DB_ENV_DEV, acessoAdmin);
    assertRegressao(
      adminDev?.banco_dados?.effective_env === DB_ENV_DEV,
      'Bootstrap admin DEV deve aplicar o ambiente DEV.'
    );

    const viewerDev = obterContextoInicialAplicacaoComAcesso_(DB_ENV_DEV, acessoViewer);
    assertRegressao(viewerDev?.banco_dados?.can_toggle === false, 'Viewer nao deve poder alternar o banco.');
    assertRegressao(
      viewerDev?.banco_dados?.effective_env === DB_ENV_PROD,
      'Bootstrap viewer deve ser forcado a PROD mesmo quando DEV foi solicitado.'
    );

    return {
      acesso_negado_sem_contexto: true,
      admin_prod: adminProd.banco_dados.effective_env,
      admin_dev: adminDev.banco_dados.effective_env,
      viewer_forcado: viewerDev.banco_dados.effective_env
    };
  } finally {
    setUserDbEnvironment_(ambienteOriginal);
  }
}

function assertThrowsRegressao(fn, trechoEsperado, mensagem) {
  let erroCapturado = '';
  try {
    fn();
  } catch (error) {
    erroCapturado = String(error?.message || error || '');
  }

  if (!erroCapturado) {
    throw new Error(mensagem || 'Era esperado erro, mas a operacao foi aceita.');
  }
  if (trechoEsperado && erroCapturado.indexOf(trechoEsperado) === -1) {
    throw new Error(
      mensagem ||
      `Mensagem inesperada. Esperado conter "${trechoEsperado}", recebido "${erroCapturado}".`
    );
  }
  return erroCapturado;
}

function criarContextoFixtureRegressao(prefixo) {
  return {
    token: `__REGTEST_${String(prefixo || 'GERAL').trim().toUpperCase()}_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
    cleanups: []
  };
}

function registrarCleanupRegressao(contexto, fn) {
  if (!contexto || typeof fn !== 'function') return;
  contexto.cleanups.push(fn);
}

function executarCleanupRegressao(contexto) {
  const erros = [];
  const lista = contexto && Array.isArray(contexto.cleanups) ? contexto.cleanups : [];

  for (let i = lista.length - 1; i >= 0; i--) {
    try {
      lista[i]();
    } catch (error) {
      erros.push(String(error?.message || error || 'Falha no cleanup'));
    }
  }

  if (erros.length > 0) {
    const msg = `Cleanup da regressao com falhas: ${erros.join(' | ')}`;
    try {
      Logger.log(msg);
    } catch (error) {
      // sem acao
    }
  }

  return erros;
}

function inativarRegistroRegressao(sheetName, idField, id, schema) {
  const registroId = String(id || '').trim();
  if (!registroId) return false;
  try {
    return updateById(sheetName, idField, registroId, { ativo: false }, schema);
  } catch (error) {
    return false;
  }
}

function inativarRegistrosPorCampoRegressao(sheetName, campo, valor, idField, schema) {
  const sheet = getDataSpreadsheet({ skipAccessCheck: true }).getSheetByName(sheetName);
  if (!sheet) return 0;

  const alvo = String(valor || '').trim();
  if (!alvo) return 0;

  const rows = rowsToObjects(sheet).filter(row =>
    String(row?.[campo] || '').trim() === alvo &&
    String(row?.ativo).toLowerCase() !== 'false'
  );

  rows.forEach(row => {
    const id = String(row?.[idField] || '').trim();
    if (!id) return;
    inativarRegistroRegressao(sheetName, idField, id, schema);
  });

  return rows.length;
}

function obterRegistroPorCampoRegressao(sheetName, campo, valor) {
  const sheet = getDataSpreadsheet({ skipAccessCheck: true }).getSheetByName(sheetName);
  if (!sheet) return null;

  const alvo = String(valor || '').trim();
  return rowsToObjects(sheet).find(row => String(row?.[campo] || '').trim() === alvo) || null;
}

function registrarCleanupItemEstoqueRegressao(contexto, estoqueId) {
  const id = String(estoqueId || '').trim();
  if (!id) return;
  registrarCleanupRegressao(contexto, () => {
    inativarRegistroRegressao(ABA_ESTOQUE, 'ID', id, ESTOQUE_SCHEMA);
  });
}

function registrarCleanupProducaoRegressao(contexto, producaoId) {
  const id = String(producaoId || '').trim();
  if (!id) return;

  registrarCleanupRegressao(contexto, () => {
    try {
      deletarProducao(id);
    } catch (error) {
      inativarRegistroRegressao(ABA_PRODUCAO, 'producao_id', id, PRODUCAO_SCHEMA);
    }

    inativarRegistrosPorCampoRegressao(ABA_PRODUCAO_ETAPAS, 'producao_id', id, 'id', PRODUCAO_ETAPAS_SCHEMA);
    inativarRegistrosPorCampoRegressao(ABA_PRODUCAO_DESTINOS, 'producao_id', id, 'id', PRODUCAO_DESTINOS_SCHEMA);
    inativarRegistrosPorCampoRegressao(ABA_PRODUCAO_SAIDAS_LOTES, 'producao_id', id, 'id', PRODUCAO_SAIDAS_LOTES_SCHEMA);
    inativarRegistrosPorCampoRegressao(ABA_PRODUCAO_CONSUMO, 'producao_id', id, 'id', PRODUCAO_CONSUMO_SCHEMA);
    inativarRegistrosPorCampoRegressao(ABA_PRODUCAO_MATERIAIS, 'producao_id', id, 'id', PRODUCAO_MATERIAIS_SCHEMA);
    inativarRegistrosPorCampoRegressao(ABA_PRODUCAO_MATERIAIS_PREVISTOS, 'producao_id', id, 'id', PRODUCAO_MATERIAIS_PREVISTOS_SCHEMA);
    inativarRegistrosPorCampoRegressao(ABA_PRODUCAO_NECESSIDADES_ENTRADA, 'producao_id', id, 'id', PRODUCAO_NECESSIDADES_ENTRADA_SCHEMA);
    inativarRegistrosPorCampoRegressao(ABA_PRODUCAO_RESERVAS_ENTRADA, 'producao_id', id, 'id', PRODUCAO_RESERVAS_ENTRADA_SCHEMA);
  });
}

function criarItemEstoqueFixtureRegressao(contexto, overrides) {
  const cfg = overrides || {};
  const token = contexto?.token || '__REGTEST';
  const id = String(cfg.ID || gerarId('EST')).trim();
  const tipo = String(cfg.tipo || 'INSUMO').trim().toUpperCase();
  const quantidade = parseNumeroBR(cfg.quantidade);
  const custoUnit = parseNumeroBR(
    Object.prototype.hasOwnProperty.call(cfg, 'custo_unitario')
      ? cfg.custo_unitario
      : (cfg.valor_unit || 0)
  );

  const novo = {
    ID: id,
    tipo,
    item: String(cfg.item || `${token} ITEM ${id}`).trim(),
    unidade: String(cfg.unidade || 'UN').trim() || 'UN',
    valor_unit: custoUnit,
    custo_unitario: custoUnit,
    preco_venda: Object.prototype.hasOwnProperty.call(cfg, 'preco_venda') ? cfg.preco_venda : '',
    ativo: true,
    criado_em: new Date(),
    quantidade,
    comprimento_cm: cfg.comprimento_cm || '',
    largura_cm: cfg.largura_cm || '',
    espessura_cm: cfg.espessura_cm || '',
    categoria: String(cfg.categoria || '').trim(),
    fornecedor: '',
    pago_por: '',
    potencia: '',
    voltagem: '',
    comprado_em: '',
    data_pagamento: '',
    forma_pagamento: '',
    parcelas: 1,
    vida_util_mes: '',
    observacao: String(cfg.observacao || token).trim(),
    origem_tipo: String(cfg.origem_tipo || '').trim(),
    origem_id: String(cfg.origem_id || '').trim(),
    op_id: String(cfg.op_id || '').trim()
  };

  insert(ABA_ESTOQUE, novo, ESTOQUE_SCHEMA);
  registrarCleanupItemEstoqueRegressao(contexto, id);
  return novo;
}

function criarProdutoModeloFixtureRegressao(contexto, config) {
  const cfg = config || {};
  const token = contexto?.token || '__REGTEST';

  const produtoId = String(cfg.produto_id || gerarId('PRD')).trim();
  const produto = {
    produto_id: produtoId,
    nome_produto: String(cfg.nome_produto || `${token} PRODUTO ${produtoId}`).trim(),
    unidade_produto: String(cfg.unidade_produto || 'UN').trim() || 'UN',
    preco_venda: parseNumeroBR(cfg.preco_venda),
    ativo: true,
    criado_em: new Date()
  };
  insert(ABA_PRODUTOS, produto, PRODUTOS_SCHEMA);
  registrarCleanupRegressao(contexto, () => {
    inativarRegistroRegressao(ABA_PRODUTOS, 'produto_id', produtoId, PRODUTOS_SCHEMA);
  });

  const receitaId = String(cfg.receita_id || gerarId('REC')).trim();
  const receita = {
    receita_id: receitaId,
    produto_id: produtoId,
    nome_receita: String(cfg.nome_receita || 'Modelo principal').trim() || 'Modelo principal',
    descricao: String(cfg.descricao || token).trim(),
    parent_receita_id: '',
    ativo: true,
    criado_em: new Date()
  };
  insert(ABA_PRODUTOS_RECEITAS, receita, PRODUTOS_RECEITAS_SCHEMA);
  registrarCleanupRegressao(contexto, () => {
    try {
      deletarReceitaProduto(receitaId);
    } catch (error) {
      inativarRegistroRegressao(ABA_PRODUTOS_RECEITAS, 'receita_id', receitaId, PRODUTOS_RECEITAS_SCHEMA);
      inativarRegistrosPorCampoRegressao(ABA_PRODUTOS_RECEITAS_ENTRADAS, 'receita_id', receitaId, 'id', PRODUTOS_RECEITAS_ENTRADAS_SCHEMA);
      inativarRegistrosPorCampoRegressao(ABA_PRODUTOS_RECEITAS_SAIDAS, 'receita_id', receitaId, 'id', PRODUTOS_RECEITAS_SAIDAS_SCHEMA);
    }
  });

  const entradas = (Array.isArray(cfg.entradas) ? cfg.entradas : []).map((entrada, idx) => {
    const tipo = String(entrada?.tipo_item || '').trim().toUpperCase();
    const novo = {
      id: String(entrada?.id || gerarId('REN')).trim(),
      receita_id: receitaId,
      tipo_item: tipo,
      nome_item: String(entrada?.nome_item || `${token} ENTRADA ${idx + 1}`).trim(),
      estoque_ref_id: '',
      produto_ref_id: '',
      receita_ref_id: '',
      categoria: String(entrada?.categoria || '').trim(),
      unidade: String(entrada?.unidade || '').trim() || (tipo === 'MADEIRA' ? 'M3' : 'UN'),
      qtd_pecas: parseNumeroBR(entrada?.qtd_pecas),
      comprimento_cm: parseNumeroBR(entrada?.comprimento_cm),
      largura_cm: parseNumeroBR(entrada?.largura_cm),
      espessura_cm: parseNumeroBR(entrada?.espessura_cm),
      custo_manual: parseNumeroBR(entrada?.custo_manual),
      observacao: '',
      ativo: true
    };
    insert(ABA_PRODUTOS_RECEITAS_ENTRADAS, novo, PRODUTOS_RECEITAS_ENTRADAS_SCHEMA);
    return novo;
  });

  const saidasBase = Array.isArray(cfg.saidas) && cfg.saidas.length > 0
    ? cfg.saidas
    : [{
      nome_saida: produto.nome_produto,
      produto_ref_id: produtoId,
      tipo_item: 'PRODUTO',
      categoria: 'PECA',
      unidade: produto.unidade_produto || 'UN',
      quantidade: 1
    }];

  const saidas = saidasBase.map((saida, idx) => {
    const novo = {
      id: String(saida?.id || gerarId('RSA')).trim(),
      receita_id: receitaId,
      nome_saida: String(saida?.nome_saida || `${token} SAIDA ${idx + 1}`).trim(),
      produto_ref_id: String(saida?.produto_ref_id || '').trim(),
      tipo_item: String(saida?.tipo_item || 'PRODUTO').trim().toUpperCase() || 'PRODUTO',
      categoria: String(saida?.categoria || '').trim(),
      unidade: String(saida?.unidade || produto.unidade_produto || 'UN').trim() || 'UN',
      quantidade: parseNumeroBR(saida?.quantidade || 1),
      ativo: true
    };
    insert(ABA_PRODUTOS_RECEITAS_SAIDAS, novo, PRODUTOS_RECEITAS_SAIDAS_SCHEMA);
    return novo;
  });

  return {
    produto,
    receita,
    entradas,
    saidas
  };
}

function criarProducaoFixtureRegressao(contexto, fixtureModelo, overrides) {
  const cfg = overrides || {};
  const produtoId = String(cfg.produto_id || fixtureModelo?.produto?.produto_id || '').trim();
  const receitaId = String(cfg.receita_id || fixtureModelo?.receita?.receita_id || '').trim();
  const qtdPlanejada = parseNumeroBR(cfg.qtd_planejada || 1);

  const op = criarProducao({
    produto_id: produtoId,
    receita_id: receitaId,
    nome_ordem: String(cfg.nome_ordem || `${contexto?.token || '__REGTEST'} OP`).trim(),
    qtd_planejada: qtdPlanejada,
    status: String(cfg.status || 'Em planejamento').trim() || 'Em planejamento',
    data_inicio: cfg.data_inicio || '',
    data_prevista_termino: cfg.data_prevista_termino || '',
    observacao: String(cfg.observacao || '').trim()
  });

  registrarCleanupProducaoRegressao(contexto, op?.producao_id);
  return op;
}

function construirConsumoPayloadCompletoRegressao(vinculos) {
  const itens = [];
  (Array.isArray(vinculos) ? vinculos : []).forEach(vinculo => {
    const reservas = Array.isArray(vinculo?.reservas) ? vinculo.reservas : [];
    reservas.forEach(reserva => {
      const quantidade = parseNumeroBR(reserva?.quantidade_restante);
      if (!reserva?.estoque_id || quantidade <= 0) return;
      itens.push({
        necessidade_id: String(vinculo?.id || '').trim(),
        reserva_id: String(reserva?.id || '').trim(),
        estoque_id: String(reserva?.estoque_id || '').trim(),
        quantidade
      });
    });
  });
  return itens;
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

  const chavePorta = normalizarTextoProducao('Porta');
  const porta = agrupadas.find(item => normalizarTextoProducao(item?.nome_saida) === chavePorta);
  assertRegressao(
    !!porta,
    `Grupo Porta nao encontrado. Grupos encontrados: ${JSON.stringify(
      (agrupadas || []).map(g => g?.nome_saida || '')
    )}`
  );
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

function regressionProducaoReservaQuantidade() {
  const contexto = criarContextoFixtureRegressao('PROD_RESERVA_QTD');
  try {
    const fixture = criarProdutoModeloFixtureRegressao(contexto, {
      nome_produto: `${contexto.token} PROD RESERVA`,
      entradas: [{
        tipo_item: 'INSUMO',
        nome_item: `${contexto.token} LIXA`,
        categoria: 'LIXA',
        unidade: 'UN',
        qtd_pecas: 2,
        custo_manual: 3
      }]
    });

    const estoqueA = criarItemEstoqueFixtureRegressao(contexto, {
      item: `${contexto.token} LIXA A`,
      tipo: 'INSUMO',
      categoria: 'LIXA',
      unidade: 'UN',
      quantidade: 5,
      custo_unitario: 3
    });
    const estoqueB = criarItemEstoqueFixtureRegressao(contexto, {
      item: `${contexto.token} LIXA B`,
      tipo: 'INSUMO',
      categoria: 'LIXA',
      unidade: 'UN',
      quantidade: 4,
      custo_unitario: 4
    });

    const op = criarProducaoFixtureRegressao(contexto, fixture, { qtd_planejada: 3 });
    const vinculosIniciais = listarVinculosMateriaisProducao(op.producao_id);
    assertRegressao(vinculosIniciais.length === 1, 'A OP deveria gerar 1 necessidade agregada.');

    const necessidade = vinculosIniciais[0];
    assertAproxRegressao(necessidade.quantidade_prevista, 6, 0.000001, 'Quantidade prevista da necessidade deveria ser 6.');
    assertRegressao(necessidade.status === 'Nao reservado', 'Necessidade inicial deveria estar sem reserva.');

    assertThrowsRegressao(
      () => salvarReservaEntradaProducao(op.producao_id, necessidade.id, {
        estoque_id: estoqueA.ID,
        quantidade_reservada: 6
      }),
      'Quantidade reservada maior que o saldo disponivel',
      'Reserva acima do disponivel deveria ser bloqueada.'
    );

    const respParcial = salvarReservaEntradaProducao(op.producao_id, necessidade.id, {
      estoque_id: estoqueA.ID,
      quantidade_reservada: 5
    });
    const vinculoParcial = (respParcial.vinculos || []).find(v => String(v.id || '') === String(necessidade.id || ''));
    assertRegressao(!!vinculoParcial, 'Necessidade parcial nao encontrada apos reserva.');
    assertRegressao(vinculoParcial.status === 'Parcial', 'Status deveria ser Parcial apos primeira reserva.');
    assertAproxRegressao(vinculoParcial.quantidade_reservada, 5, 0.000001, 'Quantidade reservada parcial incorreta.');
    assertAproxRegressao(vinculoParcial.quantidade_pendente, 1, 0.000001, 'Quantidade pendente deveria ser 1.');

    const estoqueAPosParcial = obterItemEstoque(estoqueA.ID);
    const estoqueBPosParcial = obterItemEstoque(estoqueB.ID);
    assertAproxRegressao(estoqueAPosParcial.quantidade_disponivel, 0, 0.000001, 'Estoque A deveria ficar sem saldo disponivel.');
    assertAproxRegressao(estoqueBPosParcial.quantidade_disponivel, 4, 0.000001, 'Estoque B ainda deveria estar livre.');

    const respCompleta = salvarReservaEntradaProducao(op.producao_id, necessidade.id, {
      estoque_id: estoqueB.ID,
      quantidade_reservada: 1
    });
    const vinculoCompleto = (respCompleta.vinculos || []).find(v => String(v.id || '') === String(necessidade.id || ''));
    assertRegressao(vinculoCompleto.status === 'Reservado', 'Status deveria ser Reservado apos completar a alocacao.');
    assertAproxRegressao(vinculoCompleto.quantidade_reservada, 6, 0.000001, 'Quantidade total reservada incorreta.');
    assertAproxRegressao(vinculoCompleto.quantidade_pendente, 0, 0.000001, 'Nao deveria restar saldo pendente.');
    assertRegressao(Array.isArray(vinculoCompleto.reservas) && vinculoCompleto.reservas.length === 2, 'A necessidade deveria possuir 2 reservas.');

    const estoqueBPosCompleto = obterItemEstoque(estoqueB.ID);
    assertAproxRegressao(estoqueBPosCompleto.quantidade_disponivel, 3, 0.000001, 'Estoque B deveria manter 3 unidades disponiveis.');

    return {
      producao_id: op.producao_id,
      necessidade_id: necessidade.id,
      reservas: vinculoCompleto.reservas.length
    };
  } finally {
    executarCleanupRegressao(contexto);
  }
}

function regressionProducaoMadeiraReservaExclusiva() {
  const contexto = criarContextoFixtureRegressao('PROD_MADEIRA');
  try {
    const compMin = 200;
    const largMin = 20;
    const espMin = 4;
    const volumeMin = (compMin * largMin * espMin) / 1000000;

    const fixture = criarProdutoModeloFixtureRegressao(contexto, {
      nome_produto: `${contexto.token} PROD MADEIRA`,
      entradas: [{
        tipo_item: 'MADEIRA',
        nome_item: `${contexto.token} TABUA`,
        categoria: 'TABUA',
        unidade: 'M3',
        qtd_pecas: volumeMin,
        comprimento_cm: compMin,
        largura_cm: largMin,
        espessura_cm: espMin,
        custo_manual: 2500
      }]
    });

    const madeiraIncompativel = criarItemEstoqueFixtureRegressao(contexto, {
      item: `${contexto.token} MADEIRA MENOR`,
      tipo: 'MADEIRA',
      categoria: 'TABUA',
      unidade: 'M3',
      quantidade: (190 * 20 * 4) / 1000000,
      comprimento_cm: 190,
      largura_cm: 20,
      espessura_cm: 4,
      custo_unitario: 1800
    });
    const madeiraA = criarItemEstoqueFixtureRegressao(contexto, {
      item: `${contexto.token} MADEIRA A`,
      tipo: 'MADEIRA',
      categoria: 'TABUA',
      unidade: 'M3',
      quantidade: (210 * 20 * 4) / 1000000,
      comprimento_cm: 210,
      largura_cm: 20,
      espessura_cm: 4,
      custo_unitario: 2000
    });
    const madeiraB = criarItemEstoqueFixtureRegressao(contexto, {
      item: `${contexto.token} MADEIRA B`,
      tipo: 'MADEIRA',
      categoria: 'TABUA',
      unidade: 'M3',
      quantidade: (220 * 20 * 4) / 1000000,
      comprimento_cm: 220,
      largura_cm: 20,
      espessura_cm: 4,
      custo_unitario: 2100
    });

    const op = criarProducaoFixtureRegressao(contexto, fixture, { qtd_planejada: 2 });
    const vinculos = listarVinculosMateriaisProducao(op.producao_id);
    assertRegressao(vinculos.length === 2, 'Madeira deveria explodir 2 necessidades para OP de 2 unidades.');
    assertRegressao(vinculos.every(v => v.modo_atendimento === 'PECA_UNICA'), 'Todas as necessidades de madeira deveriam ser PECA_UNICA.');
    assertRegressao(String(vinculos[0].serie_item || '') === '1', 'Primeira serie de madeira deveria ser 1.');
    assertRegressao(String(vinculos[1].serie_item || '') === '2', 'Segunda serie de madeira deveria ser 2.');

    assertThrowsRegressao(
      () => salvarReservaEntradaProducao(op.producao_id, vinculos[0].id, {
        estoque_id: madeiraIncompativel.ID
      }),
      'nao compativel',
      'Madeira menor que a dimensao minima nao deveria ser aceita.'
    );

    const respA = salvarReservaEntradaProducao(op.producao_id, vinculos[0].id, {
      estoque_id: madeiraA.ID
    });
    const vinculoA = (respA.vinculos || []).find(v => String(v.id || '') === String(vinculos[0].id || ''));
    assertRegressao(vinculoA.status === 'Reservado', 'Primeira madeira deveria ficar reservada.');
    assertRegressao(vinculoA.reservas.length === 1, 'Necessidade de madeira aceita somente uma reserva.');

    const madeiraAPosReserva = obterItemEstoque(madeiraA.ID);
    assertAproxRegressao(madeiraAPosReserva.quantidade_disponivel, 0, 0.000001, 'Peca de madeira reservada deveria ficar indisponivel.');

    assertThrowsRegressao(
      () => salvarReservaEntradaProducao(op.producao_id, vinculos[1].id, {
        estoque_id: madeiraA.ID
      }),
      'nao esta disponivel',
      'A mesma peca nao deveria ser reservada para outra necessidade.'
    );

    const respB = salvarReservaEntradaProducao(op.producao_id, vinculos[1].id, {
      estoque_id: madeiraB.ID
    });
    const vinculoB = (respB.vinculos || []).find(v => String(v.id || '') === String(vinculos[1].id || ''));
    assertRegressao(vinculoB.status === 'Reservado', 'Segunda madeira deveria ficar reservada.');

    const madeiraBPosReserva = obterItemEstoque(madeiraB.ID);
    const madeiraMenor = obterItemEstoque(madeiraIncompativel.ID);
    assertAproxRegressao(madeiraBPosReserva.quantidade_disponivel, 0, 0.000001, 'Segunda madeira deveria ficar indisponivel.');
    assertRegressao(parseNumeroBR(madeiraMenor.quantidade_disponivel) > 0, 'Madeira incompatível deveria permanecer livre.');

    return {
      producao_id: op.producao_id,
      necessidades: vinculos.length
    };
  } finally {
    executarCleanupRegressao(contexto);
  }
}

function regressionProducaoConsumoSoftDelete() {
  const contexto = criarContextoFixtureRegressao('PROD_CONSUMO');
  try {
    const fixture = criarProdutoModeloFixtureRegressao(contexto, {
      nome_produto: `${contexto.token} PROD CONSUMO`,
      entradas: [{
        tipo_item: 'INSUMO',
        nome_item: `${contexto.token} COLA`,
        categoria: 'COLA',
        unidade: 'UN',
        qtd_pecas: 2,
        custo_manual: 8
      }],
      saidas: [{
        nome_saida: `${contexto.token} SAIDA FINAL`,
        produto_ref_id: '',
        tipo_item: 'PRODUTO',
        categoria: 'PECA',
        unidade: 'UN',
        quantidade: 1
      }]
    });

    const estoqueEntrada = criarItemEstoqueFixtureRegressao(contexto, {
      item: `${contexto.token} COLA ESTOQUE`,
      tipo: 'INSUMO',
      categoria: 'COLA',
      unidade: 'UN',
      quantidade: 4,
      custo_unitario: 8
    });

    const op = criarProducaoFixtureRegressao(contexto, fixture, { qtd_planejada: 2 });
    const vinculosIniciais = listarVinculosMateriaisProducao(op.producao_id);
    const necessidade = vinculosIniciais[0];
    const respReserva = salvarReservaEntradaProducao(op.producao_id, necessidade.id, {
      estoque_id: estoqueEntrada.ID,
      quantidade_reservada: 4
    });
    const vinculoReservado = (respReserva.vinculos || []).find(v => String(v.id || '') === String(necessidade.id || ''));
    const itensBaixa = construirConsumoPayloadCompletoRegressao([vinculoReservado]);
    assertRegressao(itensBaixa.length === 1, 'Baixa deveria ser montada com uma unica reserva.');

    const respConsumo = consumirEstoque(op.producao_id, itensBaixa, {
      qtd_produzida: 2,
      bypass_entradas: false
    });

    assertRegressao(respConsumo?.producaoAtualizada?.estoque_atualizado === true, 'A OP deveria ser concluida apos produzir todo o planejado.');
    assertAproxRegressao(respConsumo?.producaoAtualizada?.qtd_restante, 0, 0.000001, 'Qtd restante deveria ser zero apos concluir a OP.');
    assertRegressao(Array.isArray(respConsumo?.saidasLotes) && respConsumo.saidasLotes.length >= 1, 'A baixa deveria gerar pelo menos um lote de saida.');

    (respConsumo.saidasLotes || []).forEach(lote => {
      registrarCleanupRegressao(contexto, () => {
        inativarRegistroRegressao(ABA_PRODUCAO_SAIDAS_LOTES, 'id', lote.id, PRODUCAO_SAIDAS_LOTES_SCHEMA);
      });
      registrarCleanupItemEstoqueRegressao(contexto, lote.estoque_id);
    });

    const linhaEntrada = obterRegistroPorCampoRegressao(ABA_ESTOQUE, 'ID', estoqueEntrada.ID);
    assertRegressao(!!linhaEntrada, 'Linha de estoque consumida nao encontrada.');
    assertAproxRegressao(parseNumeroBR(linhaEntrada.quantidade), 0, 0.000001, 'Estoque consumido deveria zerar.');
    assertRegressao(String(linhaEntrada.ativo).toLowerCase() === 'false', 'Estoque consumido deveria sofrer soft delete.');

    const vinculosFinais = listarVinculosMateriaisProducao(op.producao_id);
    const vinculoFinal = vinculosFinais.find(v => String(v.id || '') === String(necessidade.id || ''));
    assertRegressao(vinculoFinal.status === 'Baixado', 'Necessidade deveria ficar baixada apos consumo.');
    assertAproxRegressao(vinculoFinal.quantidade_baixada, 4, 0.000001, 'Quantidade baixada incorreta apos consumo.');

    return {
      producao_id: op.producao_id,
      saidas_lotes: respConsumo.saidasLotes.length
    };
  } finally {
    executarCleanupRegressao(contexto);
  }
}

function regressionProducaoAtualizacaoRegeraNecessidades() {
  const contexto = criarContextoFixtureRegressao('PROD_REPLANEJAR');
  try {
    const fixture = criarProdutoModeloFixtureRegressao(contexto, {
      nome_produto: `${contexto.token} PROD REPLANEJAR`,
      entradas: [{
        tipo_item: 'INSUMO',
        nome_item: `${contexto.token} FELTRO`,
        categoria: 'FELTRO',
        unidade: 'UN',
        qtd_pecas: 2,
        custo_manual: 5
      }]
    });

    const estoque = criarItemEstoqueFixtureRegressao(contexto, {
      item: `${contexto.token} FELTRO ESTOQUE`,
      tipo: 'INSUMO',
      categoria: 'FELTRO',
      unidade: 'UN',
      quantidade: 5,
      custo_unitario: 5
    });

    const op = criarProducaoFixtureRegressao(contexto, fixture, { qtd_planejada: 2 });
    const vinculoInicial = listarVinculosMateriaisProducao(op.producao_id)[0];
    salvarReservaEntradaProducao(op.producao_id, vinculoInicial.id, {
      estoque_id: estoque.ID,
      quantidade_reservada: 3
    });

    const estoquePosReserva = obterItemEstoque(estoque.ID);
    assertAproxRegressao(estoquePosReserva.quantidade_disponivel, 2, 0.000001, 'Reserva inicial deveria reduzir saldo disponivel.');

    const ok = atualizarProducao(op.producao_id, { qtd_planejada: 3 });
    assertRegressao(ok === true, 'Atualizacao da OP deveria retornar true.');

    const vinculosRegerados = listarVinculosMateriaisProducao(op.producao_id);
    assertRegressao(vinculosRegerados.length === 1, 'A OP deveria continuar com uma necessidade apos replanejamento.');
    assertAproxRegressao(vinculosRegerados[0].quantidade_prevista, 6, 0.000001, 'Quantidade prevista deveria ser recalculada para 6.');
    assertRegressao(vinculosRegerados[0].status === 'Nao reservado', 'Reservas antigas deveriam ser liberadas ao editar a OP.');
    assertRegressao((vinculosRegerados[0].reservas || []).length === 0, 'Nao deveria restar reserva vinculada apos replanejamento.');

    const estoquePosAtualizacao = obterItemEstoque(estoque.ID);
    assertAproxRegressao(estoquePosAtualizacao.quantidade_disponivel, 5, 0.000001, 'Saldo disponivel do estoque deveria ser restaurado apos limpar reservas.');

    return {
      producao_id: op.producao_id,
      quantidade_prevista: vinculosRegerados[0].quantidade_prevista
    };
  } finally {
    executarCleanupRegressao(contexto);
  }
}

function regressionVendaSaldoReservado() {
  const contexto = criarContextoFixtureRegressao('VENDA_RESERVA');
  try {
    const fixture = criarProdutoModeloFixtureRegressao(contexto, {
      nome_produto: `${contexto.token} PROD VENDA`,
      entradas: [{
        tipo_item: 'PRODUTO',
        nome_item: `${contexto.token} ITEM VENDAVEL`,
        categoria: 'FINALIZADO',
        unidade: 'UN',
        qtd_pecas: 4,
        custo_manual: 50
      }]
    });

    const estoqueVendavel = criarItemEstoqueFixtureRegressao(contexto, {
      item: `${contexto.token} ITEM VENDAVEL`,
      tipo: 'PRODUTO',
      categoria: 'FINALIZADO',
      unidade: 'UN',
      quantidade: 5,
      custo_unitario: 50,
      preco_venda: 120
    });

    const op = criarProducaoFixtureRegressao(contexto, fixture, { qtd_planejada: 1 });
    const vinculo = listarVinculosMateriaisProducao(op.producao_id)[0];
    salvarReservaEntradaProducao(op.producao_id, vinculo.id, {
      estoque_id: estoqueVendavel.ID,
      quantidade_reservada: 4
    });

    const vendaveis = listarItensEstoqueVendaveis();
    const itemLista = vendaveis.find(i => String(i.ID || '') === String(estoqueVendavel.ID || ''));
    assertRegressao(!!itemLista, 'Item vendavel deveria continuar aparecendo na lista.');
    assertAproxRegressao(itemLista.quantidade, 1, 0.000001, 'Lista de venda deveria expor apenas o saldo disponivel.');

    const itemEstoque = obterItemEstoqueVendavelPorId(estoqueVendavel.ID);
    assertAproxRegressao(itemEstoque.quantidade_disponivel, 1, 0.000001, 'Saldo disponivel do item vendavel deveria ser 1.');

    assertThrowsRegressao(
      () => normalizarPayloadVenda({
        estoque_id: estoqueVendavel.ID,
        quantidade: 2,
        valor_total_venda: 240,
        data_venda: '2026-04-17',
        forma_pagamento: 'PIX'
      }),
      'Quantidade de venda maior que o saldo disponivel',
      'Venda nao deveria ignorar quantidade reservada pela producao.'
    );

    const payloadValido = normalizarPayloadVenda({
      estoque_id: estoqueVendavel.ID,
      quantidade: 1,
      valor_total_venda: 120,
      data_venda: '2026-04-17',
      forma_pagamento: 'PIX'
    });
    assertAproxRegressao(payloadValido.quantidade, 1, 0.000001, 'Venda valida deveria aceitar o saldo disponivel.');

    return {
      estoque_id: estoqueVendavel.ID,
      saldo_disponivel: itemEstoque.quantidade_disponivel
    };
  } finally {
    executarCleanupRegressao(contexto);
  }
}

function regressionDocumentoCompraCalculos() {
  const itens = ratearCustosDocumentoCompra_([
    { ID: 'A', item_nome: 'Item A', valor_total: 100 },
    { ID: 'B', item_nome: 'Item B', valor_total: 300 }
  ], 40, 20);

  assertRegressao(itens.length === 2, 'Rateio deveria preservar os dois itens.');
  assertAproxRegressao(itens[0].frete_rateado, 10, 0.001, 'Frete proporcional do item A incorreto.');
  assertAproxRegressao(itens[0].desconto_rateado, 5, 0.001, 'Desconto proporcional do item A incorreto.');
  assertAproxRegressao(itens[0].custo_total, 105, 0.001, 'Custo final do item A incorreto.');
  assertAproxRegressao(itens[1].custo_total, 315, 0.001, 'Custo final do item B incorreto.');
  assertAproxRegressao(
    itens.reduce((acc, item) => acc + item.custo_total, 0),
    420,
    0.001,
    'A soma dos custos deveria reconciliar com subtotal + frete - desconto.'
  );

  const quantidadeMadeira = calcularQuantidadeEstoqueDocumentoCompra_({
    tipo: 'MADEIRA',
    quantidade: 2,
    comprimento_cm: 200,
    largura_cm: 10,
    espessura_cm: 2
  });
  assertAproxRegressao(quantidadeMadeira, 0.008, 0.000001, 'Volume de duas pecas de madeira incorreto.');
  assertRegressao(
    normalizarDestinoItemDocumentoCompra_('valor desconhecido') === 'ESTOQUE',
    'Destino desconhecido deveria exigir revisao partindo de ESTOQUE.'
  );

  return {
    ok: true,
    custo_total: itens.reduce((acc, item) => acc + item.custo_total, 0),
    quantidade_madeira_m3: quantidadeMadeira
  };
}

function regressionComprovantesFinanceiros() {
  assertRegressao(
    new Set(INBOX_DESPESAS_SCHEMA).size === INBOX_DESPESAS_SCHEMA.length,
    'Schema de INBOX_DESPESAS nao deve conter cabecalhos duplicados.'
  );
  ['classificacao', 'referencia_transacao', 'conciliado_em'].forEach(cabecalho => {
    assertRegressao(INBOX_DESPESAS_SCHEMA.includes(cabecalho), `Cabecalho ausente na Inbox: ${cabecalho}.`);
  });
  assertRegressao(
    new Set(PAGAMENTOS_SCHEMA).size === PAGAMENTOS_SCHEMA.length,
    'Schema de PAGAMENTOS nao deve conter cabecalhos duplicados.'
  );
  assertRegressao(PAGAMENTOS_SCHEMA.includes('comprovante_id'), 'Cabecalho comprovante_id ausente em PAGAMENTOS.');

  const comprovante = {
    descricao: 'Pagamento Mercado Livre Ferragens Silva',
    fornecedor: 'Ferragens Silva',
    pago_por: 'Bruno',
    data_pagamento: '2026-07-18'
  };
  const destinoExato = pontuarDestinoComprovante_(comprovante, 250, {
    origem_tipo: 'COMPRA', origem_id: 'COM_1', parcela_id: 'PAR_1',
    fornecedor: 'Ferragens Silva', origem_rotulo: 'Ferragens do Mercado Livre',
    observacao: '', pago_por: 'Bruno', data_prevista: '2026-07-18',
    valor_pendente: 250, label: 'Compra exata'
  });
  const destinoParcial = pontuarDestinoComprovante_(comprovante, 100, {
    origem_tipo: 'COMPRA', origem_id: 'COM_2', parcela_id: 'PAR_2',
    fornecedor: '', origem_rotulo: 'Outro fornecedor', observacao: '',
    pago_por: 'Socio', data_prevista: '2026-10-18',
    valor_pendente: 250, label: 'Compra parcial'
  });
  assertRegressao(destinoExato.score >= 85, 'Coincidencia exata deveria receber score alto.');
  assertRegressao(destinoExato.score > destinoParcial.score, 'Coincidencia exata deveria superar sugestao parcial fraca.');
  assertRegressao(
    destinoParcial.motivos.includes('possivel pagamento parcial'),
    'Sugestao deve reconhecer pagamentos parciais.'
  );
  assertRegressao(
    similaridadeTextoComprovante_('Depósito São José', 'Sao Jose materiais') > 0,
    'Matching textual deve ignorar acentos.'
  );
  const remapeadas = mapearParcelasAlvoRegeradasFinanceiro(
    [
      { ID: 'PGT_A', parcela_alvo_id: 'PAR_ANTIGA_2' },
      { ID: 'PGT_B', parcela_alvo_id: 'PAR_ANTIGA_2' }
    ],
    [
      { ID: 'PAR_ANTIGA_1', origem_tipo: 'COMPRA', origem_id: 'COM_1', parcela_numero: 1 },
      { ID: 'PAR_ANTIGA_2', origem_tipo: 'COMPRA', origem_id: 'COM_1', parcela_numero: 2 }
    ],
    [
      { ID: 'PAR_NOVA_1', parcela_numero: 1 },
      { ID: 'PAR_NOVA_2', parcela_numero: 2 }
    ],
    'COMPRA',
    'COM_1'
  );
  assertRegressao(
    remapeadas.every(item => item.parcela_alvo_atual_id === 'PAR_NOVA_2'),
    'Multiplos pagamentos parciais devem permanecer na mesma parcela apos recalculo.'
  );

  const historico = montarHistoricoPagamentosOrigemFinanceiro_(
    {
      tipo: 'COMPRA',
      id: 'COM_HIST',
      natureza: 'PAGAMENTO',
      total_previsto: 300,
      item: { item: 'Madeira teste', comprado_em: '2026-07-10', criado_em: '2026-07-10 08:00' }
    },
    [
      {
        ID: 'PGT_ANTIGO', origem_tipo: 'COMPRA', origem_id: 'COM_HIST', ativo: true,
        parcela_alvo_id: 'PAR_1', data_pagamento: '2026-07-11', valor_pago: 100,
        forma_pagamento: 'PIX', observacao: 'Primeira parte', criado_em: '2026-07-11 09:00',
        client_request_id: 'NAO_EXPOR'
      },
      {
        ID: 'PGT_RECENTE', origem_tipo: 'COMPRA', origem_id: 'COM_HIST', ativo: true,
        parcela_alvo_id: 'PAR_2', data_pagamento: '2026-07-15', valor_pago: 50,
        forma_pagamento: 'DINHEIRO', observacao: '', criado_em: '2026-07-15 10:00'
      },
      {
        ID: 'PGT_INATIVO', origem_tipo: 'COMPRA', origem_id: 'COM_HIST', ativo: false,
        parcela_alvo_id: 'PAR_2', data_pagamento: '2026-07-16', valor_pago: 80,
        forma_pagamento: 'PIX'
      }
    ],
    [
      {
        ID: 'PAR_1', origem_tipo: 'COMPRA', origem_id: 'COM_HIST', ativo: true,
        parcela_numero: 1, parcelas_total: 2, data_prevista: '2026-07-11',
        valor_previsto: 150, valor_pago: 100, status: 'PARCIAL'
      },
      {
        ID: 'PAR_2', origem_tipo: 'COMPRA', origem_id: 'COM_HIST', ativo: true,
        parcela_numero: 2, parcelas_total: 2, data_prevista: '2026-08-11',
        valor_previsto: 150, valor_pago: 50, status: 'PARCIAL'
      }
    ]
  );
  assertRegressao(historico.pagamentos.length === 2, 'Historico deve excluir pagamentos inativos.');
  assertRegressao(
    historico.pagamentos[0].ID === 'PGT_RECENTE',
    'Historico deve ordenar pagamentos do mais recente para o mais antigo.'
  );
  assertRegressao(
    historico.pagamentos[0].parcela_alvo?.parcela_numero === 2,
    'Historico deve identificar a parcela alvo do pagamento.'
  );
  assertAproxRegressao(historico.resumo.total_pago, 150, 0.001, 'Total pago do historico incorreto.');
  assertAproxRegressao(historico.resumo.total_pendente, 150, 0.001, 'Saldo pendente do historico incorreto.');
  assertRegressao(
    !Object.prototype.hasOwnProperty.call(historico.pagamentos[1], 'client_request_id'),
    'Historico nao deve expor identificadores internos de idempotencia.'
  );

  const historicoVenda = montarHistoricoPagamentosOrigemFinanceiro_(
    {
      tipo: 'VENDA', id: 'VND_HIST', natureza: 'RECEBIMENTO', total_previsto: 120,
      item: { item: 'Mesa teste', data_venda: '2026-07-12', criado_em: '2026-07-12 08:00' }
    },
    [{
      ID: 'RCB_HIST', origem_tipo: 'VENDA', origem_id: 'VND_HIST', ativo: true,
      data_pagamento: '2026-07-13', valor_pago: 40, forma_pagamento: 'PIX'
    }],
    []
  );
  assertRegressao(
    historicoVenda.origem.tipo === 'VENDA' && historicoVenda.origem.natureza === 'RECEBIMENTO',
    'Historico de venda deve preservar tipo e natureza de recebimento.'
  );
  assertAproxRegressao(historicoVenda.resumo.total_pago, 40, 0.001, 'Total recebido incorreto.');

  const historicoDespesa = montarHistoricoPagamentosOrigemFinanceiro_(
    {
      tipo: 'DESPESA_GERAL', id: 'DES_HIST', natureza: 'PAGAMENTO', total_previsto: 90,
      item: { descricao: 'Energia', data_competencia: '2026-07-01', criado_em: '2026-07-01 08:00' }
    },
    [],
    []
  );
  assertRegressao(
    historicoDespesa.origem.tipo === 'DESPESA_GERAL' && historicoDespesa.pagamentos.length === 0,
    'Historico de despesa deve usar o tipo canonico e aceitar lista sem pagamentos.'
  );

  let rejeitouOrigemDivergente = false;
  try {
    validarPagamentoHistoricoOrigemFinanceiro_(
      { ID: 'PGT_VALIDAR', origem_tipo: 'COMPRA', origem_id: 'COM_1', ativo: true },
      'PGT_VALIDAR',
      'VENDA',
      'VND_1'
    );
  } catch (error) {
    rejeitouOrigemDivergente = /nao pertence/i.test(String(error?.message || ''));
  }
  assertRegressao(
    rejeitouOrigemDivergente,
    'Remocao pelo historico deve rejeitar origem divergente antes de alterar dados.'
  );

  return {
    ok: true,
    score_exato: destinoExato.score,
    score_parcial: destinoParcial.score
  };
}

function regressionComprovantesRecebimentos() {
  assertRegressao(
    new Set(INBOX_RECEBIMENTOS_SCHEMA).size === INBOX_RECEBIMENTOS_SCHEMA.length,
    'Schema de INBOX_RECEBIMENTOS nao deve conter cabecalhos duplicados.'
  );
  ['arquivo_hash', 'referencia_transacao', 'pagador_nome', 'data_recebimento', 'conciliado_em']
    .forEach(cabecalho => {
      assertRegressao(
        INBOX_RECEBIMENTOS_SCHEMA.includes(cabecalho),
        `Cabecalho ausente em INBOX_RECEBIMENTOS: ${cabecalho}.`
      );
    });
  assertRegressao(VENDAS_SCHEMA.includes('cliente'), 'Cabecalho cliente ausente em VENDAS_SCHEMA.');
  assertRegressao(
    VENDAS_SCHEMA.includes('referencia_venda'),
    'Cabecalho referencia_venda ausente em VENDAS_SCHEMA.'
  );

  const comprovante = {
    pagador_nome: 'Moveis Sao Jose',
    descricao: 'Pix recebido do cliente',
    referencia_transacao: 'PEDIDO 482',
    recebido_por: 'Bruno',
    forma_pagamento: 'PIX',
    data_recebimento: '2026-07-18'
  };
  const destinoExato = pontuarDestinoRecebimento_(comprovante, 900, {
    origem_tipo: 'VENDA', origem_id: 'VND_1', parcela_id: 'PAR_1',
    cliente: 'Moveis Sao Jose', referencia_venda: 'Pedido 482',
    origem_rotulo: 'Mesa de jantar', observacao: '', recebido_por: 'Bruno',
    forma_pagamento: 'PIX', data_prevista: '2026-07-18', valor_pendente: 900,
    label: 'Venda exata'
  });
  const destinoParcial = pontuarDestinoRecebimento_(comprovante, 300, {
    origem_tipo: 'VENDA', origem_id: 'VND_2', parcela_id: 'PAR_2',
    cliente: 'Outro cliente', referencia_venda: '', origem_rotulo: 'Armario',
    observacao: '', recebido_por: 'Socio', forma_pagamento: 'DINHEIRO',
    data_prevista: '2026-10-18', valor_pendente: 900, label: 'Venda parcial'
  });
  assertRegressao(destinoExato.score >= 90, 'Venda com valor, cliente e referencia exatos deveria ter score alto.');
  assertRegressao(destinoExato.score > destinoParcial.score, 'Sugestao exata deveria superar recebimento parcial fraco.');
  assertRegressao(
    destinoParcial.motivos.includes('possivel recebimento parcial'),
    'Matching deve reconhecer recebimentos parciais.'
  );
  return { ok: true, score_exato: destinoExato.score, score_parcial: destinoParcial.score };
}

function regressionEstruturaComprovantesDrive() {
  const criarIterador = itens => {
    let indice = 0;
    return {
      hasNext: () => indice < itens.length,
      next: () => itens[indice++]
    };
  };
  const criarPasta = (id, nome, pai) => ({
    getId: () => id,
    getName: () => nome,
    getParents: () => criarIterador(pai ? [pai] : [])
  });
  const raiz = criarPasta('ROOT', COMPROVANTES_DRIVE_ROOT_FOLDER_NAME, null);
  const despesas = criarPasta('DESP', COMPROVANTES_DRIVE_DOMAIN_FOLDER_NAMES.despesas, raiz);
  const recebimentos = criarPasta('REC', COMPROVANTES_DRIVE_DOMAIN_FOLDER_NAMES.recebimentos, raiz);
  const montarDominio = (prefixo, pastaRaiz) => {
    const pastas = { raiz: pastaRaiz };
    Object.keys(COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES).forEach(chave => {
      pastas[chave] = criarPasta(
        `${prefixo}_${chave}`,
        COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES[chave],
        pastaRaiz
      );
    });
    return pastas;
  };
  const pastasDespesas = montarDominio('D', despesas);
  const pastasRecebimentos = montarDominio('R', recebimentos);
  assertRegressao(
    estruturaDominioComprovantesDriveValida_(pastasDespesas, 'despesas', raiz),
    'Despesas deve aceitar apenas a hierarquia raiz/Despesas/status.'
  );
  assertRegressao(
    estruturaDominioComprovantesDriveValida_(pastasRecebimentos, 'recebimentos', raiz),
    'Recebimentos deve aceitar apenas a hierarquia raiz/Recebimentos/status.'
  );

  const despesasComEntradaNaRaiz = {
    ...pastasDespesas,
    entrada: criarPasta('D_ENTRADA_ANTIGA', COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES.entrada, raiz)
  };
  assertRegressao(
    !estruturaDominioComprovantesDriveValida_(despesasComEntradaNaRaiz, 'despesas', raiz),
    'Uma pasta de status diretamente na raiz comum deve ser rejeitada.'
  );
  const entradaNova = pastasDespesas.entrada;
  const entradaLegada = despesasComEntradaNaRaiz.entrada;
  assertRegressao(
    selecionarPastaDestinoStatusComprovantesDrive_(
      [entradaNova, entradaLegada],
      entradaLegada,
      despesas
    )?.getId() === entradaLegada.getId(),
    'A migracao deve preservar o ID da pasta legada salva, mesmo que ja exista um destino novo.'
  );
  assertRegressao(
    resolverRaizComumComprovantesDrive_(pastasRecebimentos.entrada, COMPROVANTES_DRIVE_ROOT_FOLDER_NAME)?.getId() === raiz.getId(),
    'A raiz comum deve ser resolvida a partir de uma pasta de status aninhada.'
  );
  assertRegressao(
    obterChaveRaizComprovantesDrive_('dev') === `${COMPROVANTES_DRIVE_COMMON_ROOT_PROP}_DEV`,
    'A raiz DEV deve usar propriedade independente de PROD.'
  );
  assertRegressao(
    obterChavePastaInboxDespesasDrive_('entrada', 'dev').endsWith('_DEV') &&
      obterChavePastaInboxRecebimentosDrive_('entrada', 'dev').endsWith('_DEV'),
    'Os caminhos de despesas e recebimentos em DEV devem usar propriedades independentes.'
  );
  return { ok: true, raiz: raiz.getName(), dominios: [despesas.getName(), recebimentos.getName()] };
}

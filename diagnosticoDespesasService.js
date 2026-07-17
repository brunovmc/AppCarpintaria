function diagnosticarDespesasProd() {
  return executarDiagnosticoDespesas_('prod');
}

function diagnosticarDespesasProdResumo() {
  return executarDiagnosticoDespesasResumo_('prod');
}

function diagnosticarDespesasDev() {
  return executarDiagnosticoDespesas_('dev');
}

function diagnosticarDespesasDevResumo() {
  return executarDiagnosticoDespesasResumo_('dev');
}

function diagnosticarCriacaoDespesaProd() {
  return executarDiagnosticoCriacaoDespesa_('prod');
}

function diagnosticarCriacaoDespesaDev() {
  return executarDiagnosticoCriacaoDespesa_('dev');
}

function compararDiagnosticoDespesasProdDev() {
  const resultado = {
    gerado_em: new Date().toISOString(),
    prod: diagnosticarFluxoDespesas_('prod'),
    dev: diagnosticarFluxoDespesas_('dev')
  };

  resultado.comparacao = compararSchemasDiagnosticoDespesas_(
    resultado.prod,
    resultado.dev
  );
  registrarLogDiagnosticoDespesas_('compararDiagnosticoDespesasProdDev', resultado);
  return resultado;
}

function executarDiagnosticoDespesas_(ambiente) {
  const resultado = diagnosticarFluxoDespesas_(ambiente);
  registrarLogDiagnosticoDespesas_(`diagnosticarDespesas:${resultado.ambiente}`, resultado);
  return resultado;
}

function executarDiagnosticoDespesasResumo_(ambiente) {
  const completo = diagnosticarFluxoDespesas_(ambiente);
  const resumo = montarResumoCompactoDiagnosticoDespesas_(completo);
  registrarLogDiagnosticoDespesas_(`diagnosticarDespesasResumo:${resumo.ambiente}`, resumo);
  return resumo;
}

function montarResumoCompactoDiagnosticoDespesas_(resultado) {
  const abas = resultado.abas || {};
  const despesas = abas.despesas_gerais || {};
  const parcelas = abas.parcelas_financeiras || {};
  const resumo = resultado.resumo || {};
  return {
    ambiente: resultado.ambiente,
    gerado_em: resultado.gerado_em,
    spreadsheet_id: resultado.spreadsheet_id,
    erro: resultado.erro || '',
    alertas: resultado.alertas || [],
    abas_ok: {
      despesas_gerais: {
        existe: !!despesas.existe,
        ok: !!despesas.ok,
        total_linhas: Number(despesas.total_linhas || 0),
        total_colunas: Number(despesas.total_colunas || 0),
        colunas_esperadas_ausentes: despesas.colunas_esperadas_ausentes || [],
        headers_duplicados_exatos: despesas.headers_duplicados_exatos || [],
        headers_duplicados_normalizados: despesas.headers_duplicados_normalizados || []
      },
      parcelas_financeiras: {
        existe: !!parcelas.existe,
        ok: !!parcelas.ok,
        total_linhas: Number(parcelas.total_linhas || 0),
        total_colunas: Number(parcelas.total_colunas || 0),
        colunas_esperadas_ausentes: parcelas.colunas_esperadas_ausentes || [],
        headers_duplicados_exatos: parcelas.headers_duplicados_exatos || [],
        headers_duplicados_normalizados: parcelas.headers_duplicados_normalizados || []
      }
    },
    resumo: {
      total_despesas_linhas: Number(resumo.total_despesas_linhas || 0),
      total_despesas_ativas: Number(resumo.total_despesas_ativas || 0),
      total_parcelas_despesa_ativas: Number(resumo.total_parcelas_despesa_ativas || 0),
      total_despesas_ativas_sem_id: Number(resumo.total_despesas_ativas_sem_id || 0),
      total_ids_despesa_duplicados: Number(resumo.total_ids_despesa_duplicados || 0),
      total_despesas_ativas_sem_parcelas: Number(resumo.total_despesas_ativas_sem_parcelas || 0),
      total_parcelas_despesa_orfas: Number(resumo.total_parcelas_despesa_orfas || 0)
    },
    ids_despesa_duplicados: resumo.ids_despesa_duplicados || [],
    despesas_ativas_sem_id: resumo.despesas_ativas_sem_id || [],
    despesas_ativas_sem_parcelas_amostra: resumo.despesas_ativas_sem_parcelas_amostra || [],
    parcelas_despesa_orfas_amostra: resumo.parcelas_despesa_orfas_amostra || []
  };
}

function executarDiagnosticoCriacaoDespesa_(ambiente) {
  const env = normalizarAmbienteDiagnosticoDespesas_(ambiente);
  const resultado = {
    ambiente: env,
    gerado_em: new Date().toISOString(),
    ok: false,
    etapa: 'inicio',
    payload: {},
    despesa_id: '',
    cleanup: null,
    erro: '',
    stack: ''
  };
  let envAnterior = '';
  let clientRequestId = '';

  try {
    if (typeof getUserDbEnvironment_ === 'function') {
      envAnterior = getUserDbEnvironment_();
    }
    if (typeof setUserDbEnvironment_ === 'function') {
      setUserDbEnvironment_(env);
    }

    resultado.etapa = 'montar_payload';
    const validacoes = obterValidacoes(true);
    const categorias = Array.isArray(validacoes?.categoriasDespesas)
      ? validacoes.categoriasDespesas.map(v => String(v || '').trim()).filter(Boolean)
      : [];
    const categoria = categorias[0] || 'DIAGNOSTICO';
    const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    clientRequestId = `DIAG_DESP_${env}_${Date.now()}`;
    const payload = {
      descricao: `Diagnostico criacao despesa ${env}`,
      categoria,
      fornecedor: '',
      pago_por: '',
      valor_total: 1.23,
      data_competencia: hoje,
      data_vencimento: hoje,
      data_pagamento: '',
      forma_pagamento: '',
      parcelas: 1,
      parcelas_detalhe: [],
      fixo: false,
      observacao: 'Registro temporario gerado por diagnostico.',
      client_request_id: clientRequestId
    };
    resultado.payload = payload;

    resultado.etapa = 'criarDespesaGeral';
    const criada = criarDespesaGeral(payload);
    resultado.despesa_id = String(criada?.ID || '').trim();
    resultado.retorno_criacao = criada || null;

    if (!resultado.despesa_id) {
      resultado.etapa = 'localizar_despesa_criada';
      const sheet = getSheet(ABA_DESPESAS_GERAIS);
      const encontrada = buscarDespesaPorClientRequestIdFinanceiro(sheet, clientRequestId);
      resultado.despesa_id = String(encontrada?.ID || '').trim();
    }

    if (!resultado.despesa_id) {
      throw new Error('Criacao nao retornou ID e nao foi possivel localizar a despesa por client_request_id.');
    }

    resultado.etapa = 'cleanup';
    resultado.cleanup = deletarDespesaGeral(resultado.despesa_id);
    resultado.ok = true;
    resultado.etapa = 'concluido';
  } catch (error) {
    resultado.erro = error && error.message ? error.message : String(error);
    resultado.stack = error && error.stack ? String(error.stack) : '';
    try {
      const sheet = getSheet(ABA_DESPESAS_GERAIS);
      const encontrada = buscarDespesaPorClientRequestIdFinanceiro(sheet, clientRequestId);
      const id = String(encontrada?.ID || resultado.despesa_id || '').trim();
      if (id) {
        resultado.cleanup = deletarDespesaGeral(id);
        resultado.despesa_id = id;
      }
    } catch (cleanupError) {
      resultado.cleanup_erro = cleanupError && cleanupError.message ? cleanupError.message : String(cleanupError);
    }
  } finally {
    try {
      if (envAnterior && typeof setUserDbEnvironment_ === 'function') {
        setUserDbEnvironment_(envAnterior);
      }
    } catch (error) {
      resultado.restore_env_erro = error && error.message ? error.message : String(error);
    }
    registrarLogDiagnosticoDespesas_(`diagnosticarCriacaoDespesa:${env}`, resultado);
  }

  return resultado;
}

function diagnosticarFluxoDespesas_(ambiente) {
  const env = normalizarAmbienteDiagnosticoDespesas_(ambiente);
  const resultado = {
    ambiente: env,
    gerado_em: new Date().toISOString(),
    spreadsheet_id: '',
    abas: {},
    resumo: {},
    alertas: []
  };

  try {
    resultado.spreadsheet_id = obterSpreadsheetIdDiagnosticoDespesas_(env);
    const ss = SpreadsheetApp.openById(resultado.spreadsheet_id);
    const sheetDespesas = ss.getSheetByName(ABA_DESPESAS_GERAIS);
    const sheetParcelas = ss.getSheetByName(ABA_PARCELAS_FINANCEIRAS);

    resultado.abas.despesas_gerais = auditarAbaDiagnosticoDespesas_(
      sheetDespesas,
      DESPESAS_GERAIS_SCHEMA
    );
    resultado.abas.parcelas_financeiras = auditarAbaDiagnosticoDespesas_(
      sheetParcelas,
      PARCELAS_FINANCEIRAS_SCHEMA
    );

    resultado.resumo = resumirIntegridadeDespesas_(sheetDespesas, sheetParcelas);
    resultado.alertas = montarAlertasDiagnosticoDespesas_(resultado);
  } catch (error) {
    resultado.erro = error && error.message ? error.message : String(error);
    resultado.alertas.push('Falha ao executar diagnostico. Verifique permissoes e IDs das planilhas.');
  }

  return resultado;
}

function normalizarAmbienteDiagnosticoDespesas_(ambiente) {
  const env = String(ambiente || '').trim().toLowerCase();
  return env === 'dev' ? 'dev' : 'prod';
}

function obterSpreadsheetIdDiagnosticoDespesas_(ambiente) {
  if (typeof getDataSpreadsheetIdAtivo === 'function') {
    return String(getDataSpreadsheetIdAtivo({
      targetEnv: ambiente,
      skipAccessCheck: true
    }) || '').trim();
  }

  if (ambiente === 'dev') {
    return String(typeof DATA_SPREADSHEET_ID_DEV === 'string' ? DATA_SPREADSHEET_ID_DEV : '').trim();
  }

  return String(
    (typeof DATA_SPREADSHEET_ID_PROD === 'string' ? DATA_SPREADSHEET_ID_PROD : '')
      || (typeof DATA_SPREADSHEET_ID === 'string' ? DATA_SPREADSHEET_ID : '')
      || ''
  ).trim();
}

function auditarAbaDiagnosticoDespesas_(sheet, schema) {
  const esperado = Array.isArray(schema) ? schema : [];
  if (!sheet) {
    return {
      existe: false,
      ok: false,
      total_linhas: 0,
      total_colunas: 0,
      headers: [],
      colunas_esperadas_ausentes: esperado,
      colunas_esperadas_com_espaco: [],
      headers_duplicados_exatos: [],
      headers_duplicados_normalizados: [],
      posicoes_schema: {}
    };
  }

  const totalColunas = sheet.getLastColumn();
  const headers = totalColunas > 0
    ? sheet.getRange(1, 1, 1, totalColunas).getValues()[0].map(v => String(v ?? ''))
    : [];
  const headersNormalizados = headers.map(h => String(h || '').trim());
  const ausentes = esperado.filter(coluna => !headers.includes(coluna));
  const ausentesNormalizadas = esperado.filter(coluna => !headersNormalizados.includes(coluna));
  const colunasComEspaco = headers
    .map((h, idx) => ({
      coluna: idx + 1,
      header: h,
      normalizado: String(h || '').trim()
    }))
    .filter(item => item.header !== item.normalizado);
  const posicoes = {};

  esperado.forEach(coluna => {
    posicoes[coluna] = {
      exata: headers.indexOf(coluna) + 1,
      normalizada: headersNormalizados.indexOf(coluna) + 1
    };
  });

  return {
    existe: true,
    ok: ausentes.length === 0 &&
      detectarDuplicadosDiagnosticoDespesas_(headers).length === 0 &&
      detectarDuplicadosDiagnosticoDespesas_(headersNormalizados).length === 0,
    total_linhas: sheet.getLastRow(),
    total_colunas: totalColunas,
    headers,
    colunas_esperadas_ausentes: ausentes,
    colunas_esperadas_ausentes_mesmo_trim: ausentesNormalizadas,
    colunas_esperadas_com_espaco: colunasComEspaco,
    headers_duplicados_exatos: detectarDuplicadosDiagnosticoDespesas_(headers),
    headers_duplicados_normalizados: detectarDuplicadosDiagnosticoDespesas_(headersNormalizados),
    posicoes_schema: posicoes
  };
}

function detectarDuplicadosDiagnosticoDespesas_(lista) {
  const vistos = {};
  const duplicados = {};
  (Array.isArray(lista) ? lista : []).forEach((valor, idx) => {
    const chave = String(valor ?? '');
    if (!chave) return;
    if (!vistos[chave]) vistos[chave] = [];
    vistos[chave].push(idx + 1);
  });
  Object.keys(vistos).forEach(chave => {
    if (vistos[chave].length > 1) {
      duplicados[chave] = vistos[chave];
    }
  });
  return Object.keys(duplicados).map(chave => ({
    header: chave,
    colunas: duplicados[chave]
  }));
}

function lerObjetosAbaDiagnosticoDespesas_(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (!Array.isArray(data) || data.length <= 1) return [];

  const headers = data[0].map(h => String(h ?? ''));
  return data.slice(1).map((row, idx) => {
    const obj = { __linha: idx + 2 };
    headers.forEach((header, colIdx) => {
      obj[header] = row[colIdx];
    });
    return obj;
  });
}

function resumirIntegridadeDespesas_(sheetDespesas, sheetParcelas) {
  const despesas = lerObjetosAbaDiagnosticoDespesas_(sheetDespesas);
  const parcelas = lerObjetosAbaDiagnosticoDespesas_(sheetParcelas);
  const despesasAtivas = despesas.filter(ehLinhaAtivaDiagnosticoDespesas_);
  const parcelasAtivasDespesa = parcelas.filter(parcela =>
    ehLinhaAtivaDiagnosticoDespesas_(parcela) &&
    String(parcela.origem_tipo || '').trim().toUpperCase() === ORIGEM_TIPO_DESPESA
  );
  const idsAtivos = {};
  const parcelasPorOrigem = {};

  despesasAtivas.forEach(despesa => {
    const id = String(despesa.ID || '').trim();
    if (id) idsAtivos[id] = true;
  });

  parcelasAtivasDespesa.forEach(parcela => {
    const origemId = String(parcela.origem_id || '').trim();
    if (!origemId) return;
    parcelasPorOrigem[origemId] = (parcelasPorOrigem[origemId] || 0) + 1;
  });

  const despesasSemId = despesasAtivas
    .filter(despesa => !String(despesa.ID || '').trim())
    .map(compactarDespesaDiagnostico_);
  const despesasSemAtivoTrue = despesas
    .filter(despesa => String(despesa.ID || '').trim() && !ehLinhaAtivaDiagnosticoDespesas_(despesa))
    .slice(-10)
    .map(compactarDespesaDiagnostico_);
  const idsDuplicados = detectarDuplicadosIdsDiagnosticoDespesas_(despesas, 'ID');
  const despesasSemParcelasTodas = despesasAtivas
    .filter(despesa => {
      const id = String(despesa.ID || '').trim();
      return id && !parcelasPorOrigem[id];
    });
  const despesasSemParcelas = despesasSemParcelasTodas
    .slice(-20)
    .map(compactarDespesaDiagnostico_);
  const parcelasOrfasTodas = parcelasAtivasDespesa
    .filter(parcela => {
      const origemId = String(parcela.origem_id || '').trim();
      return origemId && !idsAtivos[origemId];
    });
  const parcelasOrfas = parcelasOrfasTodas
    .slice(-20)
    .map(compactarParcelaDiagnostico_);

  return {
    total_despesas_linhas: despesas.length,
    total_despesas_ativas: despesasAtivas.length,
    total_parcelas_despesa_ativas: parcelasAtivasDespesa.length,
    total_despesas_ativas_sem_id: despesasSemId.length,
    total_ids_despesa_duplicados: idsDuplicados.length,
    total_despesas_ativas_sem_parcelas: despesasSemParcelasTodas.length,
    total_parcelas_despesa_orfas: parcelasOrfasTodas.length,
    total_parcelas_despesa_orfas_amostra: parcelasOrfas.length,
    ids_despesa_duplicados: idsDuplicados.slice(0, 20),
    ultimas_despesas: despesas.slice(-10).reverse().map(compactarDespesaDiagnostico_),
    despesas_ativas_sem_id: despesasSemId.slice(0, 20),
    despesas_com_id_inativas_amostra: despesasSemAtivoTrue,
    despesas_ativas_sem_parcelas_amostra: despesasSemParcelas,
    parcelas_despesa_orfas_amostra: parcelasOrfas
  };
}

function ehLinhaAtivaDiagnosticoDespesas_(linha) {
  return String(linha && linha.ativo !== undefined ? linha.ativo : '')
    .trim()
    .toLowerCase() === 'true';
}

function detectarDuplicadosIdsDiagnosticoDespesas_(linhas, campo) {
  const mapa = {};
  (Array.isArray(linhas) ? linhas : []).forEach(linha => {
    const id = String(linha && linha[campo] !== undefined ? linha[campo] : '').trim();
    if (!id) return;
    if (!mapa[id]) mapa[id] = [];
    mapa[id].push(Number(linha.__linha || 0));
  });
  return Object.keys(mapa)
    .filter(id => mapa[id].length > 1)
    .map(id => ({ id, linhas: mapa[id] }));
}

function compactarDespesaDiagnostico_(despesa) {
  return {
    linha: Number(despesa.__linha || 0),
    ID: String(despesa.ID || '').trim(),
    ativo: valorDiagnosticoDespesas_(despesa.ativo),
    descricao: String(despesa.descricao || '').trim(),
    valor_total: valorDiagnosticoDespesas_(despesa.valor_total),
    criado_em: valorDiagnosticoDespesas_(despesa.criado_em),
    data_competencia: valorDiagnosticoDespesas_(despesa.data_competencia),
    data_vencimento: valorDiagnosticoDespesas_(despesa.data_vencimento),
    forma_pagamento: String(despesa.forma_pagamento || '').trim(),
    parcelas: valorDiagnosticoDespesas_(despesa.parcelas)
  };
}

function compactarParcelaDiagnostico_(parcela) {
  return {
    linha: Number(parcela.__linha || 0),
    ID: String(parcela.ID || '').trim(),
    origem_tipo: String(parcela.origem_tipo || '').trim(),
    origem_id: String(parcela.origem_id || '').trim(),
    ativo: valorDiagnosticoDespesas_(parcela.ativo),
    parcela_numero: valorDiagnosticoDespesas_(parcela.parcela_numero),
    valor_previsto: valorDiagnosticoDespesas_(parcela.valor_previsto),
    valor_pago: valorDiagnosticoDespesas_(parcela.valor_pago),
    status: String(parcela.status || '').trim()
  };
}

function valorDiagnosticoDespesas_(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor.toISOString();
  }
  if (valor === null || valor === undefined) return '';
  return valor;
}

function montarAlertasDiagnosticoDespesas_(resultado) {
  const alertas = [];
  const despesas = resultado.abas && resultado.abas.despesas_gerais;
  const parcelas = resultado.abas && resultado.abas.parcelas_financeiras;
  const resumo = resultado.resumo || {};

  if (!despesas || !despesas.existe) {
    alertas.push('Aba DESPESAS_GERAIS nao encontrada.');
  } else {
    if (despesas.colunas_esperadas_ausentes.length > 0) {
      alertas.push(`DESPESAS_GERAIS com colunas esperadas ausentes: ${despesas.colunas_esperadas_ausentes.join(', ')}`);
    }
    if (despesas.headers_duplicados_exatos.length > 0 || despesas.headers_duplicados_normalizados.length > 0) {
      alertas.push('DESPESAS_GERAIS possui headers duplicados.');
    }
    if (despesas.colunas_esperadas_com_espaco.length > 0) {
      alertas.push('DESPESAS_GERAIS possui headers com espacos antes/depois.');
    }
  }

  if (!parcelas || !parcelas.existe) {
    alertas.push('Aba PARCELAS_FINANCEIRAS nao encontrada.');
  } else {
    if (parcelas.colunas_esperadas_ausentes.length > 0) {
      alertas.push(`PARCELAS_FINANCEIRAS com colunas esperadas ausentes: ${parcelas.colunas_esperadas_ausentes.join(', ')}`);
    }
    if (parcelas.headers_duplicados_exatos.length > 0 || parcelas.headers_duplicados_normalizados.length > 0) {
      alertas.push('PARCELAS_FINANCEIRAS possui headers duplicados.');
    }
  }

  if (Number(resumo.total_despesas_ativas_sem_id || 0) > 0) {
    alertas.push('Existem despesas ativas sem ID.');
  }
  if (Number(resumo.total_ids_despesa_duplicados || 0) > 0) {
    alertas.push('Existem IDs duplicados em DESPESAS_GERAIS.');
  }
  if (Number(resumo.total_despesas_ativas_sem_parcelas || 0) > 0) {
    alertas.push('Existem despesas ativas sem parcelas financeiras.');
  }
  if (Number(resumo.total_parcelas_despesa_orfas || resumo.total_parcelas_despesa_orfas_amostra || 0) > 0) {
    alertas.push('Existem parcelas de despesa apontando para origem inexistente/inativa.');
  }

  if (alertas.length === 0) {
    alertas.push('Nenhum problema estrutural evidente encontrado no diagnostico somente leitura.');
  }

  return alertas;
}

function compararSchemasDiagnosticoDespesas_(prod, dev) {
  return {
    despesas_gerais: compararAbaSchemaDiagnosticoDespesas_(
      prod && prod.abas ? prod.abas.despesas_gerais : null,
      dev && dev.abas ? dev.abas.despesas_gerais : null
    ),
    parcelas_financeiras: compararAbaSchemaDiagnosticoDespesas_(
      prod && prod.abas ? prod.abas.parcelas_financeiras : null,
      dev && dev.abas ? dev.abas.parcelas_financeiras : null
    )
  };
}

function compararAbaSchemaDiagnosticoDespesas_(prod, dev) {
  const headersProd = prod && Array.isArray(prod.headers) ? prod.headers : [];
  const headersDev = dev && Array.isArray(dev.headers) ? dev.headers : [];
  return {
    headers_iguais_mesma_ordem: JSON.stringify(headersProd) === JSON.stringify(headersDev),
    apenas_prod: headersProd.filter(h => !headersDev.includes(h)),
    apenas_dev: headersDev.filter(h => !headersProd.includes(h)),
    prod_colunas_ausentes: prod ? prod.colunas_esperadas_ausentes : [],
    dev_colunas_ausentes: dev ? dev.colunas_esperadas_ausentes : []
  };
}

function registrarLogDiagnosticoDespesas_(nome, payload) {
  try {
    console.log(`${nome}\n${JSON.stringify(payload, null, 2)}`);
  } catch (error) {
    Logger.log(payload);
  }
}

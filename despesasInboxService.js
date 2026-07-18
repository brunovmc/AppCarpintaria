const ABA_INBOX_DESPESAS = 'INBOX_DESPESAS';
const OPENAI_COMPROVANTES_MODEL_DEFAULT = 'gpt-5-mini';
const OPENAI_COMPROVANTES_REASONING_EFFORT_DEFAULT = 'low';
const INBOX_DESPESAS_RETRY_LEITURA_TENTATIVAS = 5;
const INBOX_DESPESAS_RETRY_LEITURA_MS = 300;

const INBOX_DESPESAS_SCHEMA = [
  'ID',
  'status',
  'origem_tipo',
  'arquivo_nome',
  'arquivo_mime',
  'arquivo_drive_id',
  'arquivo_url',
  'imagem_hash',
  'descricao',
  'categoria',
  'fornecedor',
  'pago_por',
  'valor_total',
  'data_competencia',
  'data_vencimento',
  'data_pagamento',
  'forma_pagamento',
  'parcelas',
  'observacao',
  'dados_extraidos_json',
  'confianca',
  'alertas_json',
  'erro',
  'despesa_id_confirmada',
  'criado_em',
  'atualizado_em',
  'confirmado_em',
  'descartado_em'
];

function listarInboxDespesas(statusFiltro, ambiente) {
  return executarComAmbienteInboxDespesas_(ambiente, () => listarInboxDespesasNoAmbienteAtual_(statusFiltro));
}

function listarInboxDespesasNoAmbienteAtual_(statusFiltro) {
  const sheet = getSheet(ABA_INBOX_DESPESAS);
  if (!sheet) return [];

  ensureSchema(sheet, INBOX_DESPESAS_SCHEMA);
  const filtro = String(statusFiltro || '').trim().toUpperCase();
  const statusVisiveis = filtro
    ? [filtro]
    : ['PENDENTE', 'ERRO'];

  return rowsToObjects(sheet)
    .map(normalizarLinhaInboxDespesa_)
    .filter(item => {
      if (filtro === 'TODOS' || filtro === 'ALL') return true;
      return statusVisiveis.includes(String(item.status || '').trim().toUpperCase());
    })
    .sort((a, b) => {
      const da = parseDataFinanceiro(a.atualizado_em || a.criado_em)?.getTime() || 0;
      const db = parseDataFinanceiro(b.atualizado_em || b.criado_em)?.getTime() || 0;
      if (da !== db) return db - da;
      return String(b.ID || '').localeCompare(String(a.ID || ''));
    });
}

function processarComprovanteDespesaUpload(payload) {
  const ambiente = payload?.ambiente || payload?._db_env || payload?.db_env || '';
  return executarComAmbienteInboxDespesas_(ambiente, () => processarComprovanteDespesaUploadNoAmbienteAtual_(payload));
}

function processarComprovanteDespesaUploadNoAmbienteAtual_(payload) {
  assertCanWrite('Inbox de despesas');

  const upload = normalizarUploadInboxDespesa_(payload);
  const agoraAtual = new Date();
  const arquivo = salvarArquivoInboxDespesa_(upload);
  const base = {
    ID: gerarId('IBXDESP'),
    origem_tipo: 'UPLOAD_APP',
    arquivo_nome: upload.nome,
    arquivo_mime: upload.mime,
    arquivo_drive_id: arquivo.id || '',
    arquivo_url: arquivo.url || '',
    imagem_hash: upload.hash,
    despesa_id_confirmada: '',
    criado_em: agoraAtual,
    atualizado_em: agoraAtual,
    confirmado_em: '',
    descartado_em: ''
  };

  try {
    const extraido = extrairDespesaComOpenAI_(upload.dataUrl);
    const rascunho = normalizarDespesaExtraidaInbox_(extraido);
    const alertas = [
      ...rascunho.alertas,
      ...(arquivo.erro ? [arquivo.erro] : [])
    ];
    const linha = {
      ...base,
      status: 'PENDENTE',
      descricao: rascunho.descricao,
      categoria: rascunho.categoria,
      fornecedor: rascunho.fornecedor,
      pago_por: rascunho.pago_por,
      valor_total: rascunho.valor_total,
      data_competencia: rascunho.data_competencia,
      data_vencimento: rascunho.data_vencimento,
      data_pagamento: rascunho.data_pagamento,
      forma_pagamento: rascunho.forma_pagamento,
      parcelas: rascunho.parcelas,
      observacao: rascunho.observacao,
      dados_extraidos_json: JSON.stringify(extraido || {}),
      confianca: rascunho.confianca,
      alertas_json: serializarListaInboxDespesa_(alertas),
      erro: ''
    };
    insert(ABA_INBOX_DESPESAS, linha, INBOX_DESPESAS_SCHEMA);
    return obterInboxDespesaPorId_(linha.ID);
  } catch (error) {
    const linhaErro = {
      ...base,
      status: 'ERRO',
      descricao: '',
      categoria: '',
      fornecedor: '',
      pago_por: '',
      valor_total: 0,
      data_competencia: '',
      data_vencimento: '',
      data_pagamento: '',
      forma_pagamento: '',
      parcelas: 1,
      observacao: '',
      dados_extraidos_json: '',
      confianca: 0,
      alertas_json: serializarListaInboxDespesa_([arquivo.erro].filter(Boolean)),
      erro: error?.message || String(error || 'Falha ao processar comprovante.')
    };
    insert(ABA_INBOX_DESPESAS, linhaErro, INBOX_DESPESAS_SCHEMA);
    return obterInboxDespesaPorId_(linhaErro.ID);
  }
}

function confirmarInboxDespesa(id, payloadOverride) {
  const ambiente = payloadOverride?.ambiente || payloadOverride?._db_env || payloadOverride?.db_env || '';
  return executarComAmbienteInboxDespesas_(ambiente, () => confirmarInboxDespesaNoAmbienteAtual_(id, payloadOverride));
}

function confirmarInboxDespesaNoAmbienteAtual_(id, payloadOverride) {
  assertCanWrite('Confirmacao da Inbox de despesas');

  const inboxId = String(id || '').trim();
  if (!inboxId) {
    throw new Error('ID da Inbox e obrigatorio.');
  }

  const item = obterInboxDespesaPorId_(inboxId);
  if (!item) {
    throw new Error('Item da Inbox nao encontrado.');
  }

  const status = String(item.status || '').trim().toUpperCase();
  if (status === 'DESCARTADO') {
    throw new Error('Este item da Inbox ja foi descartado.');
  }

  const payload = montarPayloadConfirmacaoInboxDespesa_(item, payloadOverride);
  const traceId = Utilities.getUuid();
  const ambienteAtivo = (typeof getUserDbEnvironment_ === 'function')
    ? String(getUserDbEnvironment_() || '').trim()
    : '';
  registrarDiagnosticoConfirmacaoInbox_('inicio', {
    trace_id: traceId,
    ambiente: ambienteAtivo,
    inbox_id: inboxId,
    inbox_status: status,
    client_request_id: payload.client_request_id
  });

  if (status === 'CONFIRMADO') {
    const despesaPorId = item.despesa_id_confirmada
      ? buscarDespesaGeralAtivaPorIdInbox_(item.despesa_id_confirmada)
      : null;
    const despesaPorClient = buscarDespesaGeralPorClientRequestIdInbox_(payload.client_request_id);
    const despesaExistente = despesaPorId ||
      (String(despesaPorClient?.ativo || '').toLowerCase() === 'true' ? despesaPorClient : null);

    if (despesaExistente) {
      const despesaExistenteId = String(despesaExistente?.ID || '').trim();
      const respostaExistente = {
        ok: true,
        despesa_id: despesaExistenteId,
        inbox: item,
        despesa: {
          ...despesaExistente,
          ID: despesaExistenteId
        },
        diagnostico_confirmacao: montarDiagnosticoRespostaConfirmacaoInbox_({
          trace_id: traceId,
          ambiente: ambienteAtivo,
          inbox_id: inboxId,
          despesa_id_criada: despesaExistenteId,
          despesa_id_retornada: despesaExistenteId,
          origem_retorno: despesaPorId ? 'inbox_despesa_id' : 'client_request_id',
          tentativas_leitura: 1,
          resposta_reutilizada: true
        })
      };
      registrarDiagnosticoConfirmacaoInbox_('sucesso_reutilizado', respostaExistente.diagnostico_confirmacao);
      return respostaExistente;
    }

    updateById(ABA_INBOX_DESPESAS, 'ID', inboxId, {
      status: 'ERRO',
      despesa_id_confirmada: '',
      erro: 'Item estava confirmado, mas nenhuma despesa correspondente foi encontrada. Reenviando confirmacao.',
      atualizado_em: new Date()
    }, INBOX_DESPESAS_SCHEMA);
  }

  try {
    const despesa = criarDespesaGeral(payload);
    const despesaId = String(despesa?.ID || '').trim();
    if (!despesaId) {
      throw new Error('Nao foi possivel criar a despesa a partir da Inbox.');
    }

    const leituraConfirmada = aguardarDespesaGeralConfirmadaInbox_(despesaId, payload.client_request_id);
    const despesaConfirmada = {
      ...(leituraConfirmada?.despesa || despesa || {}),
      ID: despesaId
    };

    const agoraAtual = new Date();
    updateById(ABA_INBOX_DESPESAS, 'ID', inboxId, {
      status: 'CONFIRMADO',
      descricao: payload.descricao,
      categoria: payload.categoria,
      fornecedor: payload.fornecedor,
      pago_por: payload.pago_por,
      valor_total: payload.valor_total,
      data_competencia: payload.data_competencia,
      data_vencimento: payload.data_vencimento,
      data_pagamento: payload.data_pagamento,
      forma_pagamento: payload.forma_pagamento,
      parcelas: payload.parcelas,
      observacao: payload.observacao,
      despesa_id_confirmada: despesaId,
      erro: '',
      atualizado_em: agoraAtual,
      confirmado_em: agoraAtual
    }, INBOX_DESPESAS_SCHEMA);

    const resposta = {
      ok: true,
      despesa_id: despesaId,
      inbox: obterInboxDespesaPorId_(inboxId),
      despesa: despesaConfirmada,
      diagnostico_confirmacao: montarDiagnosticoRespostaConfirmacaoInbox_({
        trace_id: traceId,
        ambiente: ambienteAtivo,
        inbox_id: inboxId,
        despesa_id_criada: despesaId,
        despesa_id_retornada: String(despesaConfirmada?.ID || '').trim(),
        origem_retorno: leituraConfirmada?.origem || 'retorno_criacao',
        tentativas_leitura: leituraConfirmada?.tentativas || INBOX_DESPESAS_RETRY_LEITURA_TENTATIVAS,
        resposta_reutilizada: false
      })
    };
    registrarDiagnosticoConfirmacaoInbox_('sucesso', resposta.diagnostico_confirmacao);
    return resposta;
  } catch (error) {
    registrarDiagnosticoConfirmacaoInbox_('erro', {
      trace_id: traceId,
      ambiente: ambienteAtivo,
      inbox_id: inboxId,
      client_request_id: payload.client_request_id,
      erro: error?.message || String(error || 'Falha desconhecida')
    });
    updateById(ABA_INBOX_DESPESAS, 'ID', inboxId, {
      status: 'ERRO',
      descricao: payload.descricao,
      categoria: payload.categoria,
      fornecedor: payload.fornecedor,
      pago_por: payload.pago_por,
      valor_total: payload.valor_total,
      data_competencia: payload.data_competencia,
      data_vencimento: payload.data_vencimento,
      data_pagamento: payload.data_pagamento,
      forma_pagamento: payload.forma_pagamento,
      parcelas: payload.parcelas,
      observacao: payload.observacao,
      erro: error?.message || String(error || 'Falha ao confirmar item da Inbox.'),
      atualizado_em: new Date()
    }, INBOX_DESPESAS_SCHEMA);
    throw error;
  }
}

function montarDiagnosticoRespostaConfirmacaoInbox_(dados) {
  const item = dados || {};
  return {
    trace_id: String(item.trace_id || '').trim(),
    ambiente: String(item.ambiente || '').trim(),
    inbox_id: String(item.inbox_id || '').trim(),
    despesa_id_criada: String(item.despesa_id_criada || '').trim(),
    despesa_id_retornada: String(item.despesa_id_retornada || '').trim(),
    origem_retorno: String(item.origem_retorno || '').trim(),
    tentativas_leitura: Math.max(0, Math.floor(Number(item.tentativas_leitura || 0))),
    resposta_reutilizada: item.resposta_reutilizada === true
  };
}

function registrarDiagnosticoConfirmacaoInbox_(etapa, dados) {
  const registro = {
    evento: 'inbox_despesas.confirmacao',
    etapa: String(etapa || '').trim(),
    ...(dados || {})
  };
  try {
    console.log(JSON.stringify(registro));
  } catch (error) {
    try {
      Logger.log(registro);
    } catch (loggerError) {
      // Diagnostico nunca deve interferir na confirmacao.
    }
  }
}

function descartarInboxDespesa(id, ambiente) {
  return executarComAmbienteInboxDespesas_(ambiente, () => descartarInboxDespesaNoAmbienteAtual_(id));
}

function descartarInboxDespesaNoAmbienteAtual_(id) {
  assertCanWrite('Descarte da Inbox de despesas');

  const inboxId = String(id || '').trim();
  if (!inboxId) {
    throw new Error('ID da Inbox e obrigatorio.');
  }

  const item = obterInboxDespesaPorId_(inboxId);
  if (!item) {
    throw new Error('Item da Inbox nao encontrado.');
  }
  if (String(item.status || '').trim().toUpperCase() === 'CONFIRMADO') {
    throw new Error('Nao e possivel descartar um item ja confirmado.');
  }

  const agoraAtual = new Date();
  updateById(ABA_INBOX_DESPESAS, 'ID', inboxId, {
    status: 'DESCARTADO',
    atualizado_em: agoraAtual,
    descartado_em: agoraAtual
  }, INBOX_DESPESAS_SCHEMA);

  return {
    ok: true,
    inbox: obterInboxDespesaPorId_(inboxId)
  };
}

function diagnosticarInboxDespesasIA() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = String(props.getProperty('OPENAI_API_KEY') || '').trim();
  const modelo = getOpenAIComprovantesModel_();
  const reasoningEffort = getOpenAIComprovantesReasoningEffort_();
  const salvarArquivo = getSalvarArquivoInboxDespesas_();
  const itens = listarInboxDespesas('TODOS');
  return {
    ok: !!apiKey,
    openai_api_key_configurada: !!apiKey,
    modelo,
    reasoning_effort: reasoningEffort,
    salvar_arquivo_drive: salvarArquivo,
    total_inbox: itens.length,
    pendentes: itens.filter(i => String(i.status || '').toUpperCase() === 'PENDENTE').length,
    erros: itens.filter(i => String(i.status || '').toUpperCase() === 'ERRO').length
  };
}

function executarComAmbienteInboxDespesas_(ambiente, callback) {
  const env = normalizarAmbienteInboxDespesas_(ambiente);
  if (!env || typeof setUserDbEnvironment_ !== 'function' || typeof getUserDbEnvironment_ !== 'function') {
    return callback();
  }

  const anterior = getUserDbEnvironment_();
  if (anterior !== env) {
    setUserDbEnvironment_(env);
  }
  return callback();
}

function normalizarAmbienteInboxDespesas_(ambiente) {
  const raw = String(ambiente || '').trim().toLowerCase();
  if (raw === 'prod' || raw === 'dev') return raw;
  return '';
}

function diagnosticarUltimaConfirmacaoInboxDespesa() {
  return diagnosticarConfirmacaoInboxDespesa('');
}

function diagnosticarConfirmacaoInboxDespesa(id) {
  if (typeof assertCanRead === 'function') {
    assertCanRead('Diagnostico da Inbox de despesas');
  }

  const alvo = String(id || '').trim();
  const inbox = alvo
    ? obterInboxDespesaPorId_(alvo)
    : (listarInboxDespesas('TODOS')[0] || null);
  const ambiente = (typeof obterContextoBancoDados === 'function')
    ? obterContextoBancoDados()
    : {};
  const spreadsheetId = (typeof getDataSpreadsheetIdAtivo === 'function')
    ? String(getDataSpreadsheetIdAtivo({ skipAccessCheck: true }) || '').trim()
    : '';

  if (!inbox) {
    return {
      ok: false,
      erro: alvo ? 'Item da Inbox nao encontrado.' : 'Nenhum item encontrado em INBOX_DESPESAS.',
      ambiente: ambiente?.effective_env || ambiente?.selected_env || '',
      spreadsheet_id: spreadsheetId
    };
  }

  const despesaId = String(inbox.despesa_id_confirmada || '').trim();
  const clientRequestId = `INBOX_DESPESA_${String(inbox.ID || '').trim()}`;
  const despesaPorId = despesaId ? buscarDespesaGeralPorIdInbox_(despesaId) : null;
  const despesaPorClient = buscarDespesaGeralPorClientRequestIdInbox_(clientRequestId);
  const totalDespesas = contarLinhasDadosInbox_(ABA_DESPESAS_GERAIS, DESPESAS_GERAIS_SCHEMA);
  const totalInbox = contarLinhasDadosInbox_(ABA_INBOX_DESPESAS, INBOX_DESPESAS_SCHEMA);

  return {
    ok: !!(despesaPorId || despesaPorClient),
    ambiente: ambiente?.effective_env || ambiente?.selected_env || '',
    spreadsheet_id: spreadsheetId,
    inbox: resumirInboxDespesaDiagnostico_(inbox),
    client_request_id_esperado: clientRequestId,
    despesa_por_id: resumirDespesaDiagnosticoInbox_(despesaPorId),
    despesa_por_client_request_id: resumirDespesaDiagnosticoInbox_(despesaPorClient),
    totais: {
      inbox: totalInbox,
      despesas_gerais: totalDespesas
    }
  };
}

function obterInboxDespesaPorId_(id) {
  const alvo = String(id || '').trim();
  if (!alvo) return null;

  const sheet = getSheet(ABA_INBOX_DESPESAS);
  if (!sheet) return null;
  ensureSchema(sheet, INBOX_DESPESAS_SCHEMA);

  return rowsToObjects(sheet)
    .map(normalizarLinhaInboxDespesa_)
    .find(item => String(item.ID || '').trim() === alvo) || null;
}

function buscarDespesaGeralAtivaPorIdInbox_(id) {
  const item = buscarDespesaGeralPorIdInbox_(id);
  if (!item) return null;
  if (String(item.ativo).toLowerCase() !== 'true') return null;
  return item;
}

function aguardarDespesaGeralConfirmadaInbox_(despesaId, clientRequestId) {
  const tentativas = Math.max(1, INBOX_DESPESAS_RETRY_LEITURA_TENTATIVAS);
  for (let i = 0; i < tentativas; i++) {
    const porId = buscarDespesaGeralAtivaPorIdInbox_(despesaId);
    if (porId) {
      return {
        despesa: porId,
        origem: 'busca_id',
        tentativas: i + 1
      };
    }

    const porClient = buscarDespesaGeralPorClientRequestIdInbox_(clientRequestId);
    if (String(porClient?.ativo || '').toLowerCase() === 'true') {
      return {
        despesa: porClient,
        origem: 'busca_client_request_id',
        tentativas: i + 1
      };
    }

    if (i < tentativas - 1) {
      try {
        SpreadsheetApp.flush();
      } catch (error) {
        // sem acao
      }
      try {
        Utilities.sleep(INBOX_DESPESAS_RETRY_LEITURA_MS);
      } catch (error) {
        // sem acao
      }
    }
  }
  return null;
}

function buscarDespesaGeralPorIdInbox_(id) {
  return buscarDespesaGeralPorCampoInbox_('ID', id);
}

function buscarDespesaGeralPorClientRequestIdInbox_(clientRequestId) {
  return buscarDespesaGeralPorCampoInbox_('client_request_id', clientRequestId);
}

function buscarDespesaGeralPorCampoInbox_(campo, valor) {
  const field = String(campo || '').trim();
  const alvo = String(valor || '').trim();
  if (!field || !alvo) return null;

  const sheet = getSheet(ABA_DESPESAS_GERAIS);
  if (!sheet) return null;
  ensureSchema(sheet, DESPESAS_GERAIS_SCHEMA);

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastCol <= 0 || lastRow < 2) return null;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colIndex = headers.indexOf(field);
  if (colIndex < 0) return null;

  try {
    const encontrado = sheet
      .getRange(2, colIndex + 1, lastRow - 1, 1)
      .createTextFinder(alvo)
      .matchEntireCell(true)
      .matchCase(true)
      .findNext();
    if (encontrado) {
      return montarObjetoLinhaInbox_(sheet, encontrado.getRow(), headers);
    }
  } catch (error) {
    // fallback abaixo
  }

  const rows = rowsToObjects(sheet);
  return rows.find(item => String(item[field] || '').trim() === alvo) || null;
}

function montarObjetoLinhaInbox_(sheet, rowIndex, headers) {
  if (!sheet || rowIndex < 2) return null;
  const cols = Array.isArray(headers) && headers.length > 0
    ? headers
    : sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(rowIndex, 1, 1, cols.length).getValues()[0];
  const obj = {};
  cols.forEach((header, index) => {
    obj[header] = values[index];
  });
  return obj;
}

function descreverAmbienteAtivoInbox_() {
  const ambiente = (typeof obterContextoBancoDados === 'function')
    ? obterContextoBancoDados()
    : {};
  const env = String(ambiente?.effective_env || ambiente?.selected_env || '').trim();
  const spreadsheetId = (typeof getDataSpreadsheetIdAtivo === 'function')
    ? String(getDataSpreadsheetIdAtivo({ skipAccessCheck: true }) || '').trim()
    : '';
  const totalDespesas = contarLinhasDadosInbox_(ABA_DESPESAS_GERAIS, DESPESAS_GERAIS_SCHEMA);
  return `Ambiente=${env || 'N/A'}; planilha=${spreadsheetId || 'N/A'}; linhas_DESPESAS_GERAIS=${totalDespesas}.`;
}

function contarLinhasDadosInbox_(sheetName, schema) {
  const sheet = getSheet(sheetName);
  if (!sheet) return 0;
  ensureSchema(sheet, schema);
  return Math.max(0, sheet.getLastRow() - 1);
}

function resumirInboxDespesaDiagnostico_(item) {
  if (!item) return null;
  return {
    ID: String(item.ID || '').trim(),
    status: String(item.status || '').trim(),
    descricao: String(item.descricao || '').trim(),
    valor_total: round2Financeiro(parseNumeroBR(item.valor_total)),
    despesa_id_confirmada: String(item.despesa_id_confirmada || '').trim(),
    erro: String(item.erro || '').trim(),
    atualizado_em: String(item.atualizado_em || '').trim(),
    confirmado_em: String(item.confirmado_em || '').trim()
  };
}

function resumirDespesaDiagnosticoInbox_(item) {
  if (!item) return null;
  return {
    ID: String(item.ID || '').trim(),
    ativo: String(item.ativo || '').trim(),
    descricao: String(item.descricao || '').trim(),
    categoria: String(item.categoria || '').trim(),
    valor_total: round2Financeiro(parseNumeroBR(item.valor_total)),
    data_competencia: item.data_competencia ? formatarDataYmdFinanceiroSafe(item.data_competencia) : '',
    client_request_id: String(item.client_request_id || '').trim(),
    criado_em: item.criado_em ? formatarDataYmdHmFinanceiroSafe(item.criado_em) : ''
  };
}

function normalizarLinhaInboxDespesa_(row) {
  const item = { ...(row || {}) };
  return {
    ...item,
    ID: String(item.ID || '').trim(),
    status: String(item.status || '').trim().toUpperCase() || 'PENDENTE',
    origem_tipo: String(item.origem_tipo || '').trim(),
    arquivo_nome: String(item.arquivo_nome || '').trim(),
    arquivo_mime: String(item.arquivo_mime || '').trim(),
    arquivo_drive_id: String(item.arquivo_drive_id || '').trim(),
    arquivo_url: String(item.arquivo_url || '').trim(),
    imagem_hash: String(item.imagem_hash || '').trim(),
    descricao: String(item.descricao || '').trim(),
    categoria: String(item.categoria || '').trim(),
    fornecedor: String(item.fornecedor || '').trim(),
    pago_por: String(item.pago_por || '').trim(),
    valor_total: round2Financeiro(parseNumeroBR(item.valor_total)),
    data_competencia: item.data_competencia ? formatarDataYmdFinanceiroSafe(item.data_competencia) : '',
    data_vencimento: item.data_vencimento ? formatarDataYmdFinanceiroSafe(item.data_vencimento) : '',
    data_pagamento: item.data_pagamento ? formatarDataYmdFinanceiroSafe(item.data_pagamento) : '',
    forma_pagamento: String(item.forma_pagamento || '').trim(),
    parcelas: Math.max(1, Math.floor(parseNumeroBR(item.parcelas) || 1)),
    observacao: String(item.observacao || '').trim(),
    dados_extraidos: parseJsonInboxDespesa_(item.dados_extraidos_json, {}),
    confianca: Math.max(0, Math.min(1, Number(item.confianca || 0) || 0)),
    alertas: parseJsonInboxDespesa_(item.alertas_json, []),
    erro: String(item.erro || '').trim(),
    despesa_id_confirmada: String(item.despesa_id_confirmada || '').trim(),
    criado_em: item.criado_em ? formatarDataYmdHmFinanceiroSafe(item.criado_em) : '',
    atualizado_em: item.atualizado_em ? formatarDataYmdHmFinanceiroSafe(item.atualizado_em) : '',
    confirmado_em: item.confirmado_em ? formatarDataYmdHmFinanceiroSafe(item.confirmado_em) : '',
    descartado_em: item.descartado_em ? formatarDataYmdHmFinanceiroSafe(item.descartado_em) : ''
  };
}

function normalizarUploadInboxDespesa_(payload) {
  const dados = payload || {};
  const dataUrl = String(dados.data_url || dados.dataUrl || '').trim();
  const match = /^data:([^;]+);base64,(.+)$/i.exec(dataUrl);
  if (!match) {
    throw new Error('Imagem invalida. Envie um arquivo de imagem.');
  }

  const mime = String(dados.mime_type || dados.mime || match[1] || '').trim().toLowerCase();
  if (!mime || !mime.startsWith('image/')) {
    throw new Error('Apenas imagens sao aceitas nesta fase.');
  }

  const base64 = String(match[2] || '').trim();
  if (!base64) {
    throw new Error('Imagem sem conteudo.');
  }
  if (base64.length > 9 * 1024 * 1024) {
    throw new Error('Imagem muito grande. Tente uma foto menor ou recortada.');
  }

  let bytes;
  try {
    bytes = Utilities.base64Decode(base64);
  } catch (error) {
    throw new Error('Imagem em Base64 invalida.');
  }

  const nome = normalizarNomeArquivoInboxDespesa_(dados.nome || dados.name || `comprovante-${Date.now()}`);
  return {
    nome,
    mime,
    base64,
    bytes,
    dataUrl,
    hash: hashTextoInboxDespesa_(base64)
  };
}

function normalizarNomeArquivoInboxDespesa_(valor) {
  const nome = String(valor || '')
    .replace(/[\\/:*?"<>|]+/g, '-')
    .trim();
  return (nome || `comprovante-${Date.now()}`).slice(0, 180);
}

function salvarArquivoInboxDespesa_(upload) {
  if (!getSalvarArquivoInboxDespesas_()) {
    return { id: '', url: '', erro: '' };
  }

  try {
    const blob = Utilities.newBlob(upload.bytes, upload.mime, upload.nome);
    const file = DriveApp.createFile(blob).setName(upload.nome);
    return {
      id: file.getId(),
      url: file.getUrl(),
      erro: ''
    };
  } catch (error) {
    return {
      id: '',
      url: '',
      erro: `Arquivo nao salvo no Drive: ${error?.message || error}`
    };
  }
}

function getSalvarArquivoInboxDespesas_() {
  const valor = String(
    PropertiesService.getScriptProperties().getProperty('INBOX_DESPESAS_SALVAR_ARQUIVO') || ''
  ).trim().toLowerCase();
  return valor === 'true' || valor === '1' || valor === 'sim' || valor === 'yes';
}

function extrairDespesaComOpenAI_(dataUrl) {
  const apiKey = getOpenAIComprovantesApiKey_();
  const model = getOpenAIComprovantesModel_();
  const reasoning = getOpenAIComprovantesReasoning_(model);
  const payload = {
    model,
    input: [
      {
        role: 'user',
        content: [
          {
            type: 'input_text',
            text: montarPromptOpenAIInboxDespesa_()
          },
          {
            type: 'input_image',
            image_url: dataUrl
          }
        ]
      }
    ],
    text: {
      format: {
        type: 'json_schema',
        name: 'despesa_comprovante',
        strict: true,
        schema: getSchemaOpenAIInboxDespesa_()
      }
    },
    max_output_tokens: 3000
  };
  if (reasoning) {
    payload.reasoning = reasoning;
  }

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${apiKey}`
    },
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  });

  const code = response.getResponseCode();
  const body = response.getContentText();
  let json;
  try {
    json = JSON.parse(body);
  } catch (error) {
    throw new Error(`Resposta invalida da OpenAI (${code}).`);
  }

  if (code < 200 || code >= 300) {
    const detalhe = String(json?.error?.message || body || '').trim();
    throw new Error(`OpenAI falhou (${code}): ${detalhe || 'erro desconhecido'}`);
  }

  const erroStatus = montarErroStatusOpenAIResponse_(json);
  if (erroStatus) {
    throw new Error(erroStatus);
  }

  const texto = extrairTextoOpenAIResponse_(json);
  if (!texto) {
    throw new Error(`OpenAI nao retornou dados extraidos. ${resumirOpenAIResponse_(json)}`);
  }

  try {
    return JSON.parse(texto);
  } catch (error) {
    const inicio = texto.indexOf('{');
    const fim = texto.lastIndexOf('}');
    if (inicio >= 0 && fim > inicio) {
      try {
        return JSON.parse(texto.slice(inicio, fim + 1));
      } catch (innerError) {
        // continua para erro padrao
      }
    }
    throw new Error('OpenAI retornou dados em formato inesperado.');
  }
}

function getOpenAIComprovantesApiKey_() {
  const apiKey = String(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY') || '').trim();
  if (!apiKey) {
    throw new Error('OPENAI_API_KEY nao configurada nas propriedades do Apps Script.');
  }
  return apiKey;
}

function getOpenAIComprovantesModel_() {
  return String(
    PropertiesService.getScriptProperties().getProperty('OPENAI_COMPROVANTES_MODEL') ||
    OPENAI_COMPROVANTES_MODEL_DEFAULT
  ).trim() || OPENAI_COMPROVANTES_MODEL_DEFAULT;
}

function getOpenAIComprovantesReasoningEffort_() {
  const raw = String(
    PropertiesService.getScriptProperties().getProperty('OPENAI_COMPROVANTES_REASONING_EFFORT') ||
    OPENAI_COMPROVANTES_REASONING_EFFORT_DEFAULT
  ).trim().toLowerCase();
  const permitidos = ['none', 'minimal', 'low', 'medium', 'high', 'xhigh'];
  return permitidos.includes(raw) ? raw : OPENAI_COMPROVANTES_REASONING_EFFORT_DEFAULT;
}

function getOpenAIComprovantesReasoning_(model) {
  const nome = String(model || '').trim().toLowerCase();
  if (!nome) return null;
  if (nome.startsWith('gpt-5') || nome.startsWith('o')) {
    return { effort: getOpenAIComprovantesReasoningEffort_() };
  }
  return null;
}

function montarPromptOpenAIInboxDespesa_() {
  const validacoes = obterValidacoes();
  const categorias = limitarListaPromptInboxDespesa_(validacoes?.categoriasDespesas, 80);
  const pagosPor = limitarListaPromptInboxDespesa_(validacoes?.pagosPor, 40);
  const formasPagamento = limitarListaPromptInboxDespesa_(validacoes?.formasPagamento, 40);
  const hoje = formatarDataYmdFinanceiro(new Date());

  return [
    'Extraia dados de uma imagem de comprovante, nota, recibo ou tela de pagamento para cadastrar uma despesa geral.',
    'Use somente informacoes visiveis na imagem. Nao invente dados ausentes.',
    'Datas devem estar no formato yyyy-mm-dd. Valores devem ser numero decimal em reais.',
    'Se for comprovante de pagamento, use a data da transacao como data_pagamento e, se nao houver competencia melhor, tambem como data_competencia.',
    'Se uma data nao estiver visivel, deixe o campo vazio.',
    'A descricao deve ser curta e reconhecivel para uma lista financeira.',
    'Categoria, pago_por e forma_pagamento devem usar exatamente uma opcao valida quando houver correspondencia clara; caso contrario, deixe vazio.',
    `Data de hoje para contexto: ${hoje}.`,
    `Categorias validas: ${JSON.stringify(categorias)}.`,
    `Pagos por validos: ${JSON.stringify(pagosPor)}.`,
    `Formas de pagamento validas: ${JSON.stringify(formasPagamento)}.`
  ].join('\n');
}

function getSchemaOpenAIInboxDespesa_() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      descricao: { type: 'string' },
      categoria: { type: 'string' },
      fornecedor: { type: 'string' },
      pago_por: { type: 'string' },
      valor_total: { type: 'number' },
      data_competencia: { type: 'string' },
      data_vencimento: { type: 'string' },
      data_pagamento: { type: 'string' },
      forma_pagamento: { type: 'string' },
      parcelas: { type: 'integer' },
      observacao: { type: 'string' },
      confianca: { type: 'number' },
      alertas: {
        type: 'array',
        items: { type: 'string' }
      }
    },
    required: [
      'descricao',
      'categoria',
      'fornecedor',
      'pago_por',
      'valor_total',
      'data_competencia',
      'data_vencimento',
      'data_pagamento',
      'forma_pagamento',
      'parcelas',
      'observacao',
      'confianca',
      'alertas'
    ]
  };
}

function extrairTextoOpenAIResponse_(json) {
  if (typeof json?.output_text === 'string' && json.output_text.trim()) {
    return json.output_text.trim();
  }

  const partes = [];
  (Array.isArray(json?.output) ? json.output : []).forEach(item => {
    coletarTextoOpenAIResponse_(item?.content, partes, 0);
    coletarTextoOpenAIResponse_(item?.message?.content, partes, 0);
    coletarTextoOpenAIResponse_(item?.text, partes, 0);
  });
  return partes.join('\n').trim();
}

function coletarTextoOpenAIResponse_(node, partes, depth) {
  if (!node || depth > 5) return;
  if (typeof node === 'string') {
    const texto = node.trim();
    if (texto) partes.push(texto);
    return;
  }
  if (Array.isArray(node)) {
    node.forEach(item => coletarTextoOpenAIResponse_(item, partes, depth + 1));
    return;
  }
  if (typeof node !== 'object') return;

  ['text', 'output_text', 'value'].forEach(key => {
    const valor = node[key];
    if (typeof valor === 'string' && valor.trim()) {
      partes.push(valor.trim());
    } else if (valor && typeof valor === 'object') {
      coletarTextoOpenAIResponse_(valor, partes, depth + 1);
    }
  });

  if (Array.isArray(node.content)) {
    coletarTextoOpenAIResponse_(node.content, partes, depth + 1);
  }
}

function montarErroStatusOpenAIResponse_(json) {
  const status = String(json?.status || '').trim().toLowerCase();
  if (!status || status === 'completed') return '';

  if (status === 'incomplete') {
    const reason = String(json?.incomplete_details?.reason || '').trim();
    return `OpenAI retornou resposta incompleta${reason ? ` (${reason})` : ''}. Tente novamente com a imagem mais recortada ou use OPENAI_COMPROVANTES_REASONING_EFFORT=low.`;
  }

  if (status === 'failed') {
    const msg = String(json?.error?.message || json?.error || '').trim();
    return `OpenAI falhou ao gerar resposta${msg ? `: ${msg}` : '.'}`;
  }

  return `OpenAI retornou status inesperado: ${status}. ${resumirOpenAIResponse_(json)}`;
}

function resumirOpenAIResponse_(json) {
  const status = String(json?.status || '').trim() || 'sem_status';
  const reason = String(json?.incomplete_details?.reason || '').trim();
  const output = (Array.isArray(json?.output) ? json.output : [])
    .map(item => {
      const tipo = String(item?.type || 'sem_tipo').trim();
      const contentTypes = (Array.isArray(item?.content) ? item.content : [])
        .map(content => String(content?.type || 'sem_content_type').trim())
        .filter(Boolean)
        .join(',');
      return contentTypes ? `${tipo}:${contentTypes}` : tipo;
    })
    .filter(Boolean)
    .join('|');
  const tokens = json?.usage?.total_tokens ? ` tokens=${json.usage.total_tokens}` : '';
  return `status=${status}${reason ? ` reason=${reason}` : ''}${output ? ` output=${output}` : ''}${tokens}.`;
}

function normalizarDespesaExtraidaInbox_(extraido) {
  const dados = extraido || {};
  const validacoes = obterValidacoes();
  const categorias = Array.isArray(validacoes?.categoriasDespesas) ? validacoes.categoriasDespesas : [];
  const pagosPor = Array.isArray(validacoes?.pagosPor) ? validacoes.pagosPor : [];
  const formasPagamento = Array.isArray(validacoes?.formasPagamento) ? validacoes.formasPagamento : [];

  const valorTotal = round2Financeiro(parseNumeroBR(dados.valor_total));
  const dataPagamento = normalizarDataOpcionalInboxDespesa_(dados.data_pagamento);
  const dataVencimento = normalizarDataOpcionalInboxDespesa_(dados.data_vencimento);
  let dataCompetencia = normalizarDataOpcionalInboxDespesa_(dados.data_competencia);
  const alertas = Array.isArray(dados.alertas)
    ? dados.alertas.map(v => String(v || '').trim()).filter(Boolean)
    : [];

  if (!dataCompetencia && (dataPagamento || dataVencimento)) {
    dataCompetencia = dataPagamento || dataVencimento;
  }
  if (!dataCompetencia) {
    dataCompetencia = formatarDataYmdFinanceiro(new Date());
    alertas.push('Data de competencia nao identificada; usando a data atual.');
  }
  if (valorTotal <= 0) {
    alertas.push('Valor total nao identificado.');
  }

  const categoria = normalizarValorListaInboxDespesa_(dados.categoria, categorias);
  if (!categoria) {
    alertas.push('Categoria nao identificada.');
  }

  return {
    descricao: String(dados.descricao || '').trim().slice(0, 160),
    categoria,
    fornecedor: String(dados.fornecedor || '').trim().slice(0, 120),
    pago_por: normalizarValorListaInboxDespesa_(dados.pago_por, pagosPor),
    valor_total: valorTotal,
    data_competencia: dataCompetencia,
    data_vencimento: dataVencimento,
    data_pagamento: dataPagamento,
    forma_pagamento: normalizarValorListaInboxDespesa_(dados.forma_pagamento, formasPagamento),
    parcelas: Math.max(1, Math.min(60, Math.floor(parseNumeroBR(dados.parcelas) || 1))),
    observacao: String(dados.observacao || '').trim().slice(0, 500),
    confianca: Math.max(0, Math.min(1, Number(dados.confianca || 0) || 0)),
    alertas: [...new Set(alertas)]
  };
}

function montarPayloadConfirmacaoInboxDespesa_(item, payloadOverride) {
  const override = payloadOverride || {};
  return {
    descricao: String(override.descricao ?? item.descricao ?? '').trim(),
    categoria: String(override.categoria ?? item.categoria ?? '').trim(),
    fornecedor: String(override.fornecedor ?? item.fornecedor ?? '').trim(),
    pago_por: String(override.pago_por ?? item.pago_por ?? '').trim(),
    valor_total: round2Financeiro(parseNumeroBR(override.valor_total ?? item.valor_total)),
    data_competencia: String(override.data_competencia ?? item.data_competencia ?? '').trim(),
    data_vencimento: String(override.data_vencimento ?? item.data_vencimento ?? '').trim(),
    data_pagamento: String(override.data_pagamento ?? item.data_pagamento ?? '').trim(),
    forma_pagamento: String(override.forma_pagamento ?? item.forma_pagamento ?? '').trim(),
    parcelas: Math.max(1, Math.floor(parseNumeroBR(override.parcelas ?? item.parcelas) || 1)),
    parcelas_detalhe: [],
    fixo: false,
    observacao: String(override.observacao ?? item.observacao ?? '').trim(),
    client_request_id: `INBOX_DESPESA_${String(item.ID || '').trim()}`
  };
}

function normalizarDataOpcionalInboxDespesa_(valor) {
  const raw = String(valor || '').trim();
  if (!raw) return '';
  try {
    return normalizarDataFinanceiro(raw, false, 'Data');
  } catch (error) {
    return '';
  }
}

function normalizarValorListaInboxDespesa_(valor, lista) {
  const raw = String(valor || '').trim();
  if (!raw) return '';

  const opcoes = Array.isArray(lista)
    ? lista.map(v => String(v || '').trim()).filter(Boolean)
    : [];
  if (opcoes.length === 0) return raw;

  const chave = normalizarChaveInboxDespesa_(raw);
  const match = opcoes.find(v => normalizarChaveInboxDespesa_(v) === chave);
  return match || '';
}

function limitarListaPromptInboxDespesa_(lista, limite) {
  const max = Math.max(1, Math.floor(Number(limite || 40)));
  return [...new Set((Array.isArray(lista) ? lista : [])
    .map(v => String(v || '').trim())
    .filter(Boolean))]
    .slice(0, max);
}

function normalizarChaveInboxDespesa_(valor) {
  return String(valor || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
}

function parseJsonInboxDespesa_(raw, fallback) {
  if (Array.isArray(raw) || (raw && typeof raw === 'object')) return raw;
  const texto = String(raw || '').trim();
  if (!texto) return fallback;
  try {
    return JSON.parse(texto);
  } catch (error) {
    return fallback;
  }
}

function serializarListaInboxDespesa_(lista) {
  const arr = (Array.isArray(lista) ? lista : [])
    .map(v => String(v || '').trim())
    .filter(Boolean);
  return arr.length ? JSON.stringify([...new Set(arr)]) : '';
}

function hashTextoInboxDespesa_(texto) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(texto || ''),
    Utilities.Charset.UTF_8
  );
  return digest
    .map(b => {
      const n = b < 0 ? b + 256 : b;
      return n.toString(16).padStart(2, '0');
    })
    .join('');
}

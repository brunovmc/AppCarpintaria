const ABA_INBOX_RECEBIMENTOS = 'INBOX_RECEBIMENTOS';

const INBOX_RECEBIMENTOS_SCHEMA = [
  'ID',
  'status',
  'origem_tipo',
  'arquivo_nome',
  'arquivo_mime',
  'arquivo_drive_id',
  'arquivo_url',
  'arquivo_hash',
  'referencia_transacao',
  'pagador_nome',
  'banco_pagador',
  'recebido_por',
  'valor_total',
  'data_recebimento',
  'forma_pagamento',
  'descricao',
  'observacao',
  'dados_extraidos_json',
  'confianca',
  'alertas_json',
  'erro',
  'criado_em',
  'atualizado_em',
  'confirmado_em',
  'conciliado_em',
  'descartado_em'
];

function listarInboxRecebimentos(statusFiltro, ambiente) {
  return executarComAmbienteInboxRecebimentos_(ambiente, () =>
    listarInboxRecebimentosNoAmbienteAtual_(statusFiltro)
  );
}

function listarInboxRecebimentosNoAmbienteAtual_(statusFiltro) {
  const sheet = getSheet(ABA_INBOX_RECEBIMENTOS);
  if (!sheet) return [];
  ensureSchema(sheet, INBOX_RECEBIMENTOS_SCHEMA);
  const filtro = String(statusFiltro || '').trim().toUpperCase();
  return rowsToObjects(sheet)
    .map(normalizarLinhaInboxRecebimento_)
    .filter(item => {
      const status = String(item.status || '').trim().toUpperCase();
      if (!filtro) return ['PENDENTE', 'PARCIAL', 'ERRO'].includes(status);
      if (filtro === 'TODOS' || filtro === 'ALL') return true;
      return status === filtro;
    })
    .sort((a, b) => {
      const da = parseDataFinanceiro(a.atualizado_em || a.criado_em)?.getTime() || 0;
      const db = parseDataFinanceiro(b.atualizado_em || b.criado_em)?.getTime() || 0;
      return db - da || String(b.ID || '').localeCompare(String(a.ID || ''));
    });
}

function processarComprovanteRecebimentoUpload(payload) {
  const ambiente = payload?._db_env || payload?.ambiente || payload?.db_env || '';
  return executarComAmbienteInboxRecebimentos_(ambiente, () =>
    processarComprovanteRecebimentoUploadAtual_(payload)
  );
}

function processarComprovanteRecebimentoUploadAtual_(payload) {
  assertCanWrite('Inbox de recebimentos');
  const upload = normalizarUploadInboxRecebimento_(payload);
  const existente = buscarInboxRecebimentoPorHash_(upload.hash);
  if (existente) return { ...existente, reutilizado: true };

  const arquivo = salvarArquivoInboxRecebimento_(upload);
  const agora = new Date();
  const base = {
    ID: gerarId('IBXREC'),
    origem_tipo: 'UPLOAD_APP',
    arquivo_nome: upload.nome,
    arquivo_mime: upload.mime,
    arquivo_drive_id: arquivo.id || '',
    arquivo_url: arquivo.url || '',
    arquivo_hash: upload.hash,
    criado_em: agora,
    atualizado_em: agora,
    confirmado_em: '',
    conciliado_em: '',
    descartado_em: ''
  };

  try {
    const extraido = extrairRecebimentoComOpenAI_({
      dataUrl: upload.dataUrl,
      base64: upload.base64,
      mime: upload.mime,
      nome: upload.nome
    });
    const rascunho = normalizarRecebimentoExtraido_(extraido);
    const alertas = [...rascunho.alertas];
    if (arquivo.erro) alertas.push(arquivo.erro);
    if (rascunho.referencia_transacao) {
      const referenciaDuplicada = buscarInboxRecebimentoPorReferencia_(rascunho.referencia_transacao);
      if (referenciaDuplicada) {
        alertas.push(`Referencia da transacao ja encontrada no comprovante ${referenciaDuplicada.ID}.`);
      }
    }
    const linha = {
      ...base,
      status: 'PENDENTE',
      referencia_transacao: rascunho.referencia_transacao,
      pagador_nome: rascunho.pagador_nome,
      banco_pagador: rascunho.banco_pagador,
      recebido_por: rascunho.recebido_por,
      valor_total: rascunho.valor_total,
      data_recebimento: rascunho.data_recebimento,
      forma_pagamento: rascunho.forma_pagamento,
      descricao: rascunho.descricao,
      observacao: rascunho.observacao,
      dados_extraidos_json: JSON.stringify(extraido || {}),
      confianca: rascunho.confianca,
      alertas_json: serializarListaRecebimento_(alertas),
      erro: ''
    };
    if (!insert(ABA_INBOX_RECEBIMENTOS, linha, INBOX_RECEBIMENTOS_SCHEMA)) {
      throw new Error('Nao foi possivel salvar o comprovante na Inbox de recebimentos.');
    }
    const inserido = obterInboxRecebimentoPorId_(linha.ID) || normalizarLinhaInboxRecebimento_(linha);
    if (typeof moverArquivoInboxRecebimentoAposLeitura_ === 'function') {
      moverArquivoInboxRecebimentoAposLeitura_(inserido);
    }
    return inserido;
  } catch (error) {
    const linhaErro = {
      ...base,
      status: 'ERRO',
      referencia_transacao: '',
      pagador_nome: '',
      banco_pagador: '',
      recebido_por: '',
      valor_total: 0,
      data_recebimento: '',
      forma_pagamento: '',
      descricao: '',
      observacao: '',
      dados_extraidos_json: '',
      confianca: 0,
      alertas_json: serializarListaRecebimento_([arquivo.erro].filter(Boolean)),
      erro: error?.message || String(error || 'Falha ao processar comprovante.')
    };
    if (!insert(ABA_INBOX_RECEBIMENTOS, linhaErro, INBOX_RECEBIMENTOS_SCHEMA)) {
      throw new Error('Falha ao processar e ao registrar o comprovante de recebimento.');
    }
    const inserido = obterInboxRecebimentoPorId_(linhaErro.ID) || normalizarLinhaInboxRecebimento_(linhaErro);
    if (typeof moverArquivoInboxRecebimentoAposErro_ === 'function') {
      moverArquivoInboxRecebimentoAposErro_(inserido);
    }
    return inserido;
  }
}

function descartarInboxRecebimento(id, ambiente) {
  return executarComAmbienteInboxRecebimentos_(ambiente, () => {
    assertCanWrite('Descarte de comprovante de recebimento');
    const item = obterInboxRecebimentoPorId_(id);
    if (!item) throw new Error('Comprovante de recebimento nao encontrado.');
    const possuiVinculo = listarPagamentos(true).some(pagamento =>
      String(pagamento.comprovante_id || '').trim() === String(item.ID || '').trim()
    );
    if (possuiVinculo) throw new Error('Desfaca os vinculos antes de descartar o comprovante.');
    updateById(ABA_INBOX_RECEBIMENTOS, 'ID', item.ID, {
      status: 'DESCARTADO',
      atualizado_em: new Date(),
      descartado_em: new Date()
    }, INBOX_RECEBIMENTOS_SCHEMA);
    if (typeof moverArquivoInboxRecebimentoAposDescarte_ === 'function') {
      moverArquivoInboxRecebimentoAposDescarte_(item);
    }
    return { ok: true };
  });
}

function obterInboxRecebimentoPorId_(id) {
  const alvo = String(id || '').trim();
  const sheet = getSheet(ABA_INBOX_RECEBIMENTOS);
  if (!alvo || !sheet) return null;
  ensureSchema(sheet, INBOX_RECEBIMENTOS_SCHEMA);
  const item = rowsToObjects(sheet).find(row => String(row.ID || '').trim() === alvo);
  return item ? normalizarLinhaInboxRecebimento_(item) : null;
}

function buscarInboxRecebimentoPorHash_(hash) {
  const alvo = String(hash || '').trim().toLowerCase();
  const sheet = getSheet(ABA_INBOX_RECEBIMENTOS);
  if (!alvo || !sheet) return null;
  ensureSchema(sheet, INBOX_RECEBIMENTOS_SCHEMA);
  const item = rowsToObjects(sheet).find(row =>
    String(row.arquivo_hash || '').trim().toLowerCase() === alvo &&
    String(row.status || '').trim().toUpperCase() !== 'DESCARTADO'
  );
  return item ? normalizarLinhaInboxRecebimento_(item) : null;
}

function buscarInboxRecebimentoPorReferencia_(referencia) {
  const alvo = normalizarTextoRecebimento_(referencia);
  const sheet = getSheet(ABA_INBOX_RECEBIMENTOS);
  if (!alvo || !sheet) return null;
  ensureSchema(sheet, INBOX_RECEBIMENTOS_SCHEMA);
  const item = rowsToObjects(sheet).find(row =>
    normalizarTextoRecebimento_(row.referencia_transacao) === alvo &&
    String(row.status || '').trim().toUpperCase() !== 'DESCARTADO'
  );
  return item ? normalizarLinhaInboxRecebimento_(item) : null;
}

function normalizarLinhaInboxRecebimento_(row) {
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
    arquivo_hash: String(item.arquivo_hash || '').trim(),
    referencia_transacao: String(item.referencia_transacao || '').trim(),
    pagador_nome: String(item.pagador_nome || '').trim(),
    banco_pagador: String(item.banco_pagador || '').trim(),
    recebido_por: String(item.recebido_por || '').trim(),
    valor_total: round2Financeiro(parseNumeroBR(item.valor_total)),
    data_recebimento: formatarDataYmdFinanceiroSafe(item.data_recebimento),
    forma_pagamento: String(item.forma_pagamento || '').trim(),
    descricao: String(item.descricao || '').trim(),
    observacao: String(item.observacao || '').trim(),
    confianca: Math.max(0, Math.min(1, Number(item.confianca || 0) || 0)),
    alertas: parseJsonRecebimento_(item.alertas_json, []),
    erro: String(item.erro || '').trim(),
    criado_em: item.criado_em ? formatarDataYmdHmFinanceiroSafe(item.criado_em) : '',
    atualizado_em: item.atualizado_em ? formatarDataYmdHmFinanceiroSafe(item.atualizado_em) : '',
    confirmado_em: item.confirmado_em ? formatarDataYmdHmFinanceiroSafe(item.confirmado_em) : '',
    conciliado_em: item.conciliado_em ? formatarDataYmdHmFinanceiroSafe(item.conciliado_em) : '',
    descartado_em: item.descartado_em ? formatarDataYmdHmFinanceiroSafe(item.descartado_em) : ''
  };
}

function normalizarUploadInboxRecebimento_(payload) {
  const dados = payload || {};
  const dataUrl = String(dados.data_url || dados.dataUrl || '').trim();
  const match = /^data:([^;]+);base64,(.+)$/i.exec(dataUrl);
  if (!match) throw new Error('Arquivo invalido. Envie uma imagem ou PDF.');
  const mime = String(dados.mime_type || dados.mime || match[1] || '').trim().toLowerCase();
  if (!(mime.startsWith('image/') || mime === 'application/pdf')) {
    throw new Error('Formato nao suportado. Use imagem ou PDF.');
  }
  const base64 = String(match[2] || '').trim();
  if (!base64) throw new Error('Arquivo sem conteudo.');
  if (base64.length > 9 * 1024 * 1024) {
    throw new Error('Arquivo muito grande. Use uma imagem menor ou um PDF mais leve.');
  }
  let bytes;
  try {
    bytes = Utilities.base64Decode(base64);
  } catch (error) {
    throw new Error('Arquivo em Base64 invalido.');
  }
  const nome = String(dados.nome || dados.name || `recebimento-${Date.now()}`)
    .replace(/[\\/:*?"<>|]+/g, '-')
    .trim()
    .slice(0, 180) || `recebimento-${Date.now()}`;
  return { nome, mime, base64, bytes, dataUrl, hash: hashBytesRecebimento_(bytes) };
}

function salvarArquivoInboxRecebimento_(upload) {
  try {
    if (typeof salvarArquivoUploadInboxRecebimentosDrive_ === 'function') {
      const salvo = salvarArquivoUploadInboxRecebimentosDrive_(upload);
      return { id: salvo.id || '', url: salvo.url || '', erro: '' };
    }
    return { id: '', url: '', erro: '' };
  } catch (error) {
    return { id: '', url: '', erro: `Arquivo nao salvo no Drive: ${error?.message || error}` };
  }
}

function extrairRecebimentoComOpenAI_(arquivoEntrada) {
  const apiKey = getOpenAIComprovantesApiKey_();
  const model = getOpenAIComprovantesModel_();
  const reasoning = getOpenAIComprovantesReasoning_(model);
  const arquivo = normalizarArquivoOpenAIInboxDespesa_(arquivoEntrada);
  const conteudoArquivo = arquivo.mime === 'application/pdf'
    ? {
        type: 'input_file',
        filename: arquivo.nome || 'recebimento.pdf',
        file_data: `data:${arquivo.mime};base64,${arquivo.base64}`
      }
    : { type: 'input_image', image_url: arquivo.dataUrl, detail: 'high' };
  const payload = {
    model,
    input: [{
      role: 'user',
      content: [
        { type: 'input_text', text: montarPromptOpenAIInboxRecebimento_() },
        conteudoArquivo
      ]
    }],
    text: {
      format: {
        type: 'json_schema',
        name: 'comprovante_recebimento',
        strict: true,
        schema: getSchemaOpenAIInboxRecebimento_()
      }
    },
    max_output_tokens: 2200
  };
  if (reasoning) payload.reasoning = reasoning;
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${apiKey}` },
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  });
  const code = response.getResponseCode();
  const body = response.getContentText();
  let json;
  try { json = JSON.parse(body); } catch (error) {
    throw new Error(`Resposta invalida da OpenAI (${code}).`);
  }
  if (code < 200 || code >= 300) {
    throw new Error(`OpenAI falhou (${code}): ${String(json?.error?.message || body || 'erro desconhecido').trim()}`);
  }
  const erroStatus = montarErroStatusOpenAIResponse_(json);
  if (erroStatus) throw new Error(erroStatus);
  const texto = extrairTextoOpenAIResponse_(json);
  if (!texto) throw new Error(`OpenAI nao retornou dados extraidos. ${resumirOpenAIResponse_(json)}`);
  try {
    return JSON.parse(texto);
  } catch (error) {
    throw new Error('OpenAI retornou dados em formato inesperado.');
  }
}

function montarPromptOpenAIInboxRecebimento_() {
  const validacoes = obterValidacoes();
  const recebidosPor = limitarListaPromptInboxDespesa_(validacoes?.pagosPor, 40);
  const formas = limitarListaPromptInboxDespesa_(validacoes?.formasPagamento, 40);
  return [
    'Extraia os dados visiveis de um comprovante de dinheiro recebido, como PIX, TED, deposito ou transferencia.',
    'Use somente informacoes visiveis no arquivo. Nao invente cliente, venda, produto, parcela ou identificadores internos.',
    'A data deve usar yyyy-mm-dd e o valor deve ser um numero decimal em reais.',
    'pagador_nome e a pessoa ou empresa que enviou o dinheiro.',
    'referencia_transacao deve conter EndToEndId, NSU, autenticacao ou codigo equivalente, quando visivel.',
    'recebido_por e forma_pagamento devem usar exatamente uma opcao valida apenas quando houver correspondencia clara; caso contrario use string vazia.',
    `Recebidos por validos: ${JSON.stringify(recebidosPor)}.`,
    `Formas de pagamento validas: ${JSON.stringify(formas)}.`,
    `Data de hoje para contexto: ${formatarDataYmdFinanceiro(new Date())}.`
  ].join('\n');
}

function getSchemaOpenAIInboxRecebimento_() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      referencia_transacao: { type: 'string' },
      pagador_nome: { type: 'string' },
      banco_pagador: { type: 'string' },
      recebido_por: { type: 'string' },
      valor_total: { type: 'number' },
      data_recebimento: { type: 'string' },
      forma_pagamento: { type: 'string' },
      descricao: { type: 'string' },
      observacao: { type: 'string' },
      confianca: { type: 'number' },
      alertas: { type: 'array', items: { type: 'string' } }
    },
    required: [
      'referencia_transacao', 'pagador_nome', 'banco_pagador', 'recebido_por',
      'valor_total', 'data_recebimento', 'forma_pagamento', 'descricao',
      'observacao', 'confianca', 'alertas'
    ]
  };
}

function normalizarRecebimentoExtraido_(extraido) {
  const dados = extraido || {};
  const validacoes = obterValidacoes();
  const alertas = Array.isArray(dados.alertas)
    ? dados.alertas.map(item => String(item || '').trim()).filter(Boolean)
    : [];
  const valorTotal = round2Financeiro(parseNumeroBR(dados.valor_total));
  const dataRecebimento = normalizarDataOpcionalRecebimento_(dados.data_recebimento);
  const recebidoPor = normalizarValorListaRecebimento_(dados.recebido_por, validacoes?.pagosPor);
  const formaPagamento = normalizarValorListaRecebimento_(dados.forma_pagamento, validacoes?.formasPagamento);
  if (valorTotal <= 0) alertas.push('Valor recebido nao identificado.');
  if (!dataRecebimento) alertas.push('Data do recebimento nao identificada.');
  if (!recebidoPor) alertas.push('Conta ou pessoa que recebeu nao identificada.');
  if (!formaPagamento) alertas.push('Forma de recebimento nao identificada.');
  return {
    referencia_transacao: String(dados.referencia_transacao || '').trim().slice(0, 160),
    pagador_nome: String(dados.pagador_nome || '').trim().slice(0, 160),
    banco_pagador: String(dados.banco_pagador || '').trim().slice(0, 120),
    recebido_por: recebidoPor,
    valor_total: valorTotal,
    data_recebimento: dataRecebimento,
    forma_pagamento: formaPagamento,
    descricao: String(dados.descricao || '').trim().slice(0, 200),
    observacao: String(dados.observacao || '').trim().slice(0, 500),
    confianca: Math.max(0, Math.min(1, Number(dados.confianca || 0) || 0)),
    alertas: [...new Set(alertas)]
  };
}

function normalizarDataOpcionalRecebimento_(valor) {
  const raw = String(valor || '').trim();
  if (!raw) return '';
  try { return normalizarDataFinanceiro(raw, false, 'Data de recebimento'); } catch (error) { return ''; }
}

function normalizarValorListaRecebimento_(valor, lista) {
  const raw = String(valor || '').trim();
  if (!raw || !Array.isArray(lista)) return '';
  const normalizado = normalizarTextoRecebimento_(raw);
  return String(lista.find(item => normalizarTextoRecebimento_(item) === normalizado) || '').trim();
}

function normalizarTextoRecebimento_(valor) {
  return String(valor || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function executarComAmbienteInboxRecebimentos_(ambiente, callback) {
  const raw = String(ambiente || '').trim().toLowerCase();
  const env = raw === 'dev' || raw === 'prod' ? raw : '';
  if (env && typeof setUserDbEnvironment_ === 'function' && typeof getUserDbEnvironment_ === 'function') {
    if (String(getUserDbEnvironment_() || '').trim().toLowerCase() !== env) setUserDbEnvironment_(env);
  }
  return callback();
}

function parseJsonRecebimento_(raw, fallback) {
  if (Array.isArray(raw) || (raw && typeof raw === 'object')) return raw;
  const texto = String(raw || '').trim();
  if (!texto) return fallback;
  try { return JSON.parse(texto); } catch (error) { return fallback; }
}

function serializarListaRecebimento_(lista) {
  const itens = (Array.isArray(lista) ? lista : []).map(item => String(item || '').trim()).filter(Boolean);
  return itens.length ? JSON.stringify([...new Set(itens)]) : '';
}

function hashBytesRecebimento_(bytes) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes || [])
    .map(byte => (`0${(byte & 255).toString(16)}`).slice(-2)).join('');
}

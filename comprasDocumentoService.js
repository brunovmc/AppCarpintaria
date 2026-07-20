const ABA_DOCUMENTOS_COMPRA = 'DOCUMENTOS_COMPRA';
const ABA_DOCUMENTOS_COMPRA_ITENS = 'DOCUMENTOS_COMPRA_ITENS';
const ABA_MOVIMENTOS_ESTOQUE = 'MOVIMENTOS_ESTOQUE';
const DOCUMENTOS_COMPRA_MAX_ARQUIVOS = 5;
const DOCUMENTOS_COMPRA_MAX_BYTES_ARQUIVO = 10 * 1024 * 1024;
const DOCUMENTOS_COMPRA_MAX_BYTES_TOTAL = 25 * 1024 * 1024;

const DOCUMENTOS_COMPRA_SCHEMA = [
  'ID',
  'status',
  'origem_entrada',
  'arquivo_nomes_json',
  'arquivo_mimes_json',
  'arquivo_drive_ids_json',
  'arquivo_urls_json',
  'arquivo_hash',
  'tipo_documento',
  'fornecedor',
  'numero_documento',
  'numero_pedido',
  'data_compra',
  'data_vencimento',
  'subtotal',
  'frete',
  'desconto',
  'valor_total',
  'moeda',
  'pago_por',
  'forma_pagamento',
  'parcelas',
  'recebido',
  'pagamento_detectado',
  'pagamento_confirmado',
  'data_pagamento',
  'valor_pago_confirmado',
  'observacao',
  'dados_extraidos_json',
  'confianca',
  'alertas_json',
  'erro',
  'compra_id_financeira',
  'pagamento_id_confirmado',
  'client_request_id',
  'ativo',
  'criado_em',
  'atualizado_em',
  'confirmado_em',
  'descartado_em'
];

const DOCUMENTOS_COMPRA_ITENS_SCHEMA = [
  'ID',
  'documento_id',
  'ordem',
  'descricao_original',
  'sku',
  'item_nome',
  'tipo',
  'categoria',
  'unidade',
  'quantidade',
  'comprimento_cm',
  'largura_cm',
  'espessura_cm',
  'valor_unitario',
  'valor_total',
  'frete_rateado',
  'desconto_rateado',
  'custo_total',
  'custo_unitario_estoque',
  'destino',
  'estoque_id_alvo',
  'estoque_id_resultante',
  'recebido',
  'confianca',
  'alertas_json',
  'ativo',
  'criado_em',
  'atualizado_em'
];

const MOVIMENTOS_ESTOQUE_SCHEMA = [
  'ID',
  'status',
  'estoque_id',
  'tipo_movimento',
  'origem_tipo',
  'origem_id',
  'origem_item_id',
  'quantidade',
  'custo_unitario',
  'custo_total',
  'saldo_anterior',
  'saldo_posterior',
  'observacao',
  'ativo',
  'criado_em',
  'aplicado_em'
];

const DESTINOS_ITEM_DOCUMENTO_COMPRA = ['ESTOQUE', 'CONSUMO', 'EQUIPAMENTO', 'IGNORAR'];

function executarComAmbienteDocumentosCompra_(ambiente, callback) {
  if (typeof executarComAmbienteInboxDespesas_ === 'function') {
    return executarComAmbienteInboxDespesas_(ambiente, callback);
  }
  return callback();
}

function listarInboxDocumentosCompra(statusFiltro, ambiente) {
  return executarComAmbienteDocumentosCompra_(ambiente, () => listarInboxDocumentosCompraAtual_(statusFiltro));
}

function listarInboxDocumentosCompraAtual_(statusFiltro) {
  const sheet = getSheet(ABA_DOCUMENTOS_COMPRA);
  if (!sheet) return [];
  ensureSchema(sheet, DOCUMENTOS_COMPRA_SCHEMA);
  const filtro = String(statusFiltro || '').trim().toUpperCase();
  const visiveis = ['REVISAR', 'ERRO', 'ERRO_CONFIRMACAO', 'CONFIRMADO_AGUARDANDO_RECEBIMENTO'];
  const itensPorDocumento = {};
  listarTodosItensDocumentosCompra_().forEach(item => {
    if (!Array.isArray(itensPorDocumento[item.documento_id])) itensPorDocumento[item.documento_id] = [];
    itensPorDocumento[item.documento_id].push(item);
  });
  return rowsToObjects(sheet)
    .filter(row => String(row.ativo).toLowerCase() === 'true')
    .map(normalizarLinhaDocumentoCompra_)
    .filter(item => filtro === 'TODOS' || filtro === 'ALL' || (filtro ? item.status === filtro : visiveis.includes(item.status)))
    .map(item => ({ ...item, itens: itensPorDocumento[item.ID] || [] }))
    .sort((a, b) => {
      const da = parseDataFinanceiro(a.atualizado_em || a.criado_em)?.getTime() || 0;
      const db = parseDataFinanceiro(b.atualizado_em || b.criado_em)?.getTime() || 0;
      return db - da || String(b.ID).localeCompare(String(a.ID));
    });
}

function obterContextoInboxDocumentosCompra(ambiente) {
  return executarComAmbienteDocumentosCompra_(ambiente, () => {
    const validacoes = obterValidacoes();
    const estoque = listarEstoque(true).map(item => ({
      ID: String(item.ID || '').trim(),
      tipo: String(item.tipo || '').trim(),
      categoria: String(item.categoria || '').trim(),
      item: String(item.item || '').trim(),
      unidade: String(item.unidade || '').trim(),
      quantidade: parseNumeroBR(item.quantidade),
      custo_unitario: parseNumeroBR(item.custo_unitario || item.valor_unit),
      comprimento_cm: parseNumeroBR(item.comprimento_cm),
      largura_cm: parseNumeroBR(item.largura_cm),
      espessura_cm: parseNumeroBR(item.espessura_cm)
    }));
    let drive = { configurado: false };
    if (typeof obterStatusDocumentosCompraDrive === 'function') {
      try {
        drive = obterStatusDocumentosCompraDrive();
      } catch (error) {
        drive = { configurado: false, erro: error.message || String(error) };
      }
    }
    return {
      documentos: listarInboxDocumentosCompraAtual_(),
      validacoes: {
        tipos: validacoes?.tipos || [],
        unidades: validacoes?.unidades || [],
        categorias: validacoes?.categorias || [],
        categoriasPorTipo: validacoes?.categoriasPorTipo || {},
        fornecedores: validacoes?.fornecedores || [],
        pagosPor: validacoes?.pagosPor || [],
        formasPagamento: validacoes?.formasPagamento || []
      },
      estoque,
      drive
    };
  });
}

function obterDocumentoCompraPorId_(id) {
  const alvo = String(id || '').trim();
  const sheet = getSheet(ABA_DOCUMENTOS_COMPRA);
  if (!alvo || !sheet) return null;
  ensureSchema(sheet, DOCUMENTOS_COMPRA_SCHEMA);
  const row = rowsToObjects(sheet).find(item => String(item.ID || '').trim() === alvo);
  return row ? normalizarLinhaDocumentoCompra_(row) : null;
}

function listarItensDocumentoCompra_(documentoId) {
  const alvo = String(documentoId || '').trim();
  if (!alvo) return [];
  return listarTodosItensDocumentosCompra_().filter(item => item.documento_id === alvo);
}

function listarTodosItensDocumentosCompra_() {
  const sheet = getSheet(ABA_DOCUMENTOS_COMPRA_ITENS);
  if (!sheet) return [];
  ensureSchema(sheet, DOCUMENTOS_COMPRA_ITENS_SCHEMA);
  const movimentosPorItem = {};
  listarMovimentosEstoqueDocumentoCompra_().forEach(movimento => {
    movimentosPorItem[String(movimento.origem_item_id || '').trim()] = movimento;
  });
  return rowsToObjects(sheet)
    .filter(item => String(item.ativo).toLowerCase() === 'true')
    .map(normalizarLinhaItemDocumentoCompra_)
    .map(item => {
      const movimento = movimentosPorItem[item.ID];
      if (!movimento) return item;
      return {
        ...item,
        movimento_id: String(movimento.ID || '').trim(),
        movimento_status: String(movimento.status || '').trim().toUpperCase(),
        estoque_id_resultante: item.estoque_id_resultante || String(movimento.estoque_id || '').trim()
      };
    })
    .sort((a, b) => Number(a.ordem || 0) - Number(b.ordem || 0));
}

function normalizarLinhaDocumentoCompra_(row) {
  const item = { ...(row || {}) };
  return {
    ...item,
    ID: String(item.ID || '').trim(),
    status: String(item.status || 'REVISAR').trim().toUpperCase(),
    origem_entrada: String(item.origem_entrada || '').trim(),
    arquivos_nomes: parseJsonDocumentoCompra_(item.arquivo_nomes_json, []),
    arquivos_mimes: parseJsonDocumentoCompra_(item.arquivo_mimes_json, []),
    arquivos_drive_ids: parseJsonDocumentoCompra_(item.arquivo_drive_ids_json, []),
    arquivos_urls: parseJsonDocumentoCompra_(item.arquivo_urls_json, []),
    arquivo_hash: String(item.arquivo_hash || '').trim(),
    tipo_documento: String(item.tipo_documento || '').trim(),
    fornecedor: String(item.fornecedor || '').trim(),
    numero_documento: String(item.numero_documento || '').trim(),
    numero_pedido: String(item.numero_pedido || '').trim(),
    data_compra: formatarDataYmdFinanceiroSafe(item.data_compra),
    data_vencimento: formatarDataYmdFinanceiroSafe(item.data_vencimento),
    subtotal: round2Financeiro(parseNumeroBR(item.subtotal)),
    frete: round2Financeiro(parseNumeroBR(item.frete)),
    desconto: round2Financeiro(parseNumeroBR(item.desconto)),
    valor_total: round2Financeiro(parseNumeroBR(item.valor_total)),
    moeda: String(item.moeda || 'BRL').trim() || 'BRL',
    pago_por: String(item.pago_por || '').trim(),
    forma_pagamento: String(item.forma_pagamento || '').trim(),
    parcelas: Math.max(1, Math.floor(parseNumeroBR(item.parcelas) || 1)),
    recebido: parseBooleanFinanceiro(item.recebido),
    pagamento_detectado: parseBooleanFinanceiro(item.pagamento_detectado),
    pagamento_confirmado: parseBooleanFinanceiro(item.pagamento_confirmado),
    data_pagamento: formatarDataYmdFinanceiroSafe(item.data_pagamento),
    valor_pago_confirmado: round2Financeiro(parseNumeroBR(item.valor_pago_confirmado)),
    observacao: String(item.observacao || '').trim(),
    dados_extraidos: parseJsonDocumentoCompra_(item.dados_extraidos_json, {}),
    confianca: Math.max(0, Math.min(1, Number(item.confianca || 0) || 0)),
    alertas: parseJsonDocumentoCompra_(item.alertas_json, []),
    erro: String(item.erro || '').trim(),
    compra_id_financeira: String(item.compra_id_financeira || '').trim(),
    pagamento_id_confirmado: String(item.pagamento_id_confirmado || '').trim(),
    client_request_id: String(item.client_request_id || '').trim(),
    criado_em: formatarDataYmdHmFinanceiroSafe(item.criado_em),
    atualizado_em: formatarDataYmdHmFinanceiroSafe(item.atualizado_em),
    confirmado_em: formatarDataYmdHmFinanceiroSafe(item.confirmado_em),
    descartado_em: formatarDataYmdHmFinanceiroSafe(item.descartado_em)
  };
}

function normalizarLinhaItemDocumentoCompra_(row) {
  const item = { ...(row || {}) };
  return {
    ...item,
    ID: String(item.ID || '').trim(),
    documento_id: String(item.documento_id || '').trim(),
    ordem: Math.max(1, Math.floor(parseNumeroBR(item.ordem) || 1)),
    descricao_original: String(item.descricao_original || '').trim(),
    sku: String(item.sku || '').trim(),
    item_nome: String(item.item_nome || '').trim(),
    tipo: String(item.tipo || '').trim(),
    categoria: String(item.categoria || '').trim(),
    unidade: String(item.unidade || '').trim(),
    quantidade: parseNumeroBR(item.quantidade),
    comprimento_cm: parseNumeroBR(item.comprimento_cm),
    largura_cm: parseNumeroBR(item.largura_cm),
    espessura_cm: parseNumeroBR(item.espessura_cm),
    valor_unitario: round2Financeiro(parseNumeroBR(item.valor_unitario)),
    valor_total: round2Financeiro(parseNumeroBR(item.valor_total)),
    frete_rateado: round2Financeiro(parseNumeroBR(item.frete_rateado)),
    desconto_rateado: round2Financeiro(parseNumeroBR(item.desconto_rateado)),
    custo_total: round2Financeiro(parseNumeroBR(item.custo_total)),
    custo_unitario_estoque: parseNumeroBR(item.custo_unitario_estoque),
    destino: normalizarDestinoItemDocumentoCompra_(item.destino),
    estoque_id_alvo: String(item.estoque_id_alvo || '').trim(),
    estoque_id_resultante: String(item.estoque_id_resultante || '').trim(),
    recebido: parseBooleanFinanceiro(item.recebido),
    confianca: Math.max(0, Math.min(1, Number(item.confianca || 0) || 0)),
    alertas: parseJsonDocumentoCompra_(item.alertas_json, []),
    criado_em: formatarDataYmdHmFinanceiroSafe(item.criado_em),
    atualizado_em: formatarDataYmdHmFinanceiroSafe(item.atualizado_em)
  };
}

function parseJsonDocumentoCompra_(raw, fallback) {
  if (Array.isArray(raw) || (raw && typeof raw === 'object')) return raw;
  try {
    const parsed = JSON.parse(String(raw || ''));
    return parsed === null || parsed === undefined ? fallback : parsed;
  } catch (error) {
    return fallback;
  }
}

function normalizarDestinoItemDocumentoCompra_(valor) {
  const destino = String(valor || '').trim().toUpperCase();
  return DESTINOS_ITEM_DOCUMENTO_COMPRA.includes(destino) ? destino : 'ESTOQUE';
}

function processarDocumentoCompraUpload(payload) {
  const ambiente = payload?.ambiente || payload?._db_env || payload?.db_env || '';
  return executarComAmbienteDocumentosCompra_(ambiente, () => processarDocumentoCompraUploadAtual_(payload));
}

function processarDocumentoCompraUploadAtual_(payload) {
  assertCanWrite('Leitura de documento de compra');
  const arquivos = normalizarArquivosUploadDocumentoCompra_(payload);
  const hash = calcularHashCompostoDocumentoCompra_(arquivos);
  const existente = buscarDocumentoCompraPorHash_(hash);
  if (existente && existente.status !== 'DESCARTADO') {
    return { ...existente, itens: listarItensDocumentoCompra_(existente.ID), duplicado: true };
  }

  const salvos = typeof salvarArquivosUploadDocumentoCompraDrive_ === 'function'
    ? salvarArquivosUploadDocumentoCompraDrive_(arquivos)
    : arquivos.map(() => ({ id: '', url: '' }));
  const documento = criarRascunhoDocumentoCompraPorArquivos_(arquivos, {
    origemEntrada: 'UPLOAD_APP',
    hash,
    driveIds: salvos.map(item => item.id || ''),
    urls: salvos.map(item => item.url || '')
  });
  if (documento.status === 'ERRO' && typeof moverArquivosDocumentoCompraAposErro_ === 'function') {
    moverArquivosDocumentoCompraAposErro_(documento);
  } else if (documento.status !== 'ERRO' && typeof moverArquivosDocumentoCompraAposLeitura_ === 'function') {
    moverArquivosDocumentoCompraAposLeitura_(documento);
  }
  return documento;
}

function criarRascunhoDocumentoCompraPorArquivos_(arquivos, opcoes) {
  const opts = opcoes || {};
  const agora = new Date();
  const id = gerarId('DOCCOM');
  const hash = String(opts.hash || calcularHashCompostoDocumentoCompra_(arquivos)).trim();
  const base = {
    ID: id,
    status: 'ERRO',
    origem_entrada: String(opts.origemEntrada || 'UPLOAD_APP').trim(),
    arquivo_nomes_json: JSON.stringify(arquivos.map(item => item.nome)),
    arquivo_mimes_json: JSON.stringify(arquivos.map(item => item.mime)),
    arquivo_drive_ids_json: JSON.stringify(opts.driveIds || []),
    arquivo_urls_json: JSON.stringify(opts.urls || []),
    arquivo_hash: hash,
    client_request_id: `DOCUMENTO_COMPRA_${id}`,
    ativo: true,
    criado_em: agora,
    atualizado_em: agora,
    confirmado_em: '',
    descartado_em: '',
    compra_id_financeira: '',
    pagamento_id_confirmado: ''
  };

  try {
    const extraido = extrairDocumentoCompraComOpenAI_(arquivos);
    const normalizado = normalizarDocumentoCompraExtraido_(extraido);
    const documento = {
      ...base,
      ...normalizado.documento,
      status: 'REVISAR',
      dados_extraidos_json: JSON.stringify(extraido || {}),
      alertas_json: JSON.stringify(normalizado.alertas),
      erro: ''
    };
    insert(ABA_DOCUMENTOS_COMPRA, documento, DOCUMENTOS_COMPRA_SCHEMA);
    const itens = normalizado.itens.map((item, index) => ({
      ...item,
      ID: gerarId('DOCIT'),
      documento_id: id,
      ordem: index + 1,
      frete_rateado: 0,
      desconto_rateado: 0,
      custo_total: item.valor_total,
      custo_unitario_estoque: 0,
      estoque_id_resultante: '',
      alertas_json: JSON.stringify(item.alertas || []),
      ativo: true,
      criado_em: agora,
      atualizado_em: agora
    }));
    insertMany(ABA_DOCUMENTOS_COMPRA_ITENS, itens, DOCUMENTOS_COMPRA_ITENS_SCHEMA);
    return { ...normalizarLinhaDocumentoCompra_(documento), itens: itens.map(normalizarLinhaItemDocumentoCompra_) };
  } catch (error) {
    const documentoErro = {
      ...base,
      tipo_documento: '', fornecedor: '', numero_documento: '', numero_pedido: '',
      data_compra: '', data_vencimento: '', subtotal: 0, frete: 0, desconto: 0,
      valor_total: 0, moeda: 'BRL', pago_por: '', forma_pagamento: '', parcelas: 1,
      recebido: false, pagamento_detectado: false, pagamento_confirmado: false,
      data_pagamento: '', valor_pago_confirmado: 0, observacao: '',
      dados_extraidos_json: '', confianca: 0, alertas_json: '[]',
      erro: error?.message || String(error || 'Falha ao ler documento de compra.')
    };
    insert(ABA_DOCUMENTOS_COMPRA, documentoErro, DOCUMENTOS_COMPRA_SCHEMA);
    return { ...normalizarLinhaDocumentoCompra_(documentoErro), itens: [] };
  }
}

function normalizarArquivosUploadDocumentoCompra_(payload) {
  const dados = payload || {};
  const lista = Array.isArray(dados.arquivos) && dados.arquivos.length > 0 ? dados.arquivos : [dados];
  if (lista.length > DOCUMENTOS_COMPRA_MAX_ARQUIVOS) {
    throw new Error(`Envie no maximo ${DOCUMENTOS_COMPRA_MAX_ARQUIVOS} arquivos por documento.`);
  }
  let totalBytes = 0;
  const arquivos = lista.map((item, index) => {
    const dataUrl = String(item?.data_url || item?.dataUrl || '').trim();
    const match = /^data:([^;]+);base64,(.+)$/i.exec(dataUrl);
    if (!match) throw new Error(`Arquivo ${index + 1} invalido.`);
    const mime = String(item?.mime_type || item?.mime || match[1] || '').trim().toLowerCase();
    if (!(mime.startsWith('image/') || mime === 'application/pdf')) {
      throw new Error(`Formato do arquivo ${index + 1} nao suportado. Use imagem ou PDF.`);
    }
    const base64 = String(match[2] || '').trim();
    let bytes;
    try {
      bytes = Utilities.base64Decode(base64);
    } catch (error) {
      throw new Error(`Conteudo do arquivo ${index + 1} invalido.`);
    }
    if (bytes.length > DOCUMENTOS_COMPRA_MAX_BYTES_ARQUIVO) {
      throw new Error(`Arquivo ${index + 1} maior que 10 MB.`);
    }
    totalBytes += bytes.length;
    const nome = normalizarNomeArquivoDocumentoCompra_(item?.nome || item?.name || `documento-${index + 1}`);
    return { nome, mime, base64, bytes, dataUrl, hash: calcularHashBytesDocumentoCompra_(bytes) };
  });
  if (totalBytes > DOCUMENTOS_COMPRA_MAX_BYTES_TOTAL) {
    throw new Error('O conjunto de arquivos excede 25 MB. Divida o documento em menos imagens.');
  }
  return arquivos;
}

function normalizarNomeArquivoDocumentoCompra_(valor) {
  const nome = String(valor || '').replace(/[\\/:*?"<>|]+/g, '-').trim();
  return (nome || `documento-${Date.now()}`).slice(0, 180);
}

function calcularHashBytesDocumentoCompra_(bytes) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)
    .map(valor => (`0${(valor & 255).toString(16)}`).slice(-2))
    .join('');
}

function calcularHashCompostoDocumentoCompra_(arquivos) {
  const texto = (arquivos || [])
    .map(item => String(item.hash || calcularHashBytesDocumentoCompra_(item.bytes)))
    .sort()
    .join('|');
  const bytes = Utilities.newBlob(texto).getBytes();
  return calcularHashBytesDocumentoCompra_(bytes);
}

function buscarDocumentoCompraPorHash_(hash) {
  const alvo = String(hash || '').trim();
  const sheet = getSheet(ABA_DOCUMENTOS_COMPRA);
  if (!alvo || !sheet) return null;
  ensureSchema(sheet, DOCUMENTOS_COMPRA_SCHEMA);
  const row = rowsToObjects(sheet).find(item => String(item.arquivo_hash || '').trim() === alvo);
  return row ? normalizarLinhaDocumentoCompra_(row) : null;
}

function extrairDocumentoCompraComOpenAI_(arquivos) {
  const apiKey = getOpenAIComprovantesApiKey_();
  const model = getOpenAIComprovantesModel_();
  const reasoning = getOpenAIComprovantesReasoning_(model);
  const conteudo = [{ type: 'input_text', text: montarPromptOpenAIDocumentoCompra_() }];
  (arquivos || []).forEach(arquivo => {
    if (arquivo.mime === 'application/pdf') {
      conteudo.push({
        type: 'input_file',
        filename: arquivo.nome || 'documento.pdf',
        file_data: `data:${arquivo.mime};base64,${arquivo.base64}`
      });
    } else {
      conteudo.push({ type: 'input_image', image_url: arquivo.dataUrl, detail: 'high' });
    }
  });

  const payload = {
    model,
    input: [{ role: 'user', content: conteudo }],
    text: {
      format: {
        type: 'json_schema',
        name: 'documento_compra_multiplos_itens',
        strict: true,
        schema: getSchemaOpenAIDocumentoCompra_()
      }
    },
    max_output_tokens: 8000
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
  if (erroStatus) throw new Error(erroStatus);
  const texto = extrairTextoOpenAIResponse_(json);
  if (!texto) throw new Error(`OpenAI nao retornou os itens da compra. ${resumirOpenAIResponse_(json)}`);
  try {
    return JSON.parse(texto);
  } catch (error) {
    throw new Error('OpenAI retornou o documento de compra em formato inesperado.');
  }
}

function montarPromptOpenAIDocumentoCompra_() {
  const validacoes = obterValidacoes();
  const tipos = limitarListaPromptDocumentoCompra_(validacoes?.tipos, 60);
  const unidades = limitarListaPromptDocumentoCompra_(validacoes?.unidades, 60);
  const categoriasPorTipo = validacoes?.categoriasPorTipo || {};
  const formas = limitarListaPromptDocumentoCompra_(validacoes?.formasPagamento, 40);
  return [
    'Leia todos os arquivos como paginas ou capturas do mesmo documento de compra.',
    'O documento pode ser nota fiscal, cupom, recibo, pedido de e-commerce ou tela do Mercado Livre.',
    'Extraia o cabecalho e TODOS os itens visiveis. Nao invente dados ausentes e nao duplique itens repetidos entre capturas sobrepostas.',
    'Valores monetarios devem ser numeros decimais em BRL e datas devem usar yyyy-mm-dd.',
    'valor_total de cada item e o total da linha, depois de quantidade vezes valor unitario quando ambos estiverem visiveis.',
    'Frete e desconto pertencem ao cabecalho; nao os transforme em produto quando estiverem claramente identificados.',
    'Para destino_sugerido use somente ESTOQUE, CONSUMO, EQUIPAMENTO ou IGNORAR.',
    'ESTOQUE significa material mantido para uso ou revenda; CONSUMO significa gasto consumido imediatamente; EQUIPAMENTO significa ferramenta ou ativo duravel.',
    'Tipos, categorias e unidades sao apenas sugestoes. Use uma opcao valida somente quando houver correspondencia clara; caso contrario deixe string vazia.',
    'pagamento.comprovado deve ser true somente quando o documento mostrar explicitamente que a transacao foi paga ou aprovada.',
    'documento.recebido deve ser true somente quando houver evidencia de retirada, entrega ou compra presencial concluida.',
    `Tipos validos: ${JSON.stringify(tipos)}.`,
    `Categorias validas por tipo: ${JSON.stringify(categoriasPorTipo)}.`,
    `Unidades validas: ${JSON.stringify(unidades)}.`,
    `Formas de pagamento validas: ${JSON.stringify(formas)}.`
  ].join('\n');
}

function limitarListaPromptDocumentoCompra_(lista, limite) {
  return (Array.isArray(lista) ? lista : [])
    .map(item => String(item || '').trim())
    .filter(Boolean)
    .slice(0, Math.max(1, Number(limite) || 40));
}

function getSchemaOpenAIDocumentoCompra_() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      documento: {
        type: 'object',
        additionalProperties: false,
        properties: {
          tipo_documento: { type: 'string', description: 'Ex.: NOTA_FISCAL, CUPOM, RECIBO ou PEDIDO_ECOMMERCE.' },
          fornecedor: { type: 'string' },
          numero_documento: { type: 'string' },
          numero_pedido: { type: 'string' },
          data_compra: { type: 'string', description: 'Data yyyy-mm-dd ou string vazia.' },
          data_vencimento: { type: 'string', description: 'Primeiro vencimento yyyy-mm-dd ou string vazia.' },
          subtotal: { type: 'number' },
          frete: { type: 'number' },
          desconto: { type: 'number' },
          valor_total: { type: 'number' },
          moeda: { type: 'string' },
          recebido: { type: 'boolean' },
          observacao: { type: 'string' }
        },
        required: ['tipo_documento', 'fornecedor', 'numero_documento', 'numero_pedido', 'data_compra', 'data_vencimento', 'subtotal', 'frete', 'desconto', 'valor_total', 'moeda', 'recebido', 'observacao']
      },
      pagamento: {
        type: 'object',
        additionalProperties: false,
        properties: {
          comprovado: { type: 'boolean' },
          data_pagamento: { type: 'string' },
          forma_pagamento: { type: 'string' },
          parcelas: { type: 'integer' },
          valor_pago: { type: 'number' }
        },
        required: ['comprovado', 'data_pagamento', 'forma_pagamento', 'parcelas', 'valor_pago']
      },
      itens: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            descricao_original: { type: 'string' },
            sku: { type: 'string' },
            item_nome: { type: 'string' },
            tipo_sugerido: { type: 'string' },
            categoria_sugerida: { type: 'string' },
            unidade: { type: 'string' },
            quantidade: { type: 'number' },
            comprimento_cm: { type: 'number' },
            largura_cm: { type: 'number' },
            espessura_cm: { type: 'number' },
            valor_unitario: { type: 'number' },
            valor_total: { type: 'number' },
            destino_sugerido: { type: 'string' },
            recebido: { type: 'boolean' },
            confianca: { type: 'number' },
            alertas: { type: 'array', items: { type: 'string' } }
          },
          required: ['descricao_original', 'sku', 'item_nome', 'tipo_sugerido', 'categoria_sugerida', 'unidade', 'quantidade', 'comprimento_cm', 'largura_cm', 'espessura_cm', 'valor_unitario', 'valor_total', 'destino_sugerido', 'recebido', 'confianca', 'alertas']
        }
      },
      confianca: { type: 'number' },
      alertas: { type: 'array', items: { type: 'string' } }
    },
    required: ['documento', 'pagamento', 'itens', 'confianca', 'alertas']
  };
}

function normalizarDocumentoCompraExtraido_(extraido) {
  const dados = extraido || {};
  const documentoRaw = dados.documento || {};
  const pagamentoRaw = dados.pagamento || {};
  const validacoes = obterValidacoes();
  let estoqueAtivo = [];
  try {
    estoqueAtivo = listarEstoque(true);
  } catch (error) {
    estoqueAtivo = [];
  }
  const alertas = Array.isArray(dados.alertas) ? dados.alertas.map(String).filter(Boolean) : [];
  let dataCompra = normalizarDataOpcionalDocumentoCompra_(documentoRaw.data_compra);
  if (!dataCompra) {
    dataCompra = formatarDataYmdFinanceiro(new Date());
    alertas.push('Data da compra nao identificada; revise a data sugerida.');
  }
  const itens = (Array.isArray(dados.itens) ? dados.itens : [])
    .map((item, index) => normalizarItemDocumentoCompraExtraido_(item, index, documentoRaw.recebido, validacoes, estoqueAtivo))
    .filter(item => item.descricao_original || item.item_nome || item.valor_total > 0);
  if (itens.length === 0) alertas.push('Nenhum item foi identificado; revise o arquivo ou descarte este documento.');
  const somaItens = round2Financeiro(itens.reduce((acc, item) => acc + item.valor_total, 0));
  const frete = Math.max(0, round2Financeiro(parseNumeroBR(documentoRaw.frete)));
  const desconto = Math.max(0, round2Financeiro(parseNumeroBR(documentoRaw.desconto)));
  const subtotalInformado = round2Financeiro(parseNumeroBR(documentoRaw.subtotal));
  const subtotal = subtotalInformado > 0 ? subtotalInformado : somaItens;
  let valorTotal = round2Financeiro(parseNumeroBR(documentoRaw.valor_total));
  if (valorTotal <= 0) valorTotal = round2Financeiro(subtotal + frete - desconto);
  if (Math.abs(somaItens - subtotal) > 0.05) alertas.push('A soma dos itens difere do subtotal informado.');
  if (Math.abs(round2Financeiro(subtotal + frete - desconto) - valorTotal) > 0.05) alertas.push('Subtotal, frete e desconto nao reconciliam com o total.');

  return {
    documento: {
      tipo_documento: String(documentoRaw.tipo_documento || '').trim().slice(0, 60),
      fornecedor: String(documentoRaw.fornecedor || '').trim().slice(0, 120),
      numero_documento: String(documentoRaw.numero_documento || '').trim().slice(0, 80),
      numero_pedido: String(documentoRaw.numero_pedido || '').trim().slice(0, 80),
      data_compra: dataCompra,
      data_vencimento: normalizarDataOpcionalDocumentoCompra_(documentoRaw.data_vencimento),
      subtotal,
      frete,
      desconto,
      valor_total: valorTotal,
      moeda: String(documentoRaw.moeda || 'BRL').trim().toUpperCase().slice(0, 8) || 'BRL',
      pago_por: '',
      forma_pagamento: normalizarValorListaDocumentoCompra_(pagamentoRaw.forma_pagamento, validacoes?.formasPagamento),
      parcelas: Math.max(1, Math.min(60, Math.floor(parseNumeroBR(pagamentoRaw.parcelas) || 1))),
      recebido: documentoRaw.recebido === true,
      pagamento_detectado: pagamentoRaw.comprovado === true,
      pagamento_confirmado: false,
      data_pagamento: normalizarDataOpcionalDocumentoCompra_(pagamentoRaw.data_pagamento),
      valor_pago_confirmado: 0,
      observacao: String(documentoRaw.observacao || '').trim().slice(0, 500),
      confianca: Math.max(0, Math.min(1, Number(dados.confianca || 0) || 0))
    },
    itens,
    alertas: [...new Set(alertas)]
  };
}

function normalizarItemDocumentoCompraExtraido_(raw, index, documentoRecebido, validacoes, estoqueAtivo) {
  const item = raw || {};
  const alertas = Array.isArray(item.alertas) ? item.alertas.map(String).filter(Boolean) : [];
  let quantidade = parseNumeroBR(item.quantidade);
  if (quantidade <= 0) {
    quantidade = 1;
    alertas.push(`Quantidade do item ${index + 1} nao identificada; usando 1.`);
  }
  let valorUnitario = round2Financeiro(parseNumeroBR(item.valor_unitario));
  let valorTotal = round2Financeiro(parseNumeroBR(item.valor_total));
  if (valorTotal <= 0 && valorUnitario > 0) valorTotal = round2Financeiro(valorUnitario * quantidade);
  if (valorUnitario <= 0 && valorTotal > 0) valorUnitario = round2Financeiro(valorTotal / quantidade);
  const tipo = normalizarValorListaDocumentoCompra_(item.tipo_sugerido, validacoes?.tipos);
  const categorias = Array.isArray(validacoes?.categoriasPorTipo?.[tipo])
    ? validacoes.categoriasPorTipo[tipo]
    : validacoes?.categorias;
  const destino = normalizarDestinoItemDocumentoCompra_(item.destino_sugerido);
  const normalizado = {
    descricao_original: String(item.descricao_original || '').trim().slice(0, 240),
    sku: String(item.sku || '').trim().slice(0, 100),
    item_nome: String(item.item_nome || item.descricao_original || '').trim().slice(0, 160),
    tipo,
    categoria: normalizarValorListaDocumentoCompra_(item.categoria_sugerida, categorias),
    unidade: normalizarValorListaDocumentoCompra_(item.unidade, validacoes?.unidades) || String(item.unidade || '').trim().slice(0, 20),
    quantidade,
    comprimento_cm: Math.max(0, parseNumeroBR(item.comprimento_cm)),
    largura_cm: Math.max(0, parseNumeroBR(item.largura_cm)),
    espessura_cm: Math.max(0, parseNumeroBR(item.espessura_cm)),
    valor_unitario: valorUnitario,
    valor_total: valorTotal,
    destino,
    estoque_id_alvo: '',
    recebido: item.recebido === true || documentoRecebido === true,
    confianca: Math.max(0, Math.min(1, Number(item.confianca || 0) || 0)),
    alertas
  };
  normalizado.estoque_id_alvo = sugerirEstoqueIdDocumentoCompra_(normalizado, estoqueAtivo);
  return normalizado;
}

function normalizarValorListaDocumentoCompra_(valor, lista) {
  const raw = String(valor || '').trim();
  if (!raw) return '';
  const alvo = normalizarTextoDocumentoCompra_(raw);
  return (Array.isArray(lista) ? lista : []).find(item => normalizarTextoDocumentoCompra_(item) === alvo) || '';
}

function normalizarTextoDocumentoCompra_(valor) {
  return String(valor || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
}

function normalizarDataOpcionalDocumentoCompra_(valor) {
  try {
    return valor ? normalizarDataFinanceiro(valor, false, 'Data') : '';
  } catch (error) {
    return '';
  }
}

function sugerirEstoqueIdDocumentoCompra_(item, estoqueDisponivel) {
  const estoque = Array.isArray(estoqueDisponivel) ? estoqueDisponivel : [];
  const nome = normalizarTextoDocumentoCompra_(item?.item_nome);
  if (!nome) return '';
  const candidatos = estoque.filter(est => normalizarTextoDocumentoCompra_(est.item) === nome);
  if (candidatos.length !== 1) return '';
  const alvo = candidatos[0];
  if (item.tipo && normalizarTextoDocumentoCompra_(alvo.tipo) !== normalizarTextoDocumentoCompra_(item.tipo)) return '';
  if (item.unidade && normalizarTextoDocumentoCompra_(alvo.unidade) !== normalizarTextoDocumentoCompra_(item.unidade)) return '';
  return String(alvo.ID || '').trim();
}

function confirmarDocumentoCompra(id, payloadOverride) {
  const ambiente = payloadOverride?.ambiente || payloadOverride?._db_env || payloadOverride?.db_env || '';
  return executarComAmbienteDocumentosCompra_(ambiente, () => confirmarDocumentoCompraAtual_(id, payloadOverride));
}

function confirmarDocumentoCompraAtual_(id, payloadOverride) {
  assertCanWrite('Confirmacao de documento de compra');
  const documentoId = String(id || '').trim();
  if (!documentoId) throw new Error('Documento de compra invalido.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const atual = obterDocumentoCompraPorId_(documentoId);
    if (!atual || String(atual.ativo).toLowerCase() !== 'true') throw new Error('Documento de compra nao encontrado.');
    if (atual.status === 'DESCARTADO') throw new Error('Documento de compra ja descartado.');
    if (atual.status === 'CONFIRMADO') {
      return { ok: true, documento: { ...atual, itens: listarItensDocumentoCompra_(documentoId) }, reutilizado: true };
    }

    const normalizado = normalizarConfirmacaoDocumentoCompra_(atual, payloadOverride);
    persistirRevisaoDocumentoCompra_(documentoId, normalizado);

    const compraFinanceira = obterOuCriarCompraFinanceiraDocumento_(documentoId, normalizado);
    updateById(ABA_DOCUMENTOS_COMPRA, 'ID', documentoId, {
      compra_id_financeira: compraFinanceira.ID,
      atualizado_em: new Date()
    }, DOCUMENTOS_COMPRA_SCHEMA);

    const resultadosEstoque = [];
    const movimentosPorItem = {};
    listarMovimentosEstoqueDocumentoCompra_().forEach(movimento => {
      movimentosPorItem[String(movimento.origem_item_id || '').trim()] = movimento;
    });
    normalizado.itens.forEach(item => {
      if (!item.recebido || !['ESTOQUE', 'EQUIPAMENTO'].includes(item.destino)) return;
      const resultado = aplicarEntradaEstoqueDocumentoCompra_(normalizado.documento, item, movimentosPorItem[item.ID] || null);
      resultadosEstoque.push(resultado);
      updateById(ABA_DOCUMENTOS_COMPRA_ITENS, 'ID', item.ID, {
        estoque_id_resultante: resultado.estoque_id,
        custo_unitario_estoque: resultado.custo_unitario,
        atualizado_em: new Date()
      }, DOCUMENTOS_COMPRA_ITENS_SCHEMA);
    });

    let pagamento = null;
    if (normalizado.documento.pagamento_confirmado) {
      pagamento = registrarPagamentoConfirmadoDocumentoCompra_(normalizado.documento, compraFinanceira);
      updateById(ABA_DOCUMENTOS_COMPRA, 'ID', documentoId, {
        pagamento_id_confirmado: String(pagamento?.ID || '').trim(),
        atualizado_em: new Date()
      }, DOCUMENTOS_COMPRA_SCHEMA);
    }

    const agora = new Date();
    const aguardandoRecebimento = normalizado.itens.some(item =>
      ['ESTOQUE', 'EQUIPAMENTO'].includes(item.destino) && !item.recebido
    );
    const statusFinal = aguardandoRecebimento ? 'CONFIRMADO_AGUARDANDO_RECEBIMENTO' : 'CONFIRMADO';
    updateById(ABA_DOCUMENTOS_COMPRA, 'ID', documentoId, {
      status: statusFinal,
      erro: '',
      confirmado_em: atual.confirmado_em || agora,
      atualizado_em: agora
    }, DOCUMENTOS_COMPRA_SCHEMA);
    const confirmado = obterDocumentoCompraPorId_(documentoId);
    if (typeof moverArquivosDocumentoCompraAposConfirmacao_ === 'function') {
      moverArquivosDocumentoCompraAposConfirmacao_(confirmado);
    }
    return {
      ok: true,
      documento: { ...confirmado, itens: listarItensDocumentoCompra_(documentoId) },
      compraFinanceira,
      pagamento,
      entradasEstoque: resultadosEstoque,
      aguardandoRecebimento
    };
  } catch (error) {
    try {
      updateById(ABA_DOCUMENTOS_COMPRA, 'ID', documentoId, {
        status: 'ERRO_CONFIRMACAO',
        erro: error?.message || String(error),
        atualizado_em: new Date()
      }, DOCUMENTOS_COMPRA_SCHEMA);
    } catch (updateError) {
      // A falha original deve prevalecer.
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function normalizarConfirmacaoDocumentoCompra_(atual, payloadOverride) {
  const override = payloadOverride || {};
  const itensAtuais = listarItensDocumentoCompra_(atual.ID);
  const itensOverride = Array.isArray(override.itens) ? override.itens : itensAtuais;
  const mapaAtuais = {};
  itensAtuais.forEach(item => { mapaAtuais[item.ID] = item; });
  if (itensOverride.length === 0) throw new Error('O documento precisa ter ao menos um item.');
  const validacoes = obterValidacoes();
  const idsRecebidos = new Set();
  const itens = itensOverride.map((raw, index) => {
    const idRecebido = String(raw?.ID || raw?.id || '').trim();
    const novo = /^NOVO_[A-Za-z0-9_-]{1,100}$/.test(idRecebido);
    const id = novo ? gerarId('DOCIT') : idRecebido;
    if (!idRecebido || idsRecebidos.has(idRecebido)) throw new Error(`Item ${index + 1} invalido ou duplicado.`);
    idsRecebidos.add(idRecebido);
    const base = novo ? {
      ID: id,
      documento_id: atual.ID,
      descricao_original: '',
      sku: '',
      confianca: 0,
      alertas: [],
      ativo: true,
      criado_em: new Date(),
      atualizado_em: new Date()
    } : mapaAtuais[id];
    if (!base) throw new Error(`Item ${index + 1} nao pertence a este documento.`);
    const quantidade = parseNumeroBR(raw.quantidade ?? base.quantidade);
    const valorTotal = round2Financeiro(parseNumeroBR(raw.valor_total ?? base.valor_total));
    let valorUnitario = round2Financeiro(parseNumeroBR(raw.valor_unitario ?? base.valor_unitario));
    if (quantidade <= 0) throw new Error(`Quantidade invalida no item ${index + 1}.`);
    if (valorTotal <= 0) throw new Error(`Valor total invalido no item ${index + 1}.`);
    if (valorUnitario <= 0) valorUnitario = round2Financeiro(valorTotal / quantidade);
    const destino = normalizarDestinoItemDocumentoCompra_(raw.destino ?? base.destino);
    const tipo = String(raw.tipo ?? base.tipo ?? '').trim();
    const categoria = String(raw.categoria ?? base.categoria ?? '').trim();
    const unidade = String(raw.unidade ?? base.unidade ?? '').trim();
    const itemNome = String(raw.item_nome ?? base.item_nome ?? '').trim().slice(0, 160);
    const recebido = raw.recebido === true || String(raw.recebido).toLowerCase() === 'true';
    const precisaEstoque = ['ESTOQUE', 'EQUIPAMENTO'].includes(destino);
    if (precisaEstoque) {
      if (!itemNome) throw new Error(`Informe o nome do item ${index + 1}.`);
      if (!tipo) throw new Error(`Selecione o tipo do item ${index + 1}.`);
      if (!categoriaValidaDocumentoCompra_(tipo, categoria, validacoes)) {
        throw new Error(`Categoria invalida para o tipo do item ${index + 1}.`);
      }
      if (!unidade) throw new Error(`Selecione a unidade do item ${index + 1}.`);
    }
    const comprimento = Math.max(0, parseNumeroBR(raw.comprimento_cm ?? base.comprimento_cm));
    const largura = Math.max(0, parseNumeroBR(raw.largura_cm ?? base.largura_cm));
    const espessura = Math.max(0, parseNumeroBR(raw.espessura_cm ?? base.espessura_cm));
    if (precisaEstoque && normalizarTextoDocumentoCompra_(tipo) === 'MADEIRA' && recebido && (!comprimento || !largura || !espessura)) {
      throw new Error(`Informe comprimento, largura e espessura da madeira no item ${index + 1}.`);
    }
    if (base.movimento_id) {
      const camposImutaveis = [
        [quantidade, base.quantidade], [valorTotal, base.valor_total], [tipo, base.tipo],
        [categoria, base.categoria], [unidade, base.unidade], [comprimento, base.comprimento_cm],
        [largura, base.largura_cm], [espessura, base.espessura_cm], [itemNome, base.item_nome],
        [destino, base.destino], [String(raw.estoque_id_alvo ?? base.estoque_id_alvo ?? '').trim(), base.estoque_id_alvo]
      ];
      const alterou = camposImutaveis.some(([novo, anterior]) => {
        if (typeof novo === 'number') return Math.abs(parseNumeroBR(novo) - parseNumeroBR(anterior)) > 0.000001;
        return normalizarTextoDocumentoCompra_(novo) !== normalizarTextoDocumentoCompra_(anterior);
      });
      if (alterou) throw new Error(`O item ${index + 1} ja possui uma entrada de estoque iniciada e nao pode ter quantidade, custo ou classificacao alterados.`);
    }
    return {
      ...base,
      ID: id,
      ordem: index + 1,
      item_nome: itemNome,
      tipo,
      categoria,
      unidade,
      quantidade,
      comprimento_cm: comprimento,
      largura_cm: largura,
      espessura_cm: espessura,
      valor_unitario: valorUnitario,
      valor_total: valorTotal,
      destino,
      estoque_id_alvo: String(raw.estoque_id_alvo ?? base.estoque_id_alvo ?? '').trim(),
      recebido,
      _novo: novo
    };
  });
  const itensRemovidos = itensAtuais.filter(item => !idsRecebidos.has(item.ID));
  if (itensRemovidos.some(item => !!item.movimento_id)) {
    throw new Error('Um item que ja possui entrada de estoque nao pode ser removido do documento.');
  }

  const somaItens = round2Financeiro(itens.reduce((acc, item) => acc + item.valor_total, 0));
  const subtotalInformado = round2Financeiro(parseNumeroBR(override.subtotal ?? atual.subtotal));
  const subtotal = subtotalInformado > 0 ? subtotalInformado : somaItens;
  const frete = Math.max(0, round2Financeiro(parseNumeroBR(override.frete ?? atual.frete)));
  const desconto = Math.max(0, round2Financeiro(parseNumeroBR(override.desconto ?? atual.desconto)));
  const totalInformado = round2Financeiro(parseNumeroBR(override.valor_total ?? atual.valor_total));
  const totalCalculado = round2Financeiro(subtotal + frete - desconto);
  const possuiEntradaIniciada = itens.some(item => !!item.movimento_id);
  if (possuiEntradaIniciada) {
    const alterouRateio = Math.abs(frete - parseNumeroBR(atual.frete)) > 0.009 ||
      Math.abs(desconto - parseNumeroBR(atual.desconto)) > 0.009 ||
      Math.abs(totalInformado - parseNumeroBR(atual.valor_total)) > 0.009;
    if (alterouRateio) throw new Error('Frete, desconto e total nao podem ser alterados depois que um item entrou no estoque.');
  }
  if (Math.abs(somaItens - subtotal) > 0.05) {
    throw new Error(`A soma dos itens (${formatarMoedaDocumentoCompra_(somaItens)}) difere do subtotal (${formatarMoedaDocumentoCompra_(subtotal)}).`);
  }
  if (totalInformado <= 0 || Math.abs(totalInformado - totalCalculado) > 0.05) {
    throw new Error(`O total deve ser igual a subtotal + frete - desconto (${formatarMoedaDocumentoCompra_(totalCalculado)}).`);
  }
  const dataCompra = normalizarDataFinanceiro(override.data_compra ?? atual.data_compra, true, 'Data da compra');
  const formaPagamento = validarFormaPagamentoFinanceiro(override.forma_pagamento ?? atual.forma_pagamento, false);
  const parcelas = normalizarParcelasFinanceiro(override.parcelas ?? atual.parcelas, formaPagamento);
  const dataVencimento = normalizarDataFinanceiro(override.data_vencimento ?? atual.data_vencimento, false, 'Data de vencimento');
  const pagamentoConfirmado = !!String(atual.pagamento_id_confirmado || '').trim() ||
    override.pagamento_confirmado === true || String(override.pagamento_confirmado).toLowerCase() === 'true';
  const dataPagamento = normalizarDataFinanceiro(override.data_pagamento ?? atual.data_pagamento, pagamentoConfirmado, 'Data de pagamento');
  const valorPago = round2Financeiro(parseNumeroBR(override.valor_pago_confirmado ?? atual.valor_pago_confirmado));
  if (pagamentoConfirmado && valorPago <= 0) throw new Error('Informe o valor pago agora.');
  if (pagamentoConfirmado && !formaPagamento) throw new Error('Informe a forma de pagamento.');
  const rateados = ratearCustosDocumentoCompra_(itens, frete, desconto);

  return {
    documento: {
      ...atual,
      fornecedor: String(override.fornecedor ?? atual.fornecedor ?? '').trim().slice(0, 120),
      numero_documento: String(override.numero_documento ?? atual.numero_documento ?? '').trim().slice(0, 80),
      numero_pedido: String(override.numero_pedido ?? atual.numero_pedido ?? '').trim().slice(0, 80),
      tipo_documento: String(override.tipo_documento ?? atual.tipo_documento ?? '').trim().slice(0, 60),
      data_compra: dataCompra,
      data_vencimento: dataVencimento,
      subtotal,
      frete,
      desconto,
      valor_total: totalInformado,
      pago_por: validarPagoPorFinanceiro(override.pago_por ?? atual.pago_por, false),
      forma_pagamento: formaPagamento,
      parcelas,
      recebido: override.recebido === true || String(override.recebido).toLowerCase() === 'true',
      pagamento_confirmado: pagamentoConfirmado,
      data_pagamento: dataPagamento,
      valor_pago_confirmado: valorPago,
      observacao: String(override.observacao ?? atual.observacao ?? '').trim().slice(0, 500)
    },
    itens: rateados,
    itensRemovidos
  };
}

function categoriaValidaDocumentoCompra_(tipo, categoria, validacoes) {
  const t = normalizarTextoDocumentoCompra_(tipo);
  const c = normalizarTextoDocumentoCompra_(categoria);
  if (!t || !c) return false;
  const porTipo = validacoes?.categoriasPorTipo || {};
  const lista = Array.isArray(porTipo[t]) ? porTipo[t] : (validacoes?.categorias || []);
  return lista.some(item => normalizarTextoDocumentoCompra_(item) === c);
}

function formatarMoedaDocumentoCompra_(valor) {
  return `R$ ${round2Financeiro(valor).toFixed(2).replace('.', ',')}`;
}

function ratearCustosDocumentoCompra_(itens, frete, desconto) {
  const baseTotal = round2Financeiro(itens.reduce((acc, item) => acc + item.valor_total, 0));
  let freteAplicado = 0;
  let descontoAplicado = 0;
  return itens.map((item, index) => {
    const ultimo = index === itens.length - 1;
    const proporcao = baseTotal > 0 ? item.valor_total / baseTotal : 0;
    const freteRateado = ultimo ? round2Financeiro(frete - freteAplicado) : round2Financeiro(frete * proporcao);
    const descontoRateado = ultimo ? round2Financeiro(desconto - descontoAplicado) : round2Financeiro(desconto * proporcao);
    freteAplicado = round2Financeiro(freteAplicado + freteRateado);
    descontoAplicado = round2Financeiro(descontoAplicado + descontoRateado);
    const custoTotal = round2Financeiro(item.valor_total + freteRateado - descontoRateado);
    if (custoTotal < 0) throw new Error(`Rateio gerou custo negativo para ${item.item_nome || item.descricao_original}.`);
    return { ...item, frete_rateado: freteRateado, desconto_rateado: descontoRateado, custo_total: custoTotal };
  });
}

function persistirRevisaoDocumentoCompra_(documentoId, normalizado) {
  const doc = normalizado.documento;
  updateById(ABA_DOCUMENTOS_COMPRA, 'ID', documentoId, {
    tipo_documento: doc.tipo_documento,
    fornecedor: doc.fornecedor,
    numero_documento: doc.numero_documento,
    numero_pedido: doc.numero_pedido,
    data_compra: doc.data_compra,
    data_vencimento: doc.data_vencimento,
    subtotal: doc.subtotal,
    frete: doc.frete,
    desconto: doc.desconto,
    valor_total: doc.valor_total,
    pago_por: doc.pago_por,
    forma_pagamento: doc.forma_pagamento,
    parcelas: doc.parcelas,
    recebido: doc.recebido,
    pagamento_confirmado: doc.pagamento_confirmado,
    data_pagamento: doc.data_pagamento,
    valor_pago_confirmado: doc.valor_pago_confirmado,
    observacao: doc.observacao,
    atualizado_em: new Date()
  }, DOCUMENTOS_COMPRA_SCHEMA);
  (normalizado.itensRemovidos || []).forEach(item => {
    updateById(ABA_DOCUMENTOS_COMPRA_ITENS, 'ID', item.ID, {
      ativo: false,
      atualizado_em: new Date()
    }, DOCUMENTOS_COMPRA_ITENS_SCHEMA);
  });
  const itensNovos = normalizado.itens.filter(item => item._novo).map(item => ({
    ...item,
    documento_id: documentoId,
    descricao_original: item.descricao_original || item.item_nome,
    sku: item.sku || '',
    custo_unitario_estoque: 0,
    estoque_id_resultante: '',
    confianca: 0,
    alertas_json: '[]',
    ativo: true,
    criado_em: item.criado_em || new Date(),
    atualizado_em: new Date()
  }));
  if (itensNovos.length > 0) {
    insertMany(ABA_DOCUMENTOS_COMPRA_ITENS, itensNovos, DOCUMENTOS_COMPRA_ITENS_SCHEMA);
  }
  normalizado.itens.filter(item => !item._novo).forEach(item => {
    updateById(ABA_DOCUMENTOS_COMPRA_ITENS, 'ID', item.ID, {
      ordem: item.ordem,
      item_nome: item.item_nome,
      tipo: item.tipo,
      categoria: item.categoria,
      unidade: item.unidade,
      quantidade: item.quantidade,
      comprimento_cm: item.comprimento_cm,
      largura_cm: item.largura_cm,
      espessura_cm: item.espessura_cm,
      valor_unitario: item.valor_unitario,
      valor_total: item.valor_total,
      frete_rateado: item.frete_rateado,
      desconto_rateado: item.desconto_rateado,
      custo_total: item.custo_total,
      destino: item.destino,
      estoque_id_alvo: item.estoque_id_alvo,
      recebido: item.recebido,
      atualizado_em: new Date()
    }, DOCUMENTOS_COMPRA_ITENS_SCHEMA);
  });
}

function obterOuCriarCompraFinanceiraDocumento_(documentoId, normalizado) {
  const sheet = getSheet(ABA_COMPRAS);
  let existente = null;
  if (sheet) {
    ensureSchema(sheet, COMPRAS_SCHEMA);
    existente = rowsToObjects(sheet).find(item =>
      String(item.documento_compra_id || '').trim() === String(documentoId || '').trim() &&
      String(item.ativo).toLowerCase() === 'true'
    ) || null;
  }
  if (existente) {
    const pagamentos = listarPagamentos(true).filter(p =>
      String(p.origem_tipo || '').trim().toUpperCase() === ORIGEM_TIPO_COMPRA &&
      String(p.origem_id || '').trim() === String(existente.ID || '').trim()
    );
    const totalAnterior = getTotalPrevistoCompraFinanceiro(existente);
    if (pagamentos.length > 0 && Math.abs(totalAnterior - normalizado.documento.valor_total) > 0.009) {
      throw new Error('O total do documento nao pode ser alterado depois de registrar pagamento.');
    }
    const doc = normalizado.documento;
    if (pagamentos.length > 0) {
      const alterouPlano = Number(existente.parcelas || 1) !== Number(doc.parcelas || 1) ||
        normalizarTextoDocumentoCompra_(existente.forma_pagamento) !== normalizarTextoDocumentoCompra_(doc.forma_pagamento) ||
        formatarDataYmdFinanceiroSafe(existente.data_vencimento) !== formatarDataYmdFinanceiroSafe(doc.data_vencimento);
      if (alterouPlano) throw new Error('Parcelas, forma e vencimento nao podem ser alterados depois de registrar pagamento.');
    }
    updateById(ABA_COMPRAS, 'ID', existente.ID, {
      item: `Documento ${doc.numero_pedido || doc.numero_documento || documentoId}${doc.fornecedor ? ` - ${doc.fornecedor}` : ''}`.slice(0, 160),
      fornecedor: doc.fornecedor,
      pago_por: doc.pago_por,
      comprado_em: doc.data_compra,
      observacao: `Gerado pelo documento de compra ${documentoId}${doc.observacao ? `: ${doc.observacao}` : ''}`.slice(0, 500)
    }, COMPRAS_SCHEMA);
    if (pagamentos.length === 0) {
      const detalhes = normalizarParcelasDetalhePayloadFinanceiro(
        [], doc.parcelas, doc.data_vencimento || doc.data_compra || new Date(), doc.valor_total
      );
      updateById(ABA_COMPRAS, 'ID', existente.ID, {
        item: `Documento ${doc.numero_pedido || doc.numero_documento || documentoId}${doc.fornecedor ? ` - ${doc.fornecedor}` : ''}`.slice(0, 160),
        valor_unit: doc.valor_total,
        fornecedor: doc.fornecedor,
        pago_por: doc.pago_por,
        comprado_em: doc.data_compra,
        data_vencimento: doc.data_vencimento,
        forma_pagamento: doc.forma_pagamento,
        parcelas: doc.parcelas,
        parcelas_detalhe_json: serializarParcelasDetalheFinanceiro(detalhes)
      }, COMPRAS_SCHEMA);
      regerarParcelasFinanceirasOrigemComPagamentos(ORIGEM_TIPO_COMPRA, existente.ID);
      return rowsToObjects(getSheet(ABA_COMPRAS)).find(item => String(item.ID || '').trim() === String(existente.ID || '').trim()) || existente;
    }
    return existente;
  }

  const doc = normalizado.documento;
  const inicioParcelas = doc.data_vencimento || doc.data_compra || new Date();
  const parcelasDetalhe = normalizarParcelasDetalhePayloadFinanceiro(
    [],
    doc.parcelas,
    inicioParcelas,
    doc.valor_total
  );
  const referencia = doc.numero_pedido || doc.numero_documento || documentoId;
  const novo = {
    ID: gerarId('COM'),
    tipo: 'DOCUMENTO',
    item: `Documento ${referencia}${doc.fornecedor ? ` - ${doc.fornecedor}` : ''}`.slice(0, 160),
    unidade: 'UN',
    valor_unit: doc.valor_total,
    ativo: true,
    criado_em: new Date(),
    quantidade: 1,
    comprimento_cm: '',
    largura_cm: '',
    espessura_cm: '',
    categoria: 'MULTI_ITEM',
    fornecedor: doc.fornecedor,
    pago_por: doc.pago_por,
    potencia: '',
    voltagem: '',
    comprado_em: doc.data_compra,
    data_pagamento: '',
    data_vencimento: doc.data_vencimento,
    forma_pagamento: doc.forma_pagamento,
    parcelas: doc.parcelas,
    parcelas_detalhe_json: serializarParcelasDetalheFinanceiro(parcelasDetalhe),
    vida_util_mes: '',
    observacao: `Gerado pelo documento de compra ${documentoId}${doc.observacao ? `: ${doc.observacao}` : ''}`.slice(0, 500),
    adicionado_estoque: true,
    estoque_id: '',
    origem_compra_id: '',
    documento_compra_id: documentoId,
    financeiro_somente: true
  };
  insert(ABA_COMPRAS, novo, COMPRAS_SCHEMA);
  gerarParcelasFinanceirasOrigem(ORIGEM_TIPO_COMPRA, novo.ID);
  return novo;
}

function registrarPagamentoConfirmadoDocumentoCompra_(documento, compraFinanceira) {
  const compraId = String(compraFinanceira?.ID || '').trim();
  if (!compraId) throw new Error('Compra financeira nao encontrada para registrar pagamento.');
  const valor = round2Financeiro(parseNumeroBR(documento.valor_pago_confirmado));
  const payload = {
    data_pagamento: documento.data_pagamento,
    valor_pago: valor,
    forma_pagamento: documento.forma_pagamento,
    observacao: `Pagamento confirmado pelo documento ${documento.ID}`,
    client_request_id: `PAG_DOCUMENTO_COMPRA_${documento.ID}`
  };
  if (Number(documento.parcelas || 1) > 1) {
    payload.distribuir_automaticamente = true;
  }
  const resultado = registrarPagamentoCompra(compraId, payload);
  return resultado?.pagamentoCriado || resultado;
}

function calcularQuantidadeEstoqueDocumentoCompra_(item) {
  if (normalizarTextoDocumentoCompra_(item.tipo) !== 'MADEIRA') return parseNumeroBR(item.quantidade);
  const volumePeca = (parseNumeroBR(item.comprimento_cm) * parseNumeroBR(item.largura_cm) * parseNumeroBR(item.espessura_cm)) / 1000000;
  return Number((volumePeca * parseNumeroBR(item.quantidade)).toFixed(6));
}

function aplicarEntradaEstoqueDocumentoCompra_(documento, item, movimentoConhecido) {
  const movimentoExistente = arguments.length >= 3
    ? movimentoConhecido
    : buscarMovimentoEstoqueDocumentoCompra_(item.ID);
  if (movimentoExistente && String(movimentoExistente.status || '').toUpperCase() === 'APLICADO') {
    const estoqueAtual = obterItemEstoque(String(movimentoExistente.estoque_id || '').trim());
    return {
      estoque_id: String(movimentoExistente.estoque_id || '').trim(),
      movimento_id: String(movimentoExistente.ID || '').trim(),
      quantidade: parseNumeroBR(movimentoExistente.quantidade),
      custo_unitario: parseNumeroBR(estoqueAtual?.custo_unitario || estoqueAtual?.valor_unit || movimentoExistente.custo_unitario),
      reutilizado: true
    };
  }

  const quantidadeEntrada = calcularQuantidadeEstoqueDocumentoCompra_(item);
  if (quantidadeEntrada <= 0) throw new Error(`Quantidade de estoque invalida para ${item.item_nome}.`);
  const custoUnitarioEntrada = Number((parseNumeroBR(item.custo_total) / quantidadeEntrada).toFixed(6));
  const estoqueIdAlvo = String(item.estoque_id_alvo || '').trim();
  const sheetEstoque = getSheet(ABA_ESTOQUE);
  if (!sheetEstoque) throw new Error('Aba ESTOQUE nao encontrada.');
  ensureSchema(sheetEstoque, ESTOQUE_SCHEMA);
  const linhas = rowsToObjects(sheetEstoque);

  if (estoqueIdAlvo) {
    const alvo = linhas.find(est => String(est.ID || '').trim() === estoqueIdAlvo && String(est.ativo).toLowerCase() === 'true');
    if (!alvo) throw new Error(`Item de estoque selecionado nao encontrado para ${item.item_nome}.`);
    validarCompatibilidadeEstoqueDocumentoCompra_(alvo, item);
    if (movimentoExistente && String(movimentoExistente.estoque_id || '').trim() !== estoqueIdAlvo) {
      throw new Error(`A entrada de ${item.item_nome} foi iniciada para outro item de estoque.`);
    }
    const saldoAnterior = movimentoExistente
      ? parseNumeroBR(movimentoExistente.saldo_anterior)
      : parseNumeroBR(alvo.quantidade);
    const quantidadeMovimento = movimentoExistente
      ? parseNumeroBR(movimentoExistente.quantidade)
      : quantidadeEntrada;
    const saldoPosterior = movimentoExistente
      ? parseNumeroBR(movimentoExistente.saldo_posterior)
      : Number((saldoAnterior + quantidadeEntrada).toFixed(6));
    const custoTotalMovimento = movimentoExistente
      ? parseNumeroBR(movimentoExistente.custo_total)
      : parseNumeroBR(item.custo_total);
    if (Math.abs(quantidadeMovimento - quantidadeEntrada) > 0.000001 ||
        Math.abs(custoTotalMovimento - parseNumeroBR(item.custo_total)) > 0.009) {
      throw new Error(`Os dados da entrada pendente de ${item.item_nome} nao correspondem ao documento atual.`);
    }
    const movimento = movimentoExistente || criarMovimentoPendenteDocumentoCompra_(
      estoqueIdAlvo, documento.ID, item, quantidadeEntrada, custoUnitarioEntrada, saldoAnterior, saldoPosterior
    );
    const saldoAtual = parseNumeroBR(obterItemEstoque(estoqueIdAlvo)?.quantidade);
    let custoMedio;
    if (Math.abs(saldoAtual - saldoAnterior) <= 0.000001) {
      const custoAnterior = parseNumeroBR(alvo.custo_unitario || alvo.valor_unit);
      custoMedio = saldoPosterior > 0
        ? Number((((saldoAnterior * custoAnterior) + custoTotalMovimento) / saldoPosterior).toFixed(6))
        : custoUnitarioEntrada;
      updateById(ABA_ESTOQUE, 'ID', estoqueIdAlvo, {
        quantidade: saldoPosterior,
        custo_unitario: custoMedio,
        valor_unit: custoMedio
      }, ESTOQUE_SCHEMA);
    } else if (Math.abs(saldoAtual - saldoPosterior) > 0.000001) {
      throw new Error(`Saldo de ${item.item_nome} mudou durante a confirmacao. Revise o documento.`);
    } else {
      // O saldo ja foi aplicado em uma tentativa anterior interrompida.
      custoMedio = parseNumeroBR(alvo.custo_unitario || alvo.valor_unit);
    }
    marcarMovimentoEstoqueAplicadoDocumentoCompra_(movimento.ID);
    return { estoque_id: estoqueIdAlvo, movimento_id: movimento.ID, quantidade: quantidadeEntrada, custo_unitario: custoMedio };
  }

  const origemExistente = linhas.find(est =>
    String(est.origem_tipo || '').trim() === 'DOCUMENTO_COMPRA_ITEM' &&
    String(est.origem_id || '').trim() === String(item.ID || '').trim()
  );
  let estoqueId = String(origemExistente?.ID || '').trim();
  if (!estoqueId) {
    estoqueId = gerarId('EST');
    const novo = {
      ID: estoqueId,
      tipo: item.tipo,
      item: item.item_nome,
      unidade: normalizarTextoDocumentoCompra_(item.tipo) === 'MADEIRA' ? 'M3' : item.unidade,
      valor_unit: custoUnitarioEntrada,
      custo_unitario: custoUnitarioEntrada,
      preco_venda: '',
      ativo: true,
      criado_em: new Date(),
      quantidade: quantidadeEntrada,
      comprimento_cm: item.comprimento_cm || '',
      largura_cm: item.largura_cm || '',
      espessura_cm: item.espessura_cm || '',
      categoria: item.categoria,
      fornecedor: documento.fornecedor,
      pago_por: documento.pago_por,
      potencia: '',
      voltagem: '',
      comprado_em: documento.data_compra,
      data_pagamento: documento.pagamento_confirmado ? documento.data_pagamento : '',
      forma_pagamento: documento.forma_pagamento,
      parcelas: documento.parcelas,
      vida_util_mes: '',
      observacao: `Entrada pelo documento ${documento.ID}`,
      origem_tipo: 'DOCUMENTO_COMPRA_ITEM',
      origem_id: item.ID,
      op_id: ''
    };
    insert(ABA_ESTOQUE, novo, ESTOQUE_SCHEMA);
  }
  const movimento = movimentoExistente || {
    ID: gerarId('MVE'), status: 'APLICADO', estoque_id: estoqueId,
    tipo_movimento: 'ENTRADA_COMPRA', origem_tipo: 'DOCUMENTO_COMPRA', origem_id: documento.ID,
    origem_item_id: item.ID, quantidade: quantidadeEntrada, custo_unitario: custoUnitarioEntrada,
    custo_total: item.custo_total, saldo_anterior: 0, saldo_posterior: quantidadeEntrada,
    observacao: `Entrada de ${item.item_nome}`, ativo: true, criado_em: new Date(), aplicado_em: new Date()
  };
  if (!movimentoExistente) insert(ABA_MOVIMENTOS_ESTOQUE, movimento, MOVIMENTOS_ESTOQUE_SCHEMA);
  return { estoque_id: estoqueId, movimento_id: movimento.ID, quantidade: quantidadeEntrada, custo_unitario: custoUnitarioEntrada };
}

function validarCompatibilidadeEstoqueDocumentoCompra_(estoque, item) {
  const pares = [
    ['tipo', estoque.tipo, item.tipo],
    ['categoria', estoque.categoria, item.categoria],
    ['unidade', estoque.unidade, normalizarTextoDocumentoCompra_(item.tipo) === 'MADEIRA' ? 'M3' : item.unidade]
  ];
  pares.forEach(([campo, atual, novo]) => {
    if (normalizarTextoDocumentoCompra_(atual) !== normalizarTextoDocumentoCompra_(novo)) {
      throw new Error(`O ${campo} de ${item.item_nome} nao e compativel com o item de estoque selecionado.`);
    }
  });
  if (normalizarTextoDocumentoCompra_(item.tipo) === 'MADEIRA') {
    ['comprimento_cm', 'largura_cm', 'espessura_cm'].forEach(campo => {
      if (Math.abs(parseNumeroBR(estoque[campo]) - parseNumeroBR(item[campo])) > 0.000001) {
        throw new Error(`As dimensoes de ${item.item_nome} diferem do lote de madeira selecionado.`);
      }
    });
  }
}

function buscarMovimentoEstoqueDocumentoCompra_(itemId) {
  return listarMovimentosEstoqueDocumentoCompra_().find(mov =>
    String(mov.origem_item_id || '').trim() === String(itemId || '').trim()
  ) || null;
}

function listarMovimentosEstoqueDocumentoCompra_() {
  const sheet = getSheet(ABA_MOVIMENTOS_ESTOQUE);
  if (!sheet) return [];
  ensureSchema(sheet, MOVIMENTOS_ESTOQUE_SCHEMA);
  return rowsToObjects(sheet).filter(mov =>
    String(mov.origem_tipo || '').trim() === 'DOCUMENTO_COMPRA' &&
    String(mov.ativo).toLowerCase() === 'true'
  );
}

function criarMovimentoPendenteDocumentoCompra_(estoqueId, documentoId, item, quantidade, custoUnitario, saldoAnterior, saldoPosterior) {
  const movimento = {
    ID: gerarId('MVE'), status: 'PENDENTE', estoque_id: estoqueId,
    tipo_movimento: 'ENTRADA_COMPRA', origem_tipo: 'DOCUMENTO_COMPRA', origem_id: documentoId,
    origem_item_id: item.ID, quantidade, custo_unitario: custoUnitario,
    custo_total: item.custo_total, saldo_anterior: saldoAnterior, saldo_posterior: saldoPosterior,
    observacao: `Entrada de ${item.item_nome}`, ativo: true, criado_em: new Date(), aplicado_em: ''
  };
  insert(ABA_MOVIMENTOS_ESTOQUE, movimento, MOVIMENTOS_ESTOQUE_SCHEMA);
  return movimento;
}

function marcarMovimentoEstoqueAplicadoDocumentoCompra_(movimentoId) {
  updateById(ABA_MOVIMENTOS_ESTOQUE, 'ID', movimentoId, {
    status: 'APLICADO', aplicado_em: new Date()
  }, MOVIMENTOS_ESTOQUE_SCHEMA);
}

function descartarDocumentoCompra(id, ambiente) {
  return executarComAmbienteDocumentosCompra_(ambiente, () => {
    assertCanWrite('Descarte de documento de compra');
    const documento = obterDocumentoCompraPorId_(id);
    if (!documento) throw new Error('Documento de compra nao encontrado.');
    if (documento.status === 'CONFIRMADO') throw new Error('Documento confirmado nao pode ser descartado.');
    const itens = listarItensDocumentoCompra_(documento.ID);
    const possuiMovimento = itens.some(item => !!item.movimento_id);
    if (documento.compra_id_financeira || documento.pagamento_id_confirmado || possuiMovimento) {
      throw new Error('Este documento ja possui integracoes financeiras ou de estoque. Corrija a pendencia e conclua a confirmacao.');
    }
    const agora = new Date();
    updateById(ABA_DOCUMENTOS_COMPRA, 'ID', documento.ID, {
      status: 'DESCARTADO', descartado_em: agora, atualizado_em: agora
    }, DOCUMENTOS_COMPRA_SCHEMA);
    if (typeof moverArquivosDocumentoCompraAposDescarte_ === 'function') {
      moverArquivosDocumentoCompraAposDescarte_(documento);
    }
    return { ok: true, ID: documento.ID };
  });
}

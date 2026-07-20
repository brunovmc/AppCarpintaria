const DOCUMENTOS_COMPRA_DRIVE_ROOT_FOLDER_NAME = 'Documentos de Compra CarpintariaZizu';
const DOCUMENTOS_COMPRA_DRIVE_BATCH_DEFAULT = 5;
const DOCUMENTOS_COMPRA_DRIVE_TRIGGER_HANDLER = 'executarImportacaoDocumentosCompraDriveTrigger';
const DOCUMENTOS_COMPRA_DRIVE_FOLDER_PROPS = {
  raiz: 'DOCUMENTOS_COMPRA_DRIVE_ROOT_FOLDER_ID',
  entrada: 'DOCUMENTOS_COMPRA_DRIVE_ENTRADA_FOLDER_ID',
  processados: 'DOCUMENTOS_COMPRA_DRIVE_PROCESSADOS_FOLDER_ID',
  erros: 'DOCUMENTOS_COMPRA_DRIVE_ERROS_FOLDER_ID',
  descartados: 'DOCUMENTOS_COMPRA_DRIVE_DESCARTADOS_FOLDER_ID'
};

function configurarPastasDocumentosCompraDrive(pastaRaizId) {
  assertCanWrite('Configuracao do Drive para documentos de compra');
  return configurarPastasDocumentosCompraDrive_(pastaRaizId);
}

function configurarPastasDocumentosCompraDrive_(pastaRaizId) {
  const props = PropertiesService.getScriptProperties();
  const ambiente = obterAmbienteDocumentosCompraDrive_();
  const idInformado = String(pastaRaizId || '').trim();
  const idSalvo = String(props.getProperty(obterChavePastaDocumentosCompraDrive_('raiz', ambiente)) || '').trim();
  let raiz = null;
  for (const id of [idInformado, idSalvo].filter(Boolean)) {
    try {
      raiz = DriveApp.getFolderById(id);
      break;
    } catch (error) {
      raiz = null;
    }
  }
  const nomeRaiz = ambiente === 'dev'
    ? `${DOCUMENTOS_COMPRA_DRIVE_ROOT_FOLDER_NAME} - DEV`
    : DOCUMENTOS_COMPRA_DRIVE_ROOT_FOLDER_NAME;
  if (!raiz) raiz = DriveApp.createFolder(nomeRaiz);
  const pastas = {
    raiz,
    entrada: obterOuCriarSubpastaDocumentoCompraDrive_(raiz, 'Entrada'),
    processados: obterOuCriarSubpastaDocumentoCompraDrive_(raiz, 'Processados'),
    erros: obterOuCriarSubpastaDocumentoCompraDrive_(raiz, 'Erros'),
    descartados: obterOuCriarSubpastaDocumentoCompraDrive_(raiz, 'Descartados')
  };
  Object.keys(pastas).forEach(chave => props.setProperty(obterChavePastaDocumentosCompraDrive_(chave, ambiente), pastas[chave].getId()));
  return montarStatusPastasDocumentosCompraDrive_(pastas, ambiente);
}

function obterStatusDocumentosCompraDrive() {
  if (typeof assertCanRead === 'function') assertCanRead('Pastas de documentos de compra no Drive');
  const pastas = obterPastasDocumentosCompraDrive_(false);
  return pastas
    ? montarStatusPastasDocumentosCompraDrive_(pastas)
    : { configurado: false, gatilho_instalado: temTriggerDocumentosCompraDrive_() };
}

function salvarArquivosUploadDocumentoCompraDrive_(arquivos) {
  const pastas = obterPastasDocumentosCompraDrive_(true);
  const salvos = [];
  try {
    (arquivos || []).forEach(arquivo => {
      const blob = Utilities.newBlob(arquivo.bytes, arquivo.mime, arquivo.nome);
      const file = pastas.entrada.createFile(blob).setName(arquivo.nome);
      salvos.push({ id: file.getId(), url: file.getUrl(), file });
    });
    return salvos.map(item => ({ id: item.id, url: item.url }));
  } catch (error) {
    salvos.forEach(item => {
      try {
        moverArquivoDocumentoCompraDrive_(item.file, pastas.erros);
      } catch (moveError) {
        // A falha original deve prevalecer.
      }
    });
    throw error;
  }
}

function importarDocumentosCompraDriveParaInbox(limite, ambiente) {
  return executarComAmbienteDocumentosCompra_(ambiente, () => importarDocumentosCompraDriveParaInboxAtual_(limite));
}

function importarDocumentosCompraDriveParaInboxAtual_(limite) {
  assertCanWrite('Importacao de documentos de compra do Drive');
  const quantidade = Math.max(1, Math.min(20, Math.floor(Number(limite) || DOCUMENTOS_COMPRA_DRIVE_BATCH_DEFAULT)));
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('Ja existe uma importacao de documentos em andamento.');
  try {
    const pastas = obterPastasDocumentosCompraDrive_(true);
    const files = pastas.entrada.getFiles();
    const resultado = { processados: 0, adicionados: 0, ignorados: 0, erros: 0, limite: quantidade };
    while (files.hasNext() && resultado.processados < quantidade) {
      const file = files.next();
      resultado.processados++;
      if (buscarDocumentoCompraPorDriveId_(file.getId())) {
        tentarMoverArquivoDocumentoCompraDrive_(file, pastas.processados);
        resultado.ignorados++;
        continue;
      }
      try {
        const arquivo = normalizarArquivoDriveDocumentoCompra_(file);
        const existenteHash = buscarDocumentoCompraPorHash_(calcularHashCompostoDocumentoCompra_([arquivo]));
        if (existenteHash && existenteHash.status !== 'DESCARTADO') {
          tentarMoverArquivoDocumentoCompraDrive_(file, pastas.processados);
          resultado.ignorados++;
          continue;
        }
        const documento = criarRascunhoDocumentoCompraPorArquivos_([arquivo], {
          origemEntrada: 'DRIVE_FOLDER',
          driveIds: [file.getId()],
          urls: [file.getUrl()]
        });
        if (documento.status === 'ERRO') {
          tentarMoverArquivoDocumentoCompraDrive_(file, pastas.erros);
          resultado.erros++;
        } else {
          tentarMoverArquivoDocumentoCompraDrive_(file, pastas.processados);
          resultado.adicionados++;
        }
      } catch (error) {
        registrarErroArquivoDriveDocumentoCompra_(file, error);
        tentarMoverArquivoDocumentoCompraDrive_(file, pastas.erros);
        resultado.erros++;
      }
    }
    return { ...resultado, pastas: montarStatusPastasDocumentosCompraDrive_(pastas) };
  } finally {
    lock.releaseLock();
  }
}

function executarImportacaoDocumentosCompraDriveTrigger() {
  return executarComAmbienteDocumentosCompra_('prod', () => importarDocumentosCompraDriveParaInboxAtual_(DOCUMENTOS_COMPRA_DRIVE_BATCH_DEFAULT));
}

function instalarTriggerImportacaoDocumentosCompraDrive(minutos, ambiente) {
  assertCanWrite('Gatilho de importacao de documentos de compra');
  const env = String(ambiente || obterAmbienteDocumentosCompraDrive_()).trim().toLowerCase();
  if (env === 'dev') {
    throw new Error('A busca automatica usa somente a pasta de PROD. Em DEV, use o botao Buscar no Drive.');
  }
  const intervalo = Math.floor(Number(minutos) || 5);
  if (![1, 5, 10, 15, 30].includes(intervalo)) throw new Error('Intervalo invalido. Use 1, 5, 10, 15 ou 30 minutos.');
  removerTriggerImportacaoDocumentosCompraDrive_();
  ScriptApp.newTrigger(DOCUMENTOS_COMPRA_DRIVE_TRIGGER_HANDLER).timeBased().everyMinutes(intervalo).create();
  return { ok: true, minutos: intervalo };
}

function removerTriggerImportacaoDocumentosCompraDrive() {
  assertCanWrite('Remocao do gatilho de documentos de compra');
  return { ok: true, removidos: removerTriggerImportacaoDocumentosCompraDrive_() };
}

function removerTriggerImportacaoDocumentosCompraDrive_() {
  let removidos = 0;
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === DOCUMENTOS_COMPRA_DRIVE_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
      removidos++;
    }
  });
  return removidos;
}

function normalizarArquivoDriveDocumentoCompra_(file) {
  const mime = String(file.getMimeType() || '').trim().toLowerCase();
  if (!(mime.startsWith('image/') || mime === 'application/pdf')) throw new Error('Formato nao suportado. Use imagem ou PDF.');
  if (file.getSize() > DOCUMENTOS_COMPRA_MAX_BYTES_ARQUIVO) throw new Error('Arquivo maior que 10 MB.');
  const bytes = file.getBlob().getBytes();
  const base64 = Utilities.base64Encode(bytes);
  return {
    nome: normalizarNomeArquivoDocumentoCompra_(file.getName()),
    mime,
    bytes,
    base64,
    dataUrl: `data:${mime};base64,${base64}`,
    hash: calcularHashBytesDocumentoCompra_(bytes)
  };
}

function registrarErroArquivoDriveDocumentoCompra_(file, error) {
  const agora = new Date();
  insert(ABA_DOCUMENTOS_COMPRA, {
    ID: gerarId('DOCCOM'), status: 'ERRO', origem_entrada: 'DRIVE_FOLDER',
    arquivo_nomes_json: JSON.stringify([file.getName()]),
    arquivo_mimes_json: JSON.stringify([file.getMimeType()]),
    arquivo_drive_ids_json: JSON.stringify([file.getId()]),
    arquivo_urls_json: JSON.stringify([file.getUrl()]), arquivo_hash: '',
    tipo_documento: '', fornecedor: '', numero_documento: '', numero_pedido: '',
    data_compra: '', data_vencimento: '', subtotal: 0, frete: 0, desconto: 0, valor_total: 0,
    moeda: 'BRL', pago_por: '', forma_pagamento: '', parcelas: 1, recebido: false,
    pagamento_detectado: false, pagamento_confirmado: false, data_pagamento: '', valor_pago_confirmado: 0,
    observacao: '', dados_extraidos_json: '', confianca: 0, alertas_json: '[]',
    erro: error?.message || String(error), compra_id_financeira: '', pagamento_id_confirmado: '',
    client_request_id: '', ativo: true, criado_em: agora, atualizado_em: agora,
    confirmado_em: '', descartado_em: ''
  }, DOCUMENTOS_COMPRA_SCHEMA);
}

function buscarDocumentoCompraPorDriveId_(driveId) {
  const alvo = String(driveId || '').trim();
  const sheet = getSheet(ABA_DOCUMENTOS_COMPRA);
  if (!alvo || !sheet) return null;
  ensureSchema(sheet, DOCUMENTOS_COMPRA_SCHEMA);
  return rowsToObjects(sheet).find(row => {
    const ids = parseJsonDocumentoCompra_(row.arquivo_drive_ids_json, []);
    return ids.some(id => String(id || '').trim() === alvo);
  }) || null;
}

function moverArquivosDocumentoCompraAposConfirmacao_(documento) {
  moverArquivosDocumentoCompraPorStatus_(documento, 'processados');
}

function moverArquivosDocumentoCompraAposLeitura_(documento) {
  moverArquivosDocumentoCompraPorStatus_(documento, 'processados');
}

function moverArquivosDocumentoCompraAposDescarte_(documento) {
  moverArquivosDocumentoCompraPorStatus_(documento, 'descartados');
}

function moverArquivosDocumentoCompraAposErro_(documento) {
  moverArquivosDocumentoCompraPorStatus_(documento, 'erros');
}

function moverArquivosDocumentoCompraPorStatus_(documento, destino) {
  const ids = Array.isArray(documento?.arquivos_drive_ids)
    ? documento.arquivos_drive_ids
    : parseJsonDocumentoCompra_(documento?.arquivo_drive_ids_json, []);
  if (ids.length === 0) return;
  try {
    const pastas = obterPastasDocumentosCompraDrive_(false);
    if (!pastas?.[destino]) return;
    ids.filter(Boolean).forEach(id => {
      try {
        moverArquivoDocumentoCompraDrive_(DriveApp.getFileById(id), pastas[destino]);
      } catch (error) {
        console.log(JSON.stringify({ evento: 'documentos_compra.drive.movimento_falhou', arquivo_id: id, erro: String(error) }));
      }
    });
  } catch (error) {
    console.log(JSON.stringify({ evento: 'documentos_compra.drive.pastas_indisponiveis', erro: String(error) }));
  }
}

function moverArquivoDocumentoCompraDrive_(file, pasta) {
  file.moveTo(pasta);
}

function tentarMoverArquivoDocumentoCompraDrive_(file, pasta) {
  try {
    moverArquivoDocumentoCompraDrive_(file, pasta);
    return true;
  } catch (error) {
    console.log(JSON.stringify({
      evento: 'documentos_compra.drive.movimento_falhou',
      arquivo_id: String(file?.getId?.() || ''),
      erro: String(error)
    }));
    return false;
  }
}

function obterPastasDocumentosCompraDrive_(criarSeAusente) {
  const props = PropertiesService.getScriptProperties();
  const ambiente = obterAmbienteDocumentosCompraDrive_();
  const ids = {};
  for (const chave of Object.keys(DOCUMENTOS_COMPRA_DRIVE_FOLDER_PROPS)) {
    ids[chave] = String(props.getProperty(obterChavePastaDocumentosCompraDrive_(chave, ambiente)) || '').trim();
    if (!ids[chave]) return criarSeAusente ? pastasDocumentosCompraPorStatus_(configurarPastasDocumentosCompraDrive_()) : null;
  }
  try {
    const pastas = {};
    Object.keys(ids).forEach(chave => { pastas[chave] = DriveApp.getFolderById(ids[chave]); });
    return pastas;
  } catch (error) {
    return criarSeAusente ? pastasDocumentosCompraPorStatus_(configurarPastasDocumentosCompraDrive_()) : null;
  }
}

function pastasDocumentosCompraPorStatus_(status) {
  const pastas = {};
  Object.keys(status.pastas || {}).forEach(chave => { pastas[chave] = DriveApp.getFolderById(status.pastas[chave].id); });
  return pastas;
}

function obterAmbienteDocumentosCompraDrive_() {
  if (typeof getDbEnvironmentExecutionEffective_ === 'function') {
    return getDbEnvironmentExecutionEffective_();
  }
  if (typeof getUserDbEnvironment_ !== 'function') return 'prod';
  return String(getUserDbEnvironment_() || '').trim().toLowerCase() === 'dev' ? 'dev' : 'prod';
}

function obterChavePastaDocumentosCompraDrive_(chave, ambiente) {
  const base = DOCUMENTOS_COMPRA_DRIVE_FOLDER_PROPS[chave];
  return String(ambiente || '').trim().toLowerCase() === 'dev' ? `${base}_DEV` : base;
}

function obterOuCriarSubpastaDocumentoCompraDrive_(raiz, nome) {
  const encontradas = raiz.getFoldersByName(nome);
  return encontradas.hasNext() ? encontradas.next() : raiz.createFolder(nome);
}

function montarStatusPastasDocumentosCompraDrive_(pastas, ambiente) {
  const env = String(ambiente || obterAmbienteDocumentosCompraDrive_()).trim().toLowerCase() === 'dev' ? 'dev' : 'prod';
  const status = { configurado: true, ambiente: env, gatilho_instalado: temTriggerDocumentosCompraDrive_(), pastas: {} };
  Object.keys(pastas).forEach(chave => {
    status.pastas[chave] = { id: pastas[chave].getId(), nome: pastas[chave].getName(), url: pastas[chave].getUrl() };
  });
  return status;
}

function temTriggerDocumentosCompraDrive_() {
  return ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === DOCUMENTOS_COMPRA_DRIVE_TRIGGER_HANDLER);
}

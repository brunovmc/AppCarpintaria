const INBOX_RECEBIMENTOS_DRIVE_MAX_BYTES = 10 * 1024 * 1024;
const INBOX_RECEBIMENTOS_DRIVE_BATCH_DEFAULT = 5;
const INBOX_RECEBIMENTOS_DRIVE_TRIGGER_HANDLER = 'executarImportacaoRecebimentosDriveTrigger';
const INBOX_RECEBIMENTOS_DRIVE_FOLDER_PROPS = {
  raiz: 'INBOX_RECEBIMENTOS_DRIVE_ROOT_FOLDER_ID',
  entrada: 'INBOX_RECEBIMENTOS_DRIVE_ENTRADA_FOLDER_ID',
  processados: 'INBOX_RECEBIMENTOS_DRIVE_PROCESSADOS_FOLDER_ID',
  erros: 'INBOX_RECEBIMENTOS_DRIVE_ERROS_FOLDER_ID',
  descartados: 'INBOX_RECEBIMENTOS_DRIVE_DESCARTADOS_FOLDER_ID'
};

function configurarPastasInboxRecebimentosDrive(pastaRaizId) {
  assertCanWrite('Configuracao do Drive para recebimentos');
  return configurarPastasInboxRecebimentosDrive_(pastaRaizId);
}

function configurarPastasInboxRecebimentosDrive_(pastaRaizId) {
  const props = PropertiesService.getScriptProperties();
  const ambiente = obterAmbienteInboxRecebimentosDrive_();
  const pastasAnteriores = lerPastasComprovantesDriveSalvas_(
    INBOX_RECEBIMENTOS_DRIVE_FOLDER_PROPS,
    obterChavePastaInboxRecebimentosDrive_,
    ambiente
  );
  const raizComum = obterOuCriarRaizComprovantesDrive_(ambiente, pastaRaizId);
  const pastas = criarEstruturaDominioComprovantesDrive_(raizComum, 'recebimentos', pastasAnteriores);
  const migracao = migrarArquivosPastasComprovantesDrive_(pastasAnteriores, pastas, 'recebimentos');
  Object.keys(pastas).forEach(chave => {
    props.setProperty(obterChavePastaInboxRecebimentosDrive_(chave, ambiente), pastas[chave].getId());
  });
  return { ...montarStatusPastasInboxRecebimentos_(pastas, ambiente), migracao };
}

function obterStatusInboxRecebimentosDrive() {
  if (typeof assertCanRead === 'function') assertCanRead('Pastas de recebimentos no Drive');
  const pastas = obterPastasInboxRecebimentosDrive_(false);
  return pastas
    ? montarStatusPastasInboxRecebimentos_(pastas)
    : { configurado: false, gatilho_instalado: temTriggerInboxRecebimentosDrive_() };
}

function salvarArquivoUploadInboxRecebimentosDrive_(upload) {
  const pastas = obterPastasInboxRecebimentosDrive_(true);
  const arquivo = pastas.entrada.createFile(
    Utilities.newBlob(upload.bytes, upload.mime, upload.nome)
  ).setName(upload.nome);
  return { id: arquivo.getId(), url: arquivo.getUrl() };
}

function importarComprovantesRecebimentosDrive(limite, ambiente) {
  return executarComAmbienteInboxRecebimentos_(ambiente, () =>
    importarComprovantesRecebimentosDriveAtual_(limite)
  );
}

function importarComprovantesRecebimentosDriveAtual_(limite) {
  assertCanWrite('Importacao de recebimentos do Drive');
  const quantidade = Math.max(1, Math.min(20,
    Math.floor(Number(limite) || INBOX_RECEBIMENTOS_DRIVE_BATCH_DEFAULT)));
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    throw new Error('Ja existe uma busca de recebimentos em andamento. Tente novamente em instantes.');
  }
  try {
    const pastas = obterPastasInboxRecebimentosDrive_(true);
    const arquivos = pastas.entrada.getFiles();
    const resultado = { processados: 0, adicionados: 0, ignorados: 0, erros: 0, limite: quantidade };
    while (arquivos.hasNext() && resultado.processados < quantidade) {
      const arquivo = arquivos.next();
      resultado.processados++;
      if (buscarInboxRecebimentoPorArquivoDriveId_(arquivo.getId())) {
        tentarMoverArquivoInboxRecebimentos_(arquivo, pastas.processados);
        resultado.ignorados++;
        continue;
      }
      try {
        const processado = processarArquivoDriveInboxRecebimento_(arquivo);
        tentarMoverArquivoInboxRecebimentos_(arquivo, pastas.processados);
        if (processado?.reutilizado) resultado.ignorados++;
        else resultado.adicionados++;
      } catch (error) {
        try { registrarErroArquivoDriveInboxRecebimento_(arquivo, error); } catch (registroError) {
          console.log(JSON.stringify({
            evento: 'inbox_recebimentos.drive.registro_erro_falhou',
            arquivo_id: arquivo.getId(),
            erro: String(registroError)
          }));
        }
        tentarMoverArquivoInboxRecebimentos_(arquivo, pastas.erros);
        resultado.erros++;
      }
    }
    return { ...resultado, pastas: montarStatusPastasInboxRecebimentos_(pastas) };
  } finally {
    lock.releaseLock();
  }
}

function executarImportacaoRecebimentosDriveTrigger() {
  return executarComAmbienteInboxRecebimentos_('prod', () =>
    importarComprovantesRecebimentosDriveAtual_(INBOX_RECEBIMENTOS_DRIVE_BATCH_DEFAULT)
  );
}

function instalarTriggerImportacaoRecebimentosDrive(minutos, ambiente) {
  assertCanWrite('Gatilho de importacao de recebimentos');
  const env = String(ambiente || obterAmbienteInboxRecebimentosDrive_()).trim().toLowerCase();
  if (env === 'dev') {
    throw new Error('A busca automatica usa somente a pasta de PROD. Em DEV, use Buscar no Drive.');
  }
  const intervalo = Math.floor(Number(minutos) || 5);
  if (![1, 5, 10, 15, 30].includes(intervalo)) {
    throw new Error('Intervalo invalido. Use 1, 5, 10, 15 ou 30 minutos.');
  }
  removerTriggerImportacaoRecebimentosDrive_();
  ScriptApp.newTrigger(INBOX_RECEBIMENTOS_DRIVE_TRIGGER_HANDLER)
    .timeBased().everyMinutes(intervalo).create();
  return { ok: true, minutos: intervalo };
}

function removerTriggerImportacaoRecebimentosDrive() {
  assertCanWrite('Gatilho de importacao de recebimentos');
  return { ok: true, removidos: removerTriggerImportacaoRecebimentosDrive_() };
}

function removerTriggerImportacaoRecebimentosDrive_() {
  let removidos = 0;
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === INBOX_RECEBIMENTOS_DRIVE_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
      removidos++;
    }
  });
  return removidos;
}

function processarArquivoDriveInboxRecebimento_(arquivo) {
  const mime = String(arquivo.getMimeType() || '').trim().toLowerCase();
  if (!(mime.startsWith('image/') || mime === 'application/pdf')) {
    throw new Error('Formato nao suportado. Use imagem ou PDF.');
  }
  if (arquivo.getSize() > INBOX_RECEBIMENTOS_DRIVE_MAX_BYTES) {
    throw new Error('Arquivo maior que 10 MB.');
  }
  const bytes = arquivo.getBlob().getBytes();
  const hash = hashBytesRecebimento_(bytes);
  const existente = buscarInboxRecebimentoPorHash_(hash);
  if (existente) return { ...existente, reutilizado: true };
  const extraido = extrairRecebimentoComOpenAI_({
    base64: Utilities.base64Encode(bytes),
    mime,
    nome: arquivo.getName()
  });
  const rascunho = normalizarRecebimentoExtraido_(extraido);
  const alertas = [...rascunho.alertas];
  if (rascunho.referencia_transacao) {
    const referenciaDuplicada = buscarInboxRecebimentoPorReferencia_(rascunho.referencia_transacao);
    if (referenciaDuplicada) {
      alertas.push(`Referencia da transacao ja encontrada no comprovante ${referenciaDuplicada.ID}.`);
    }
  }
  const agora = new Date();
  const linha = {
    ID: gerarId('IBXREC'),
    status: 'PENDENTE',
    origem_tipo: 'DRIVE_FOLDER',
    arquivo_nome: arquivo.getName(),
    arquivo_mime: mime,
    arquivo_drive_id: arquivo.getId(),
    arquivo_url: arquivo.getUrl(),
    arquivo_hash: hash,
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
    erro: '',
    criado_em: agora,
    atualizado_em: agora,
    confirmado_em: '',
    conciliado_em: '',
    descartado_em: ''
  };
  if (!insert(ABA_INBOX_RECEBIMENTOS, linha, INBOX_RECEBIMENTOS_SCHEMA)) {
    throw new Error('Nao foi possivel salvar o recebimento importado.');
  }
  return normalizarLinhaInboxRecebimento_(linha);
}

function registrarErroArquivoDriveInboxRecebimento_(arquivo, error) {
  const agora = new Date();
  const linha = {
    ID: gerarId('IBXREC'), status: 'ERRO', origem_tipo: 'DRIVE_FOLDER',
    arquivo_nome: arquivo.getName(), arquivo_mime: arquivo.getMimeType(),
    arquivo_drive_id: arquivo.getId(), arquivo_url: arquivo.getUrl(), arquivo_hash: '',
    referencia_transacao: '', pagador_nome: '', banco_pagador: '', recebido_por: '',
    valor_total: 0, data_recebimento: '', forma_pagamento: '', descricao: '', observacao: '',
    dados_extraidos_json: '', confianca: 0, alertas_json: '',
    erro: error?.message || String(error), criado_em: agora, atualizado_em: agora,
    confirmado_em: '', conciliado_em: '', descartado_em: ''
  };
  if (!insert(ABA_INBOX_RECEBIMENTOS, linha, INBOX_RECEBIMENTOS_SCHEMA)) {
    throw new Error('Nao foi possivel registrar o erro de importacao do recebimento.');
  }
}

function buscarInboxRecebimentoPorArquivoDriveId_(arquivoId) {
  const alvo = String(arquivoId || '').trim();
  const sheet = getSheet(ABA_INBOX_RECEBIMENTOS);
  if (!alvo || !sheet) return null;
  ensureSchema(sheet, INBOX_RECEBIMENTOS_SCHEMA);
  const item = rowsToObjects(sheet).find(row => String(row.arquivo_drive_id || '').trim() === alvo);
  return item ? normalizarLinhaInboxRecebimento_(item) : null;
}

function moverArquivoInboxRecebimentoAposLeitura_(item) {
  moverArquivoInboxRecebimentoPorItem_(item, 'processados');
}

function moverArquivoInboxRecebimentoAposConfirmacao_(item) {
  moverArquivoInboxRecebimentoPorItem_(item, 'processados');
}

function moverArquivoInboxRecebimentoAposErro_(item) {
  moverArquivoInboxRecebimentoPorItem_(item, 'erros');
}

function moverArquivoInboxRecebimentoAposDescarte_(item) {
  moverArquivoInboxRecebimentoPorItem_(item, 'descartados');
}

function moverArquivoInboxRecebimentoPorItem_(item, destino) {
  if (!item?.arquivo_drive_id) return;
  try {
    const pastas = obterPastasInboxRecebimentosDrive_(false);
    if (pastas?.[destino]) DriveApp.getFileById(item.arquivo_drive_id).moveTo(pastas[destino]);
  } catch (error) {
    console.log(JSON.stringify({
      evento: 'inbox_recebimentos.drive.movimento_falhou',
      arquivo_id: String(item?.arquivo_drive_id || ''),
      erro: String(error)
    }));
  }
}

function tentarMoverArquivoInboxRecebimentos_(arquivo, pasta) {
  try { arquivo.moveTo(pasta); return true; } catch (error) {
    console.log(JSON.stringify({
      evento: 'inbox_recebimentos.drive.movimento_falhou',
      arquivo_id: String(arquivo?.getId?.() || ''),
      erro: String(error)
    }));
    return false;
  }
}

function obterPastasInboxRecebimentosDrive_(criarSeAusente) {
  const props = PropertiesService.getScriptProperties();
  const ambiente = obterAmbienteInboxRecebimentosDrive_();
  const ids = {};
  for (const chave of Object.keys(INBOX_RECEBIMENTOS_DRIVE_FOLDER_PROPS)) {
    ids[chave] = String(props.getProperty(
      obterChavePastaInboxRecebimentosDrive_(chave, ambiente)
    ) || '').trim();
    if (!ids[chave]) {
      return criarSeAusente
        ? pastasStatusParaObjetosRecebimentos_(configurarPastasInboxRecebimentosDrive_())
        : null;
    }
  }
  try {
    const pastas = {};
    Object.keys(ids).forEach(chave => { pastas[chave] = DriveApp.getFolderById(ids[chave]); });
    if (estruturaDominioComprovantesDriveValida_(pastas, 'recebimentos')) return pastas;
    return criarSeAusente
      ? pastasStatusParaObjetosRecebimentos_(configurarPastasInboxRecebimentosDrive_())
      : null;
  } catch (error) {
    return criarSeAusente
      ? pastasStatusParaObjetosRecebimentos_(configurarPastasInboxRecebimentosDrive_())
      : null;
  }
}

function pastasStatusParaObjetosRecebimentos_(status) {
  const pastas = {};
  Object.keys(status.pastas || {}).forEach(chave => {
    pastas[chave] = DriveApp.getFolderById(status.pastas[chave].id);
  });
  return pastas;
}

function montarStatusPastasInboxRecebimentos_(pastas, ambiente) {
  const status = {
    configurado: true,
    ambiente: String(ambiente || obterAmbienteInboxRecebimentosDrive_()).trim().toLowerCase(),
    gatilho_instalado: temTriggerInboxRecebimentosDrive_(),
    pastas: {}
  };
  Object.keys(pastas).forEach(chave => {
    status.pastas[chave] = {
      id: pastas[chave].getId(),
      nome: pastas[chave].getName(),
      url: pastas[chave].getUrl()
    };
  });
  return status;
}

function temTriggerInboxRecebimentosDrive_() {
  return ScriptApp.getProjectTriggers().some(trigger =>
    trigger.getHandlerFunction() === INBOX_RECEBIMENTOS_DRIVE_TRIGGER_HANDLER
  );
}

function obterAmbienteInboxRecebimentosDrive_() {
  if (typeof getUserDbEnvironment_ === 'function') {
    return String(getUserDbEnvironment_() || '').trim().toLowerCase() === 'dev' ? 'dev' : 'prod';
  }
  return 'prod';
}

function obterChavePastaInboxRecebimentosDrive_(chave, ambiente) {
  const base = INBOX_RECEBIMENTOS_DRIVE_FOLDER_PROPS[chave];
  return String(ambiente || '').trim().toLowerCase() === 'dev' ? `${base}_DEV` : base;
}

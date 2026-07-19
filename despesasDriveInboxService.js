const INBOX_DESPESAS_DRIVE_ROOT_FOLDER_NAME = 'Comprovantes CarpintariaZizu';
const INBOX_DESPESAS_DRIVE_MAX_BYTES = 10 * 1024 * 1024;
const INBOX_DESPESAS_DRIVE_BATCH_DEFAULT = 5;
const INBOX_DESPESAS_DRIVE_TRIGGER_HANDLER = 'executarImportacaoComprovantesDriveTrigger';
const INBOX_DESPESAS_DRIVE_FOLDER_PROPS = {
  raiz: 'INBOX_DESPESAS_DRIVE_ROOT_FOLDER_ID',
  entrada: 'INBOX_DESPESAS_DRIVE_ENTRADA_FOLDER_ID',
  processados: 'INBOX_DESPESAS_DRIVE_PROCESSADOS_FOLDER_ID',
  erros: 'INBOX_DESPESAS_DRIVE_ERROS_FOLDER_ID',
  descartados: 'INBOX_DESPESAS_DRIVE_DESCARTADOS_FOLDER_ID'
};

function configurarPastasInboxDespesasDrive(pastaRaizId) {
  assertCanWrite('Configuracao do Drive para comprovantes');
  return configurarPastasInboxDespesasDrive_(pastaRaizId);
}

function configurarPastasInboxDespesasDrive_(pastaRaizId) {
  const props = PropertiesService.getScriptProperties();
  const raizIdInformada = String(pastaRaizId || '').trim();
  const raizIdSalva = String(props.getProperty(INBOX_DESPESAS_DRIVE_FOLDER_PROPS.raiz) || '').trim();
  let raiz;
  if (raizIdInformada) {
    raiz = DriveApp.getFolderById(raizIdInformada);
  } else if (raizIdSalva) {
    try {
      raiz = DriveApp.getFolderById(raizIdSalva);
    } catch (error) {
      raiz = null;
    }
  }
  if (!raiz) raiz = DriveApp.createFolder(INBOX_DESPESAS_DRIVE_ROOT_FOLDER_NAME);

  const pastas = {
    raiz,
    entrada: obterOuCriarSubpastaInboxDrive_(raiz, 'Entrada'),
    processados: obterOuCriarSubpastaInboxDrive_(raiz, 'Processados'),
    erros: obterOuCriarSubpastaInboxDrive_(raiz, 'Erros'),
    descartados: obterOuCriarSubpastaInboxDrive_(raiz, 'Descartados')
  };
  Object.keys(pastas).forEach(chave => {
    props.setProperty(INBOX_DESPESAS_DRIVE_FOLDER_PROPS[chave], pastas[chave].getId());
  });
  return montarStatusPastasInboxDrive_(pastas);
}

function obterStatusInboxDespesasDrive() {
  if (typeof assertCanRead === 'function') assertCanRead('Pastas de comprovantes no Drive');
  const pastas = obterPastasInboxDespesasDrive_(false);
  return pastas ? montarStatusPastasInboxDrive_(pastas) : { configurado: false, gatilho_instalado: temTriggerInboxDespesasDrive_() };
}

function importarComprovantesDriveParaInbox(limite, ambiente) {
  return executarComAmbienteInboxDespesas_(ambiente, () => importarComprovantesDriveParaInboxNoAmbienteAtual_(limite));
}

function importarComprovantesDriveParaInboxNoAmbienteAtual_(limite) {
  assertCanWrite('Importacao de comprovantes do Drive');
  const quantidade = Math.max(1, Math.min(20, Math.floor(Number(limite) || INBOX_DESPESAS_DRIVE_BATCH_DEFAULT)));
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('Ja existe uma busca de comprovantes em andamento. Tente novamente em instantes.');
  try {
    const pastas = obterPastasInboxDespesasDrive_(true);
    const arquivos = pastas.entrada.getFiles();
    const resultado = { processados: 0, adicionados: 0, ignorados: 0, erros: 0, limite: quantidade };
    while (arquivos.hasNext() && resultado.processados < quantidade) {
      const arquivo = arquivos.next();
      resultado.processados++;
      const existente = buscarInboxDespesaPorArquivoDriveId_(arquivo.getId());
      if (existente) {
        resultado.ignorados++;
        continue;
      }
      try {
        processarArquivoDriveInboxDespesa_(arquivo);
        resultado.adicionados++;
      } catch (error) {
        registrarErroArquivoDriveInboxDespesa_(arquivo, error);
        moverArquivoInboxDrive_(arquivo, pastas.erros);
        resultado.erros++;
      }
    }
    return { ...resultado, pastas: montarStatusPastasInboxDrive_(pastas) };
  } finally {
    lock.releaseLock();
  }
}

function executarImportacaoComprovantesDriveTrigger() {
  return executarComAmbienteInboxDespesas_('prod', () => importarComprovantesDriveParaInboxNoAmbienteAtual_(INBOX_DESPESAS_DRIVE_BATCH_DEFAULT));
}

function instalarTriggerImportacaoComprovantesDrive(minutos) {
  assertCanWrite('Gatilho de importacao de comprovantes');
  const intervalo = Math.floor(Number(minutos) || 5);
  if (![1, 5, 10, 15, 30].includes(intervalo)) throw new Error('Intervalo invalido. Use 1, 5, 10, 15 ou 30 minutos.');
  removerTriggerImportacaoComprovantesDrive_();
  ScriptApp.newTrigger(INBOX_DESPESAS_DRIVE_TRIGGER_HANDLER).timeBased().everyMinutes(intervalo).create();
  return { ok: true, minutos: intervalo };
}

function removerTriggerImportacaoComprovantesDrive() {
  assertCanWrite('Gatilho de importacao de comprovantes');
  return { ok: true, removidos: removerTriggerImportacaoComprovantesDrive_() };
}

function removerTriggerImportacaoComprovantesDrive_() {
  let removidos = 0;
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === INBOX_DESPESAS_DRIVE_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
      removidos++;
    }
  });
  return removidos;
}

function processarArquivoDriveInboxDespesa_(arquivo) {
  const mime = String(arquivo.getMimeType() || '').toLowerCase();
  if (!(mime.startsWith('image/') || mime === 'application/pdf')) throw new Error('Formato nao suportado. Use imagem ou PDF.');
  if (arquivo.getSize() > INBOX_DESPESAS_DRIVE_MAX_BYTES) throw new Error('Arquivo maior que 10 MB.');
  const blob = arquivo.getBlob();
  const bytes = blob.getBytes();
  const base64 = Utilities.base64Encode(bytes);
  const hash = calcularHashArquivoInboxDrive_(bytes);
  const extraido = extrairDespesaComOpenAI_({ base64, mime, nome: arquivo.getName() });
  const rascunho = normalizarDespesaExtraidaInbox_(extraido);
  const agora = new Date();
  const linha = {
    ID: gerarId('IBXDESP'), status: 'PENDENTE', origem_tipo: 'DRIVE_FOLDER',
    arquivo_nome: arquivo.getName(), arquivo_mime: mime, arquivo_drive_id: arquivo.getId(),
    arquivo_url: arquivo.getUrl(), imagem_hash: hash, descricao: rascunho.descricao,
    categoria: rascunho.categoria, fornecedor: rascunho.fornecedor, pago_por: rascunho.pago_por,
    valor_total: rascunho.valor_total, data_competencia: rascunho.data_competencia,
    data_vencimento: rascunho.data_vencimento, data_pagamento: rascunho.data_pagamento,
    forma_pagamento: rascunho.forma_pagamento, parcelas: rascunho.parcelas,
    observacao: rascunho.observacao, dados_extraidos_json: JSON.stringify(extraido || {}),
    confianca: rascunho.confianca, alertas_json: serializarListaInboxDespesa_(rascunho.alertas),
    erro: '', despesa_id_confirmada: '', criado_em: agora, atualizado_em: agora,
    confirmado_em: '', descartado_em: ''
  };
  insert(ABA_INBOX_DESPESAS, linha, INBOX_DESPESAS_SCHEMA);
  return linha;
}

function registrarErroArquivoDriveInboxDespesa_(arquivo, error) {
  const agora = new Date();
  insert(ABA_INBOX_DESPESAS, {
    ID: gerarId('IBXDESP'), status: 'ERRO', origem_tipo: 'DRIVE_FOLDER',
    arquivo_nome: arquivo.getName(), arquivo_mime: arquivo.getMimeType(), arquivo_drive_id: arquivo.getId(),
    arquivo_url: arquivo.getUrl(), imagem_hash: '', descricao: '', categoria: '', fornecedor: '', pago_por: '',
    valor_total: 0, data_competencia: '', data_vencimento: '', data_pagamento: '', forma_pagamento: '',
    parcelas: 1, observacao: '', dados_extraidos_json: '', confianca: 0, alertas_json: '[]',
    erro: error?.message || String(error), despesa_id_confirmada: '', criado_em: agora, atualizado_em: agora,
    confirmado_em: '', descartado_em: ''
  }, INBOX_DESPESAS_SCHEMA);
}

function buscarInboxDespesaPorArquivoDriveId_(arquivoId) {
  const alvo = String(arquivoId || '').trim();
  const sheet = getSheet(ABA_INBOX_DESPESAS);
  if (!alvo || !sheet) return null;
  ensureSchema(sheet, INBOX_DESPESAS_SCHEMA);
  return rowsToObjects(sheet).map(normalizarLinhaInboxDespesa_).find(item => item.arquivo_drive_id === alvo) || null;
}

function moverArquivoInboxDriveAposConfirmacao_(item) {
  moverArquivoInboxDrivePorItem_(item, 'processados');
}

function moverArquivoInboxDriveAposDescarte_(item) {
  moverArquivoInboxDrivePorItem_(item, 'descartados');
}

function moverArquivoInboxDrivePorItem_(item, destino) {
  if (String(item?.origem_tipo || '').toUpperCase() !== 'DRIVE_FOLDER' || !item?.arquivo_drive_id) return;
  try {
    const pastas = obterPastasInboxDespesasDrive_(false);
    if (pastas?.[destino]) moverArquivoInboxDrive_(DriveApp.getFileById(item.arquivo_drive_id), pastas[destino]);
  } catch (error) {
    console.log(JSON.stringify({ evento: 'inbox_despesas.drive.movimento_falhou', arquivo_id: item.arquivo_drive_id, erro: String(error) }));
  }
}

function moverArquivoInboxDrive_(arquivo, pasta) {
  arquivo.moveTo(pasta);
}

function obterPastasInboxDespesasDrive_(criarSeAusente) {
  const props = PropertiesService.getScriptProperties();
  const ids = {};
  for (const chave of Object.keys(INBOX_DESPESAS_DRIVE_FOLDER_PROPS)) {
    ids[chave] = String(props.getProperty(INBOX_DESPESAS_DRIVE_FOLDER_PROPS[chave]) || '').trim();
    if (!ids[chave]) return criarSeAusente ? pastasDeStatusParaDrive_(configurarPastasInboxDespesasDrive_()) : null;
  }
  try {
    const pastas = {};
    Object.keys(ids).forEach(chave => { pastas[chave] = DriveApp.getFolderById(ids[chave]); });
    return pastas;
  } catch (error) {
    return criarSeAusente ? pastasDeStatusParaDrive_(configurarPastasInboxDespesasDrive_()) : null;
  }
}

function pastasDeStatusParaDrive_(status) {
  const pastas = {};
  Object.keys(status.pastas || {}).forEach(chave => { pastas[chave] = DriveApp.getFolderById(status.pastas[chave].id); });
  return pastas;
}

function obterOuCriarSubpastaInboxDrive_(raiz, nome) {
  const encontradas = raiz.getFoldersByName(nome);
  return encontradas.hasNext() ? encontradas.next() : raiz.createFolder(nome);
}

function montarStatusPastasInboxDrive_(pastas) {
  const status = { configurado: true, gatilho_instalado: temTriggerInboxDespesasDrive_(), pastas: {} };
  Object.keys(pastas).forEach(chave => {
    status.pastas[chave] = { id: pastas[chave].getId(), nome: pastas[chave].getName(), url: pastas[chave].getUrl() };
  });
  return status;
}

function temTriggerInboxDespesasDrive_() {
  return ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === INBOX_DESPESAS_DRIVE_TRIGGER_HANDLER);
}

function calcularHashArquivoInboxDrive_(bytes) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)
    .map(valor => (`0${(valor & 255).toString(16)}`).slice(-2)).join('');
}

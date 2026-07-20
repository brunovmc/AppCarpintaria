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
  const ambiente = obterAmbienteInboxDespesasDrive_();
  const pastasAnteriores = lerPastasComprovantesDriveSalvas_(
    INBOX_DESPESAS_DRIVE_FOLDER_PROPS,
    obterChavePastaInboxDespesasDrive_,
    ambiente
  );
  const raizComum = obterOuCriarRaizComprovantesDrive_(ambiente, pastaRaizId);
  const pastas = criarEstruturaDominioComprovantesDrive_(raizComum, 'despesas', pastasAnteriores);
  const migracao = migrarArquivosPastasComprovantesDrive_(pastasAnteriores, pastas, 'despesas');
  Object.keys(pastas).forEach(chave => {
    props.setProperty(obterChavePastaInboxDespesasDrive_(chave, ambiente), pastas[chave].getId());
  });
  return { ...montarStatusPastasInboxDrive_(pastas, ambiente), migracao };
}

function obterStatusInboxDespesasDrive() {
  if (typeof assertCanRead === 'function') assertCanRead('Pastas de comprovantes no Drive');
  const pastas = obterPastasInboxDespesasDrive_(false);
  return pastas ? montarStatusPastasInboxDrive_(pastas) : { configurado: false, gatilho_instalado: temTriggerInboxDespesasDrive_() };
}

function salvarArquivoUploadInboxDespesasDrive_(upload) {
  const pastas = obterPastasInboxDespesasDrive_(true);
  const blob = Utilities.newBlob(upload.bytes, upload.mime, upload.nome);
  const file = pastas.entrada.createFile(blob).setName(upload.nome);
  return { id: file.getId(), url: file.getUrl() };
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
        tentarMoverArquivoInboxDrive_(arquivo, pastas.processados);
        resultado.ignorados++;
        continue;
      }
      try {
        const processado = processarArquivoDriveInboxDespesa_(arquivo);
        tentarMoverArquivoInboxDrive_(arquivo, pastas.processados);
        if (processado?.reutilizado) resultado.ignorados++;
        else resultado.adicionados++;
      } catch (error) {
        try {
          registrarErroArquivoDriveInboxDespesa_(arquivo, error);
        } catch (registroError) {
          console.log(JSON.stringify({
            evento: 'inbox_despesas.drive.registro_erro_falhou',
            arquivo_id: arquivo.getId(),
            erro: String(registroError)
          }));
        }
        tentarMoverArquivoInboxDrive_(arquivo, pastas.erros);
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

function instalarTriggerImportacaoComprovantesDrive(minutos, ambiente) {
  assertCanWrite('Gatilho de importacao de comprovantes');
  const env = String(ambiente || obterAmbienteInboxDespesasDrive_()).trim().toLowerCase();
  if (env === 'dev') {
    throw new Error('A busca automatica usa somente a pasta de PROD. Em DEV, use o botao Buscar no Drive.');
  }
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
  const existenteHash = buscarInboxDespesaPorHash_(hash);
  if (existenteHash) return { ...existenteHash, reutilizado: true };
  const extraido = extrairDespesaComOpenAI_({ base64, mime, nome: arquivo.getName() });
  const rascunho = normalizarDespesaExtraidaInbox_(extraido);
  const agora = new Date();
  const linha = {
    ID: gerarId('IBXDESP'), status: 'PENDENTE', classificacao: 'NAO_CLASSIFICADO', origem_tipo: 'DRIVE_FOLDER',
    arquivo_nome: arquivo.getName(), arquivo_mime: mime, arquivo_drive_id: arquivo.getId(),
    arquivo_url: arquivo.getUrl(), imagem_hash: hash, referencia_transacao: rascunho.referencia_transacao,
    descricao: rascunho.descricao,
    categoria: rascunho.categoria, fornecedor: rascunho.fornecedor, pago_por: rascunho.pago_por,
    valor_total: rascunho.valor_total, data_competencia: rascunho.data_competencia,
    data_vencimento: rascunho.data_vencimento, data_pagamento: rascunho.data_pagamento,
    forma_pagamento: rascunho.forma_pagamento, parcelas: rascunho.parcelas,
    observacao: rascunho.observacao, dados_extraidos_json: JSON.stringify(extraido || {}),
    confianca: rascunho.confianca, alertas_json: serializarListaInboxDespesa_(rascunho.alertas),
    erro: '', despesa_id_confirmada: '', criado_em: agora, atualizado_em: agora,
    confirmado_em: '', conciliado_em: '', descartado_em: ''
  };
  const inseriu = insert(ABA_INBOX_DESPESAS, linha, INBOX_DESPESAS_SCHEMA);
  if (!inseriu) throw new Error('Nao foi possivel salvar o comprovante importado na Inbox.');
  return linha;
}

function registrarErroArquivoDriveInboxDespesa_(arquivo, error) {
  const agora = new Date();
  const inseriu = insert(ABA_INBOX_DESPESAS, {
    ID: gerarId('IBXDESP'), status: 'ERRO', classificacao: 'NAO_CLASSIFICADO', origem_tipo: 'DRIVE_FOLDER',
    arquivo_nome: arquivo.getName(), arquivo_mime: arquivo.getMimeType(), arquivo_drive_id: arquivo.getId(),
    arquivo_url: arquivo.getUrl(), imagem_hash: '', referencia_transacao: '', descricao: '', categoria: '', fornecedor: '', pago_por: '',
    valor_total: 0, data_competencia: '', data_vencimento: '', data_pagamento: '', forma_pagamento: '',
    parcelas: 1, observacao: '', dados_extraidos_json: '', confianca: 0, alertas_json: '[]',
    erro: error?.message || String(error), despesa_id_confirmada: '', criado_em: agora, atualizado_em: agora,
    confirmado_em: '', conciliado_em: '', descartado_em: ''
  }, INBOX_DESPESAS_SCHEMA);
  if (!inseriu) throw new Error('Nao foi possivel registrar o erro de importacao na Inbox.');
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

function moverArquivoInboxDriveAposLeitura_(item) {
  moverArquivoInboxDrivePorItem_(item, 'processados');
}

function moverArquivoInboxDriveAposErro_(item) {
  moverArquivoInboxDrivePorItem_(item, 'erros');
}

function moverArquivoInboxDriveAposDescarte_(item) {
  moverArquivoInboxDrivePorItem_(item, 'descartados');
}

function moverArquivoInboxDrivePorItem_(item, destino) {
  if (!item?.arquivo_drive_id) return;
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

function tentarMoverArquivoInboxDrive_(arquivo, pasta) {
  try {
    moverArquivoInboxDrive_(arquivo, pasta);
    return true;
  } catch (error) {
    console.log(JSON.stringify({
      evento: 'inbox_despesas.drive.movimento_falhou',
      arquivo_id: String(arquivo?.getId?.() || ''),
      erro: String(error)
    }));
    return false;
  }
}

function obterPastasInboxDespesasDrive_(criarSeAusente) {
  const props = PropertiesService.getScriptProperties();
  const ambiente = obterAmbienteInboxDespesasDrive_();
  const ids = {};
  for (const chave of Object.keys(INBOX_DESPESAS_DRIVE_FOLDER_PROPS)) {
    ids[chave] = String(props.getProperty(obterChavePastaInboxDespesasDrive_(chave, ambiente)) || '').trim();
    if (!ids[chave]) return criarSeAusente ? pastasDeStatusParaDrive_(configurarPastasInboxDespesasDrive_()) : null;
  }
  try {
    const pastas = {};
    Object.keys(ids).forEach(chave => { pastas[chave] = DriveApp.getFolderById(ids[chave]); });
    if (estruturaDominioComprovantesDriveValida_(pastas, 'despesas')) return pastas;
    return criarSeAusente ? pastasDeStatusParaDrive_(configurarPastasInboxDespesasDrive_()) : null;
  } catch (error) {
    return criarSeAusente ? pastasDeStatusParaDrive_(configurarPastasInboxDespesasDrive_()) : null;
  }
}

function pastasDeStatusParaDrive_(status) {
  const pastas = {};
  Object.keys(status.pastas || {}).forEach(chave => { pastas[chave] = DriveApp.getFolderById(status.pastas[chave].id); });
  return pastas;
}

function montarStatusPastasInboxDrive_(pastas, ambiente) {
  const status = {
    configurado: true,
    ambiente: String(ambiente || obterAmbienteInboxDespesasDrive_()).trim().toLowerCase(),
    gatilho_instalado: temTriggerInboxDespesasDrive_(),
    pastas: {}
  };
  Object.keys(pastas).forEach(chave => {
    status.pastas[chave] = { id: pastas[chave].getId(), nome: pastas[chave].getName(), url: pastas[chave].getUrl() };
  });
  return status;
}

function temTriggerInboxDespesasDrive_() {
  return ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === INBOX_DESPESAS_DRIVE_TRIGGER_HANDLER);
}

function obterAmbienteInboxDespesasDrive_() {
  if (typeof getUserDbEnvironment_ === 'function') {
    return String(getUserDbEnvironment_() || '').trim().toLowerCase() === 'dev' ? 'dev' : 'prod';
  }
  return 'prod';
}

function obterChavePastaInboxDespesasDrive_(chave, ambiente) {
  const base = INBOX_DESPESAS_DRIVE_FOLDER_PROPS[chave];
  return String(ambiente || '').trim().toLowerCase() === 'dev' ? `${base}_DEV` : base;
}

function calcularHashArquivoInboxDrive_(bytes) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)
    .map(valor => (`0${(valor & 255).toString(16)}`).slice(-2)).join('');
}

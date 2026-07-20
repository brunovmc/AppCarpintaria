const COMPROVANTES_DRIVE_ROOT_FOLDER_NAME = 'Comprovantes CarpintariaZizu';
const COMPROVANTES_DRIVE_COMMON_ROOT_PROP = 'COMPROVANTES_DRIVE_ROOT_FOLDER_ID';
const COMPROVANTES_DRIVE_STRUCTURE_VERSION = '2';
const COMPROVANTES_DRIVE_STRUCTURE_VERSION_PROP = 'COMPROVANTES_DRIVE_STRUCTURE_VERSION';
const COMPROVANTES_DRIVE_DOMAIN_FOLDER_NAMES = {
  despesas: 'Despesas',
  recebimentos: 'Recebimentos'
};
const COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES = {
  entrada: 'Entrada',
  processados: 'Processados',
  erros: 'Erros',
  descartados: 'Descartados'
};

function reconciliarEstruturaPastasComprovantesDrive_(ambiente, pastaRaizId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return reconciliarEstruturaPastasComprovantesDriveSemLock_(ambiente, pastaRaizId);
  } finally {
    lock.releaseLock();
  }
}

function reconciliarEstruturaPastasComprovantesDriveSemLock_(ambiente, pastaRaizId) {
  const env = normalizarAmbienteComprovantesDrive_(ambiente);
  const props = PropertiesService.getScriptProperties();
  const configuracoes = obterConfiguracoesDominiosComprovantesDrive_();
  const anteriores = {};
  Object.keys(configuracoes).forEach(dominio => {
    const config = configuracoes[dominio];
    anteriores[dominio] = lerPastasComprovantesDriveSalvas_(
      config.folderProps,
      config.obterChave,
      env
    );
  });

  const raizComum = obterOuCriarRaizComprovantesDrive_(env, pastaRaizId);
  const resultado = {
    ambiente: env,
    raizComum,
    pastas: {},
    migracao: { arquivos_movidos: 0, pastas_movidas: 0, pastas_descartadas: 0 }
  };

  const raizesDominios = {};
  Object.keys(COMPROVANTES_DRIVE_DOMAIN_FOLDER_NAMES).forEach(dominio => {
    raizesDominios[dominio] = obterOuReaproveitarRaizDominioComprovantesDrive_(
      raizComum,
      dominio,
      anteriores[dominio]?.raiz,
      resultado.migracao
    );
  });

  Object.keys(COMPROVANTES_DRIVE_DOMAIN_FOLDER_NAMES).forEach(dominio => {
    const raizDominio = raizesDominios[dominio].raiz;
    const pastasDominio = { raiz: raizDominio };
    Object.keys(COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES).forEach(chave => {
      pastasDominio[chave] = reconciliarPastaStatusComprovantesDrive_(
        raizComum,
        raizDominio,
        dominio,
        chave,
        anteriores[dominio]?.[chave],
        raizesDominios[dominio].duplicadas,
        resultado.migracao
      );
    });
    resultado.pastas[dominio] = pastasDominio;
  });

  Object.keys(COMPROVANTES_DRIVE_DOMAIN_FOLDER_NAMES).forEach(dominio => {
    const info = raizesDominios[dominio];
    info.duplicadas.forEach(duplicada => {
      migrarConteudoPastaComprovantesDrive_(
        duplicada,
        info.raiz,
        dominio,
        'raiz',
        resultado.migracao
      );
      descartarPastaVaziaComprovantesDrive_(duplicada, resultado.migracao);
    });
  });

  validarEstruturaCompletaComprovantesDrive_(raizComum, resultado.pastas);
  const propriedades = {};
  propriedades[obterChaveRaizComprovantesDrive_(env)] = raizComum.getId();
  propriedades[obterChaveVersaoEstruturaComprovantesDrive_(env)] = COMPROVANTES_DRIVE_STRUCTURE_VERSION;
  Object.keys(configuracoes).forEach(dominio => {
    const config = configuracoes[dominio];
    Object.keys(config.folderProps).forEach(chave => {
      propriedades[config.obterChave(chave, env)] = resultado.pastas[dominio][chave].getId();
    });
  });
  props.setProperties(propriedades, false);
  return resultado;
}

function tentarReconciliarEstruturaComprovantesDriveNoAcesso_() {
  try {
    const acesso = typeof obterContextoUsuario === 'function'
      ? obterContextoUsuario(false)
      : { can_write: true };
    if (acesso?.can_write !== true) return null;
    const ambiente = obterAmbienteAtualComprovantesDrive_();
    const props = PropertiesService.getScriptProperties();
    const versao = String(props.getProperty(
      obterChaveVersaoEstruturaComprovantesDrive_(ambiente)
    ) || '').trim();
    if (versao === COMPROVANTES_DRIVE_STRUCTURE_VERSION) return null;
    return reconciliarEstruturaPastasComprovantesDrive_(ambiente);
  } catch (error) {
    console.log(JSON.stringify({
      evento: 'comprovantes.drive.reconciliacao_automatica_falhou',
      erro: String(error)
    }));
    return null;
  }
}

function obterConfiguracoesDominiosComprovantesDrive_() {
  return {
    despesas: {
      folderProps: INBOX_DESPESAS_DRIVE_FOLDER_PROPS,
      obterChave: obterChavePastaInboxDespesasDrive_
    },
    recebimentos: {
      folderProps: INBOX_RECEBIMENTOS_DRIVE_FOLDER_PROPS,
      obterChave: obterChavePastaInboxRecebimentosDrive_
    }
  };
}

function obterOuCriarRaizComprovantesDrive_(ambiente, pastaRaizId) {
  const env = normalizarAmbienteComprovantesDrive_(ambiente);
  const props = PropertiesService.getScriptProperties();
  const nomeEsperado = obterNomeRaizComprovantesDrive_(env);
  const candidatos = [
    String(pastaRaizId || '').trim(),
    String(props.getProperty(obterChaveRaizComprovantesDrive_(env)) || '').trim(),
    String(props.getProperty(obterChavePastaInboxDespesasDrive_('raiz', env)) || '').trim(),
    String(props.getProperty(obterChavePastaInboxRecebimentosDrive_('raiz', env)) || '').trim()
  ].filter(Boolean);

  for (const id of [...new Set(candidatos)]) {
    const raiz = resolverRaizComumComprovantesDrive_(
      tentarObterPastaComprovantesDrive_(id),
      nomeEsperado
    );
    if (raiz) return raiz;
  }
  const raizesPorNome = DriveApp.getFoldersByName(nomeEsperado);
  if (raizesPorNome.hasNext()) return raizesPorNome.next();
  return DriveApp.createFolder(nomeEsperado);
}

function obterOuReaproveitarRaizDominioComprovantesDrive_(
  raizComum,
  dominio,
  pastaAnterior,
  migracao
) {
  const nomeDominio = COMPROVANTES_DRIVE_DOMAIN_FOLDER_NAMES[dominio];
  if (!nomeDominio) throw new Error(`Dominio de comprovantes invalido: ${dominio}.`);
  const encontradas = listarSubpastasPorNomeComprovantesDrive_(raizComum, nomeDominio);
  if (pastaAnterior && pastaAnterior.getName() === nomeDominio) {
    if (!pastaTemPaiComprovantesDrive_(pastaAnterior, raizComum)) {
      pastaAnterior.moveTo(raizComum);
      migracao.pastas_movidas++;
    }
    return {
      raiz: pastaAnterior,
      duplicadas: encontradas.filter(pasta => pasta.getId() !== pastaAnterior.getId())
    };
  }
  if (encontradas.length) {
    return { raiz: encontradas[0], duplicadas: encontradas.slice(1) };
  }
  return { raiz: raizComum.createFolder(nomeDominio), duplicadas: [] };
}

function reconciliarPastaStatusComprovantesDrive_(
  raizComum,
  raizDominio,
  dominio,
  chave,
  pastaAnterior,
  raizesDuplicadas,
  migracao
) {
  const nome = COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES[chave];
  if (!nome) throw new Error(`Status de comprovantes invalido: ${chave}.`);
  const candidatas = listarSubpastasPorNomeComprovantesDrive_(raizDominio, nome);
  (raizesDuplicadas || []).forEach(raizDuplicada => {
    candidatas.push(...listarSubpastasPorNomeComprovantesDrive_(raizDuplicada, nome));
  });
  if (pastaAnterior && pastaAnterior.getName() === nome &&
      (pastaTemPaiComprovantesDrive_(pastaAnterior, raizDominio) ||
        (dominio === 'despesas' && pastaTemPaiComprovantesDrive_(pastaAnterior, raizComum)))) {
    candidatas.push(pastaAnterior);
  }

  // A estrutura anterior as separava por status diretamente na raiz comum.
  // Essas pastas pertencem a despesas; recebimentos sempre teve dominio proprio.
  if (dominio === 'despesas') {
    candidatas.push(...listarSubpastasPorNomeComprovantesDrive_(raizComum, nome));
  }

  const unicas = deduplicarPastasComprovantesDrive_(candidatas);
  let destino = selecionarPastaDestinoStatusComprovantesDrive_(
    unicas,
    pastaAnterior,
    raizDominio
  );
  if (!destino) destino = unicas[0] || raizDominio.createFolder(nome);
  if (!pastaTemPaiComprovantesDrive_(destino, raizDominio)) {
    destino.moveTo(raizDominio);
    migracao.pastas_movidas++;
  }

  unicas.forEach(origem => {
    if (origem.getId() === destino.getId()) return;
    migrarConteudoPastaComprovantesDrive_(origem, destino, dominio, chave, migracao);
    descartarPastaVaziaComprovantesDrive_(origem, migracao);
  });
  return destino;
}

function selecionarPastaDestinoStatusComprovantesDrive_(candidatas, pastaAnterior, raizDominio) {
  return (candidatas || []).find(pasta =>
    pastaAnterior && pasta.getId() === pastaAnterior.getId()
  ) || (candidatas || []).find(pasta =>
    pastaTemPaiComprovantesDrive_(pasta, raizDominio)
  ) || null;
}

function migrarConteudoPastaComprovantesDrive_(origem, destino, dominio, chave, migracao) {
  const arquivos = origem.getFiles();
  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    try {
      arquivo.moveTo(destino);
      migracao.arquivos_movidos++;
    } catch (error) {
      registrarFalhaMigracaoComprovantesDrive_(error, origem, destino, arquivo, dominio, chave);
      throw new Error(`Nao foi possivel mover um arquivo da pasta ${origem.getName()}.`);
    }
  }
  const subpastas = origem.getFolders();
  while (subpastas.hasNext()) {
    const subpasta = subpastas.next();
    try {
      subpasta.moveTo(destino);
      migracao.pastas_movidas++;
    } catch (error) {
      registrarFalhaMigracaoComprovantesDrive_(error, origem, destino, subpasta, dominio, chave);
      throw new Error(`Nao foi possivel mover uma subpasta de ${origem.getName()}.`);
    }
  }
}

function registrarFalhaMigracaoComprovantesDrive_(error, origem, destino, item, dominio, chave) {
  console.log(JSON.stringify({
    evento: 'comprovantes.drive.migracao_falhou',
    dominio: String(dominio || ''),
    status: String(chave || ''),
    pasta_origem_id: origem.getId(),
    pasta_destino_id: destino.getId(),
    item_id: String(item?.getId?.() || ''),
    erro: String(error)
  }));
}

function descartarPastaVaziaComprovantesDrive_(pasta, migracao) {
  if (pasta.getFiles().hasNext() || pasta.getFolders().hasNext()) {
    throw new Error(`A pasta antiga ${pasta.getName()} ainda contem itens e nao pode ser removida.`);
  }
  pasta.setTrashed(true);
  migracao.pastas_descartadas++;
}

function lerPastasComprovantesDriveSalvas_(folderProps, obterChave, ambiente) {
  const props = PropertiesService.getScriptProperties();
  const pastas = {};
  Object.keys(folderProps || {}).forEach(chave => {
    const id = String(props.getProperty(obterChave(chave, ambiente)) || '').trim();
    const pasta = tentarObterPastaComprovantesDrive_(id);
    if (pasta) pastas[chave] = pasta;
  });
  return pastas;
}

function estruturaDominioComprovantesDriveValida_(pastas, dominio, raizComum) {
  const nomeDominio = COMPROVANTES_DRIVE_DOMAIN_FOLDER_NAMES[dominio];
  if (!nomeDominio || !pastas?.raiz || pastas.raiz.getName() !== nomeDominio) return false;
  if (raizComum && !pastaTemPaiComprovantesDrive_(pastas.raiz, raizComum)) return false;
  return Object.keys(COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES).every(chave =>
    pastas[chave] &&
    pastas[chave].getName() === COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES[chave] &&
    pastaTemPaiComprovantesDrive_(pastas[chave], pastas.raiz)
  );
}

function validarEstruturaCompletaComprovantesDrive_(raizComum, dominios) {
  Object.keys(COMPROVANTES_DRIVE_DOMAIN_FOLDER_NAMES).forEach(dominio => {
    if (!estruturaDominioComprovantesDriveValida_(dominios[dominio], dominio, raizComum)) {
      throw new Error(`Estrutura de comprovantes invalida para ${dominio}.`);
    }
    if (listarSubpastasPorNomeComprovantesDrive_(
      raizComum,
      COMPROVANTES_DRIVE_DOMAIN_FOLDER_NAMES[dominio]
    ).length !== 1) {
      throw new Error(`Existe mais de uma pasta de dominio para ${dominio}.`);
    }
    Object.keys(COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES).forEach(chave => {
      if (listarSubpastasPorNomeComprovantesDrive_(
        dominios[dominio].raiz,
        COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES[chave]
      ).length !== 1) {
        throw new Error(`Estrutura duplicada em ${dominio}/${chave}.`);
      }
    });
  });
  Object.values(COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES).forEach(nome => {
    if (listarSubpastasPorNomeComprovantesDrive_(raizComum, nome).length) {
      throw new Error(`A pasta antiga ${nome} ainda esta na raiz de comprovantes.`);
    }
  });
}

function resolverRaizComumComprovantesDrive_(pasta, nomeEsperado) {
  if (!pasta) return null;
  const fila = [pasta];
  const visitadas = {};
  while (fila.length) {
    const atual = fila.shift();
    const id = atual.getId();
    if (visitadas[id]) continue;
    visitadas[id] = true;
    if (atual.getName() === nomeEsperado) return atual;
    const pais = atual.getParents();
    while (pais.hasNext()) fila.push(pais.next());
  }
  return null;
}

function tentarObterPastaComprovantesDrive_(id) {
  if (!id) return null;
  try {
    return DriveApp.getFolderById(id);
  } catch (error) {
    return null;
  }
}

function listarSubpastasPorNomeComprovantesDrive_(raiz, nome) {
  const resultado = [];
  const pastas = raiz.getFoldersByName(nome);
  while (pastas.hasNext()) resultado.push(pastas.next());
  return resultado;
}

function deduplicarPastasComprovantesDrive_(pastas) {
  const ids = {};
  return (pastas || []).filter(pasta => {
    if (!pasta) return false;
    const id = pasta.getId();
    if (ids[id]) return false;
    ids[id] = true;
    return true;
  });
}

function pastaTemPaiComprovantesDrive_(pasta, paiEsperado) {
  if (!pasta || !paiEsperado) return false;
  const pais = pasta.getParents();
  while (pais.hasNext()) {
    if (pais.next().getId() === paiEsperado.getId()) return true;
  }
  return false;
}

function obterNomeRaizComprovantesDrive_(ambiente) {
  return normalizarAmbienteComprovantesDrive_(ambiente) === 'dev'
    ? `${COMPROVANTES_DRIVE_ROOT_FOLDER_NAME} - DEV`
    : COMPROVANTES_DRIVE_ROOT_FOLDER_NAME;
}

function obterChaveRaizComprovantesDrive_(ambiente) {
  return normalizarAmbienteComprovantesDrive_(ambiente) === 'dev'
    ? `${COMPROVANTES_DRIVE_COMMON_ROOT_PROP}_DEV`
    : COMPROVANTES_DRIVE_COMMON_ROOT_PROP;
}

function obterChaveVersaoEstruturaComprovantesDrive_(ambiente) {
  return normalizarAmbienteComprovantesDrive_(ambiente) === 'dev'
    ? `${COMPROVANTES_DRIVE_STRUCTURE_VERSION_PROP}_DEV`
    : COMPROVANTES_DRIVE_STRUCTURE_VERSION_PROP;
}

function obterAmbienteAtualComprovantesDrive_() {
  if (typeof getDbEnvironmentExecutionEffective_ === 'function') {
    return normalizarAmbienteComprovantesDrive_(getDbEnvironmentExecutionEffective_());
  }
  if (typeof getUserDbEnvironment_ === 'function') {
    return normalizarAmbienteComprovantesDrive_(getUserDbEnvironment_());
  }
  return 'prod';
}

function normalizarAmbienteComprovantesDrive_(ambiente) {
  return String(ambiente || '').trim().toLowerCase() === 'dev' ? 'dev' : 'prod';
}

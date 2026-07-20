const COMPROVANTES_DRIVE_ROOT_FOLDER_NAME = 'Comprovantes CarpintariaZizu';
const COMPROVANTES_DRIVE_COMMON_ROOT_PROP = 'COMPROVANTES_DRIVE_ROOT_FOLDER_ID';
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

function obterOuCriarRaizComprovantesDrive_(ambiente, pastaRaizId) {
  const env = normalizarAmbienteComprovantesDrive_(ambiente);
  const props = PropertiesService.getScriptProperties();
  const nomeEsperado = obterNomeRaizComprovantesDrive_(env);
  const candidatos = [
    String(pastaRaizId || '').trim(),
    String(props.getProperty(obterChaveRaizComprovantesDrive_(env)) || '').trim(),
    String(props.getProperty(obterChavePastaInboxDespesasDrive_('raiz', env)) || '').trim()
  ].filter(Boolean);

  let raiz = null;
  for (const id of candidatos) {
    const pasta = tentarObterPastaComprovantesDrive_(id);
    raiz = resolverRaizComumComprovantesDrive_(pasta, nomeEsperado);
    if (raiz) break;
  }
  if (!raiz) raiz = DriveApp.createFolder(nomeEsperado);
  props.setProperty(obterChaveRaizComprovantesDrive_(env), raiz.getId());
  return raiz;
}

function criarEstruturaDominioComprovantesDrive_(raizComum, dominio, pastasAnteriores) {
  const nomeDominio = COMPROVANTES_DRIVE_DOMAIN_FOLDER_NAMES[dominio];
  if (!nomeDominio) throw new Error(`Dominio de comprovantes invalido: ${dominio}.`);
  const raizDominio = obterOuCriarSubpastaComprovantesDrive_(raizComum, nomeDominio);
  const pastas = { raiz: raizDominio };
  Object.keys(COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES).forEach(chave => {
    pastas[chave] = obterOuReaproveitarPastaStatusComprovantesDrive_(
      raizDominio,
      COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES[chave],
      pastasAnteriores?.[chave],
      dominio
    );
  });
  return pastas;
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

function migrarArquivosPastasComprovantesDrive_(pastasOrigem, pastasDestino, contexto) {
  const resultado = { movidos: 0, falhas: 0 };
  Object.keys(COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES).forEach(chave => {
    const origem = pastasOrigem?.[chave];
    const destino = pastasDestino?.[chave];
    if (!origem || !destino || origem.getId() === destino.getId()) return;
    const arquivos = origem.getFiles();
    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      try {
        arquivo.moveTo(destino);
        resultado.movidos++;
      } catch (error) {
        resultado.falhas++;
        console.log(JSON.stringify({
          evento: 'comprovantes.drive.migracao_arquivo_falhou',
          contexto: String(contexto || ''),
          pasta_origem_id: origem.getId(),
          pasta_destino_id: destino.getId(),
          arquivo_id: arquivo.getId(),
          erro: String(error)
        }));
      }
    }
  });
  return resultado;
}

function estruturaDominioComprovantesDriveValida_(pastas, dominio) {
  const nomeDominio = COMPROVANTES_DRIVE_DOMAIN_FOLDER_NAMES[dominio];
  if (!nomeDominio || !pastas?.raiz || pastas.raiz.getName() !== nomeDominio) return false;
  return Object.keys(COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES).every(chave =>
    pastas[chave] && pastas[chave].getName() === COMPROVANTES_DRIVE_STATUS_FOLDER_NAMES[chave]
  );
}

function resolverRaizComumComprovantesDrive_(pasta, nomeEsperado) {
  if (!pasta) return null;
  if (pasta.getName() === nomeEsperado) return pasta;
  const nomesDominio = Object.values(COMPROVANTES_DRIVE_DOMAIN_FOLDER_NAMES);
  if (!nomesDominio.includes(pasta.getName())) return null;
  const pais = pasta.getParents();
  while (pais.hasNext()) {
    const pai = pais.next();
    if (pai.getName() === nomeEsperado) return pai;
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

function obterOuCriarSubpastaComprovantesDrive_(raiz, nome) {
  const encontradas = raiz.getFoldersByName(nome);
  return encontradas.hasNext() ? encontradas.next() : raiz.createFolder(nome);
}

function obterOuReaproveitarPastaStatusComprovantesDrive_(raizDominio, nome, pastaAnterior, dominio) {
  const encontradas = raizDominio.getFoldersByName(nome);
  if (encontradas.hasNext()) return encontradas.next();
  if (pastaAnterior && pastaAnterior.getName() === nome) {
    try {
      pastaAnterior.moveTo(raizDominio);
      return pastaAnterior;
    } catch (error) {
      console.log(JSON.stringify({
        evento: 'comprovantes.drive.migracao_pasta_falhou',
        dominio: String(dominio || ''),
        pasta_id: pastaAnterior.getId(),
        pasta_nome: nome,
        destino_id: raizDominio.getId(),
        erro: String(error)
      }));
    }
  }
  return raizDominio.createFolder(nome);
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

function normalizarAmbienteComprovantesDrive_(ambiente) {
  return String(ambiente || '').trim().toLowerCase() === 'dev' ? 'dev' : 'prod';
}

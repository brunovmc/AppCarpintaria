const DB_ENV_PROD = 'prod';
const DB_ENV_DEV = 'dev';
const DB_ENV_USER_PROP_KEY = 'APP_DB_ENV_ACTIVE';

const DB_SYNC_ARMED_ACTION_KEY = 'DB_SYNC_ARMED_ACTION';
const DB_SYNC_ARMED_UNTIL_MS_KEY = 'DB_SYNC_ARMED_UNTIL_MS';
const DB_SYNC_ARMED_BY_KEY = 'DB_SYNC_ARMED_BY';
const DB_SYNC_ARMED_AT_ISO_KEY = 'DB_SYNC_ARMED_AT_ISO';

const DB_SYNC_WRITE_LOCK_UNTIL_MS_KEY = 'DB_SYNC_WRITE_LOCK_UNTIL_MS';
const DB_SYNC_WRITE_LOCK_ACTION_KEY = 'DB_SYNC_WRITE_LOCK_ACTION';
const DB_SYNC_WRITE_LOCK_BY_KEY = 'DB_SYNC_WRITE_LOCK_BY';

const DB_SYNC_ACTION_DEV_TO_PROD = 'DEV_TO_PROD';
const DB_SYNC_ACTION_PROD_TO_DEV = 'PROD_TO_DEV';
const DB_SYNC_ARM_TTL_MS = 10 * 60 * 1000; // 10 minutos
const DB_SYNC_WRITE_LOCK_TTL_MS = 20 * 60 * 1000; // 20 minutos

function normalizarAmbienteBancoDados_(valor) {
  const env = String(valor || '').trim().toLowerCase();
  if (env === DB_ENV_DEV) return DB_ENV_DEV;
  return DB_ENV_PROD;
}

function getSpreadsheetIdsConfigurados_() {
  const prod = String(
    (typeof DATA_SPREADSHEET_ID_PROD === 'string' ? DATA_SPREADSHEET_ID_PROD : '')
      || (typeof DATA_SPREADSHEET_ID === 'string' ? DATA_SPREADSHEET_ID : '')
      || ''
  ).trim();
  const dev = String(typeof DATA_SPREADSHEET_ID_DEV === 'string' ? DATA_SPREADSHEET_ID_DEV : '').trim();

  if (!prod) {
    throw new Error('DATA_SPREADSHEET_ID_PROD nao configurado.');
  }

  return { prod, dev };
}

function getSpreadsheetIdPorAmbiente_(ambiente) {
  const ids = getSpreadsheetIdsConfigurados_();
  const env = normalizarAmbienteBancoDados_(ambiente);
  if (env === DB_ENV_DEV) {
    if (!ids.dev) {
      throw new Error('DATA_SPREADSHEET_ID_DEV nao configurado.');
    }
    return ids.dev;
  }
  return ids.prod;
}

function getUserDbEnvProperties_() {
  return PropertiesService.getUserProperties();
}

function getUserDbEnvironment_() {
  try {
    const raw = getUserDbEnvProperties_().getProperty(DB_ENV_USER_PROP_KEY);
    return normalizarAmbienteBancoDados_(raw);
  } catch (error) {
    return DB_ENV_PROD;
  }
}

function setUserDbEnvironment_(ambiente) {
  const env = normalizarAmbienteBancoDados_(ambiente);
  getUserDbEnvProperties_().setProperty(DB_ENV_USER_PROP_KEY, env);
  return env;
}

function getDataSpreadsheetIdAtivo(opcoes) {
  const opts = opcoes || {};
  const envExplicito = String(opts.targetEnv || opts.env || '').trim().toLowerCase();
  if (envExplicito) {
    return getSpreadsheetIdPorAmbiente_(envExplicito);
  }
  return getSpreadsheetIdPorAmbiente_(getUserDbEnvironment_());
}

function getDbEnvironmentContext_() {
  const acesso = (typeof obterContextoUsuario === 'function')
    ? obterContextoUsuario(false)
    : { role: 'admin', can_write: true, can_read: true, email: '' };
  const isAdmin = String(acesso?.role || '').toLowerCase() === 'admin' && acesso?.can_write === true;
  const envUsuario = getUserDbEnvironment_();
  const ambienteAtivo = isAdmin ? envUsuario : DB_ENV_PROD;
  const ids = getSpreadsheetIdsConfigurados_();

  return {
    can_read: acesso?.can_read !== false,
    can_write: acesso?.can_write === true,
    role: String(acesso?.role || '').trim() || (isAdmin ? 'admin' : 'viewer'),
    email: String(acesso?.email || '').trim(),
    can_toggle: isAdmin,
    selected_env: ambienteAtivo,
    effective_env: ambienteAtivo,
    labels: {
      prod: 'PROD',
      dev: 'DEV'
    },
    ids_configured: {
      prod: !!ids.prod,
      dev: !!ids.dev
    }
  };
}

function obterContextoBancoDados() {
  return getDbEnvironmentContext_();
}

function definirContextoBancoDados(ambiente) {
  const ctx = getDbEnvironmentContext_();
  if (!ctx.can_toggle) {
    setUserDbEnvironment_(DB_ENV_PROD);
    return getDbEnvironmentContext_();
  }

  const env = normalizarAmbienteBancoDados_(ambiente);
  if (env === DB_ENV_DEV) {
    const ids = getSpreadsheetIdsConfigurados_();
    if (!ids.dev) {
      throw new Error('Ambiente DEV indisponivel: DATA_SPREADSHEET_ID_DEV nao configurado.');
    }
  }

  setUserDbEnvironment_(env);
  return getDbEnvironmentContext_();
}

function resetarContextoBancoDadosSessao() {
  setUserDbEnvironment_(DB_ENV_PROD);
  return getDbEnvironmentContext_();
}

function limparChavesSyncArmadas_() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(DB_SYNC_ARMED_ACTION_KEY);
  props.deleteProperty(DB_SYNC_ARMED_UNTIL_MS_KEY);
  props.deleteProperty(DB_SYNC_ARMED_BY_KEY);
  props.deleteProperty(DB_SYNC_ARMED_AT_ISO_KEY);
}

function getArmedSyncStatus_() {
  const props = PropertiesService.getScriptProperties();
  const action = String(props.getProperty(DB_SYNC_ARMED_ACTION_KEY) || '').trim();
  const untilMs = Number(props.getProperty(DB_SYNC_ARMED_UNTIL_MS_KEY) || 0);
  const armedBy = String(props.getProperty(DB_SYNC_ARMED_BY_KEY) || '').trim();
  const armedAtIso = String(props.getProperty(DB_SYNC_ARMED_AT_ISO_KEY) || '').trim();
  const now = Date.now();

  if (!action || !Number.isFinite(untilMs) || untilMs <= now) {
    if (action || untilMs) {
      limparChavesSyncArmadas_();
    }
    return {
      armed: false,
      action: '',
      untilMs: 0,
      armedBy: '',
      armedAtIso: ''
    };
  }

  return {
    armed: true,
    action,
    untilMs,
    armedBy,
    armedAtIso
  };
}

function descreverAcaoSync_(action) {
  if (action === DB_SYNC_ACTION_DEV_TO_PROD) return 'DEV -> PROD';
  if (action === DB_SYNC_ACTION_PROD_TO_DEV) return 'PROD -> DEV';
  return String(action || '').trim() || 'N/A';
}

function assertAdminParaSync_(acao) {
  if (typeof obterContextoUsuario !== 'function') return true;
  const ctx = obterContextoUsuario(false);
  const isAdmin = String(ctx?.role || '').toLowerCase() === 'admin' && ctx?.can_write === true;
  if (isAdmin) return true;
  throw new Error(`${String(acao || 'Operacao').trim() || 'Operacao'} bloqueada. Apenas admin pode executar sync de banco.`);
}

function armarSyncBancoDados_(action) {
  assertAdminParaSync_('Armamento de sync de banco');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const now = Date.now();
    const untilMs = now + DB_SYNC_ARM_TTL_MS;
    const email = (typeof obterEmailUsuarioAtualAcesso === 'function')
      ? String(obterEmailUsuarioAtualAcesso() || '').trim()
      : '';
    const props = PropertiesService.getScriptProperties();
    props.setProperty(DB_SYNC_ARMED_ACTION_KEY, action);
    props.setProperty(DB_SYNC_ARMED_UNTIL_MS_KEY, String(untilMs));
    props.setProperty(DB_SYNC_ARMED_BY_KEY, email || 'admin');
    props.setProperty(DB_SYNC_ARMED_AT_ISO_KEY, new Date(now).toISOString());
    return {
      ok: true,
      armed: true,
      action,
      action_label: descreverAcaoSync_(action),
      armed_by: email || 'admin',
      armed_at: new Date(now),
      expires_at: new Date(untilMs)
    };
  } finally {
    lock.releaseLock();
  }
}

function armSyncDevToProdDB() {
  return armarSyncBancoDados_(DB_SYNC_ACTION_DEV_TO_PROD);
}

function armRevertProdToDevDB() {
  return armarSyncBancoDados_(DB_SYNC_ACTION_PROD_TO_DEV);
}

function limparBloqueioEscritaSyncDB_() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(DB_SYNC_WRITE_LOCK_UNTIL_MS_KEY);
  props.deleteProperty(DB_SYNC_WRITE_LOCK_ACTION_KEY);
  props.deleteProperty(DB_SYNC_WRITE_LOCK_BY_KEY);
}

function setBloqueioEscritaSyncDB_(action) {
  const now = Date.now();
  const untilMs = now + DB_SYNC_WRITE_LOCK_TTL_MS;
  const email = (typeof obterEmailUsuarioAtualAcesso === 'function')
    ? String(obterEmailUsuarioAtualAcesso() || '').trim()
    : '';
  const props = PropertiesService.getScriptProperties();
  props.setProperty(DB_SYNC_WRITE_LOCK_UNTIL_MS_KEY, String(untilMs));
  props.setProperty(DB_SYNC_WRITE_LOCK_ACTION_KEY, String(action || '').trim());
  props.setProperty(DB_SYNC_WRITE_LOCK_BY_KEY, email || 'admin');
  return {
    locked: true,
    action: action,
    by: email || 'admin',
    untilMs
  };
}

function obterStatusBloqueioSyncDB() {
  const props = PropertiesService.getScriptProperties();
  const untilMs = Number(props.getProperty(DB_SYNC_WRITE_LOCK_UNTIL_MS_KEY) || 0);
  const action = String(props.getProperty(DB_SYNC_WRITE_LOCK_ACTION_KEY) || '').trim();
  const by = String(props.getProperty(DB_SYNC_WRITE_LOCK_BY_KEY) || '').trim();
  const now = Date.now();

  if (!Number.isFinite(untilMs) || untilMs <= now) {
    if (untilMs || action || by) {
      limparBloqueioEscritaSyncDB_();
    }
    return {
      blocked: false,
      action: '',
      by: '',
      untilMs: 0
    };
  }

  return {
    blocked: true,
    action,
    by,
    untilMs
  };
}

function assertPodeEscreverNoBancoDados(acao) {
  const status = obterStatusBloqueioSyncDB();
  if (!status.blocked) return true;

  const contexto = String(acao || 'Operacao de escrita').trim() || 'Operacao de escrita';
  const ate = Utilities.formatDate(
    new Date(status.untilMs),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
  const descricaoAcao = descreverAcaoSync_(status.action);
  throw new Error(`${contexto} bloqueada. Sincronizacao manual (${descricaoAcao}) em andamento ate ${ate}.`);
}

function ajustarDimensaoAbaParaSync_(sheet, rowsNeed, colsNeed) {
  const minRows = Math.max(1, Number(rowsNeed) || 1);
  const minCols = Math.max(1, Number(colsNeed) || 1);

  const currentRows = sheet.getMaxRows();
  const currentCols = sheet.getMaxColumns();

  if (currentRows < minRows) {
    sheet.insertRowsAfter(currentRows, minRows - currentRows);
  }
  if (currentCols < minCols) {
    sheet.insertColumnsAfter(currentCols, minCols - currentCols);
  }
}

function copiarDadosPlanilhaCompleta_(sourceSpreadsheetId, targetSpreadsheetId) {
  const source = SpreadsheetApp.openById(sourceSpreadsheetId);
  const target = SpreadsheetApp.openById(targetSpreadsheetId);

  const sourceSheets = source.getSheets();
  if (!Array.isArray(sourceSheets) || sourceSheets.length === 0) {
    throw new Error('Planilha de origem sem abas para sincronizar.');
  }

  const sourceNameSet = {};
  sourceSheets.forEach(s => {
    sourceNameSet[s.getName()] = true;
  });

  const targetByName = {};
  target.getSheets().forEach(s => {
    targetByName[s.getName()] = s;
  });

  const resumoAbas = [];

  sourceSheets.forEach((sourceSheet, idx) => {
    const name = sourceSheet.getName();
    let targetSheet = targetByName[name];
    if (!targetSheet) {
      targetSheet = target.insertSheet(name);
      targetByName[name] = targetSheet;
    }

    ajustarDimensaoAbaParaSync_(
      targetSheet,
      sourceSheet.getMaxRows(),
      sourceSheet.getMaxColumns()
    );

    targetSheet.clear();

    const sourceLastRow = sourceSheet.getLastRow();
    const sourceLastCol = sourceSheet.getLastColumn();
    if (sourceLastRow > 0 && sourceLastCol > 0) {
      const sourceRange = sourceSheet.getRange(1, 1, sourceLastRow, sourceLastCol);
      const targetRange = targetSheet.getRange(1, 1, sourceLastRow, sourceLastCol);
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    }

    targetSheet.setFrozenRows(sourceSheet.getFrozenRows());
    targetSheet.setFrozenColumns(sourceSheet.getFrozenColumns());
    try {
      targetSheet.setTabColor(sourceSheet.getTabColor());
    } catch (error) {
      // sem acao
    }

    target.setActiveSheet(targetSheet);
    target.moveActiveSheet(idx + 1);

    resumoAbas.push({
      aba: name,
      linhas: sourceLastRow,
      colunas: sourceLastCol
    });
  });

  const sheetsToDelete = target
    .getSheets()
    .filter(s => !sourceNameSet[s.getName()]);

  sheetsToDelete.forEach(sheet => {
    if (target.getSheets().length <= 1) return;
    target.deleteSheet(sheet);
  });

  SpreadsheetApp.flush();

  return {
    total_abas_copiadas: resumoAbas.length,
    abas: resumoAbas,
    abas_removidas: sheetsToDelete.map(s => s.getName())
  };
}

function limparCachesAposSyncBancoDados_() {
  const envOriginal = getUserDbEnvironment_();
  const ids = getSpreadsheetIdsConfigurados_();
  const ambientes = [DB_ENV_PROD];
  if (ids.dev) ambientes.push(DB_ENV_DEV);

  try {
    ambientes.forEach(env => {
      setUserDbEnvironment_(env);
      if (typeof limparCacheValidacoes === 'function') limparCacheValidacoes();
      if (typeof limparCacheEstoque === 'function') limparCacheEstoque();
      if (typeof limparCacheCompras === 'function') limparCacheCompras();
      if (typeof limparCacheVendas === 'function') limparCacheVendas();
      if (typeof limparCacheProdutos === 'function') limparCacheProdutos();
      if (typeof limparCacheProducao === 'function') limparCacheProducao();
      if (typeof limparCacheDespesasGerais === 'function') limparCacheDespesasGerais();
      if (typeof limparCachePagamentos === 'function') limparCachePagamentos();
      if (typeof limparCacheUsuariosAcesso === 'function') limparCacheUsuariosAcesso();
      if (typeof limparCacheDashboardFinanceiro === 'function') limparCacheDashboardFinanceiro();
    });
  } finally {
    setUserDbEnvironment_(envOriginal);
  }
}

function executarSyncBancoDadosArmado_(expectedAction) {
  assertAdminParaSync_('Sync de banco');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const armed = getArmedSyncStatus_();
    if (!armed.armed) {
      throw new Error(`Sync nao autorizado. Execute arm primeiro para ${descreverAcaoSync_(expectedAction)}.`);
    }
    if (armed.action !== expectedAction) {
      throw new Error(`Sync armado para ${descreverAcaoSync_(armed.action)}, mas solicitado ${descreverAcaoSync_(expectedAction)}.`);
    }

    setBloqueioEscritaSyncDB_(expectedAction);

    try {
      const sourceEnv = expectedAction === DB_SYNC_ACTION_DEV_TO_PROD ? DB_ENV_DEV : DB_ENV_PROD;
      const targetEnv = expectedAction === DB_SYNC_ACTION_DEV_TO_PROD ? DB_ENV_PROD : DB_ENV_DEV;
      const sourceId = getSpreadsheetIdPorAmbiente_(sourceEnv);
      const targetId = getSpreadsheetIdPorAmbiente_(targetEnv);
      const startedAt = new Date();

      const copyResult = copiarDadosPlanilhaCompleta_(sourceId, targetId);
      limparCachesAposSyncBancoDados_();

      return {
        ok: true,
        action: expectedAction,
        action_label: descreverAcaoSync_(expectedAction),
        source_env: sourceEnv,
        target_env: targetEnv,
        source_id: sourceId,
        target_id: targetId,
        started_at: startedAt,
        finished_at: new Date(),
        resumo: copyResult
      };
    } finally {
      limparBloqueioEscritaSyncDB_();
      limparChavesSyncArmadas_();
    }
  } finally {
    lock.releaseLock();
  }
}

function syncDevToProdDB() {
  return executarSyncBancoDadosArmado_(DB_SYNC_ACTION_DEV_TO_PROD);
}

function revertProdToDevDB() {
  return executarSyncBancoDadosArmado_(DB_SYNC_ACTION_PROD_TO_DEV);
}

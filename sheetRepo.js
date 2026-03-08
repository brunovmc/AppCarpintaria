function getDataSpreadsheet() {
  const id = (typeof DATA_SPREADSHEET_ID === 'string')
    ? DATA_SPREADSHEET_ID.trim()
    : '';

  if (!id) {
    throw new Error('DATA_SPREADSHEET_ID nao configurado em main.js');
  }

  try {
    return SpreadsheetApp.openById(id);
  } catch (error) {
    throw new Error('Nao foi possivel abrir a planilha de dados. Verifique o DATA_SPREADSHEET_ID e as permissoes de acesso. Detalhes: ' + error.message);
  }
}

const APP_CACHE_VERSION = 'v1';

function getAppCacheKey(scope) {
  const id = (typeof DATA_SPREADSHEET_ID === 'string')
    ? DATA_SPREADSHEET_ID.trim()
    : '';
  const escopo = String(scope || '').trim().toUpperCase();
  return `APP_CACHE:${id || 'SEM_ID'}:${APP_CACHE_VERSION}:${escopo}`;
}

function appCacheGetJson(scope) {
  const key = getAppCacheKey(scope);
  try {
    const raw = CacheService.getScriptCache().get(key);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
}

function appCachePutJson(scope, payload, ttlSec) {
  const key = getAppCacheKey(scope);
  const ttl = Number(ttlSec) > 0 ? Number(ttlSec) : 60;
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(payload ?? null), ttl);
    return { ok: true, key, ttl };
  } catch (error) {
    return { ok: false, key, ttl, erro: error.message };
  }
}

function appCacheRemove(scope) {
  const key = getAppCacheKey(scope);
  try {
    CacheService.getScriptCache().remove(key);
    return { ok: true, key };
  } catch (error) {
    return { ok: false, key, erro: error.message };
  }
}

function invalidarCachesRelacionadosAba(sheetName) {
  const aba = String(sheetName || '').trim().toUpperCase();
  if (!aba) return;

  if (aba === 'ESTOQUE' && typeof limparCacheEstoque === 'function') {
    limparCacheEstoque();
  }

  if (aba === 'COMPRAS' && typeof limparCacheCompras === 'function') {
    limparCacheCompras();
  }

  if (aba === 'DESPESAS_GERAIS' && typeof limparCacheDespesasGerais === 'function') {
    limparCacheDespesasGerais();
  }

  if (aba === 'PAGAMENTOS') {
    if (typeof limparCachePagamentos === 'function') {
      limparCachePagamentos();
    }
    if (typeof limparCacheCompras === 'function') {
      limparCacheCompras();
    }
    if (typeof limparCacheDespesasGerais === 'function') {
      limparCacheDespesasGerais();
    }
  }

  if ((aba === 'VALIDACAO' || aba === 'VALIDACAO_TIPO_CATEGORIA') && typeof limparCacheValidacoes === 'function') {
    limparCacheValidacoes();
  }

  if (
    (aba === 'ESTOQUE' || aba === 'COMPRAS' || aba === 'DESPESAS_GERAIS' || aba === 'PAGAMENTOS') &&
    typeof limparCacheDashboardFinanceiro === 'function'
  ) {
    limparCacheDashboardFinanceiro();
  }
}

function getSheet(nome) {
  return getDataSpreadsheet().getSheetByName(nome);
}

function getHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => map[h] = i);
  return map;
}

function ensureSchema(sheet, schema) {
  if (!sheet || !Array.isArray(schema)) return;

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    sheet.getRange(1, 1, 1, schema.length).setValues([schema]);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const missing = schema.filter(h => !headers.includes(h));
  if (missing.length === 0) return;

  const startCol = headers.length + 1;
  sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
}

function rowsToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  return data.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function insert(sheetName, payload, schema) {
  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  ensureSchema(sheet, schema);
  const headerMap = getHeaderMap(sheet);
  const row = Array(sheet.getLastColumn()).fill('');

  schema.forEach(key => {
    if (key in headerMap) {
      row[headerMap[key]] = payload[key] ?? '';
    }
  });

  sheet.appendRow(row);
  invalidarCachesRelacionadosAba(sheetName);
  return true;
}


function updateById(sheetName, idField, id, payload, schema) {
  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  ensureSchema(sheet, schema);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf(idField);

  if (idCol === -1) return false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) {
      const headerMap = getHeaderMap(sheet);
      const row = [...data[i]];
      let alterou = false;

      Object.keys(payload || {}).forEach(key => {
        if (key in headerMap) {
          row[headerMap[key]] = payload[key];
          alterou = true;
        }
      });

      if (alterou) {
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      }

      invalidarCachesRelacionadosAba(sheetName);
      return true;
    }
  }
  return false;
}

const EXECUTION_MEMORY_CACHE_KEY = '__APP_EXECUTION_MEMORY_CACHE__';

function getExecutionMemoryCache() {
  const globalScope = (typeof globalThis !== 'undefined' && globalThis)
    ? globalThis
    : (typeof this !== 'undefined' ? this : {});

  if (!globalScope[EXECUTION_MEMORY_CACHE_KEY]) {
    globalScope[EXECUTION_MEMORY_CACHE_KEY] = {};
  }

  return globalScope[EXECUTION_MEMORY_CACHE_KEY];
}

function getDataSpreadsheet(opcoes) {
  const opts = opcoes || {};
  if (!opts.skipAccessCheck && typeof assertCanRead === 'function') {
    assertCanRead('Acesso aos dados');
  }

  const id = (typeof getDataSpreadsheetIdAtivo === 'function')
    ? String(getDataSpreadsheetIdAtivo(opts) || '').trim()
    : String((typeof DATA_SPREADSHEET_ID === 'string' ? DATA_SPREADSHEET_ID : '') || '').trim();

  if (!id) {
    throw new Error('ID da planilha de dados nao configurado.');
  }

  const executionCache = getExecutionMemoryCache();
  if (executionCache[id]) {
    return executionCache[id];
  }

  try {
    const ss = SpreadsheetApp.openById(id);
    executionCache[id] = ss;
    return ss;
  } catch (error) {
    throw new Error('Nao foi possivel abrir a planilha de dados. Verifique os IDs de PROD/DEV e as permissoes de acesso. Detalhes: ' + error.message);
  }
}

const APP_CACHE_VERSION = 'v1';
const ROW_LOOKUP_CACHE_VERSION = 'v1';
const ROW_LOOKUP_CACHE_TTL_SEC = 21600; // 6h

function getAppCacheKey(scope) {
  const id = (typeof getDataSpreadsheetIdAtivo === 'function')
    ? String(getDataSpreadsheetIdAtivo({ skipAccessCheck: true }) || '').trim()
    : String((typeof DATA_SPREADSHEET_ID === 'string' ? DATA_SPREADSHEET_ID : '') || '').trim();
  const escopo = String(scope || '').trim().toUpperCase();
  return `APP_CACHE:${id || 'SEM_ID'}:${APP_CACHE_VERSION}:${escopo}`;
}

function appCacheGetJson(scope) {
  const escopo = String(scope || '').trim().toUpperCase();
  const scopeLiberadoAcesso = (typeof USUARIOS_ACESSO_CACHE_SCOPE === 'string')
    ? String(USUARIOS_ACESSO_CACHE_SCOPE || '').trim().toUpperCase()
    : 'USUARIOS_ACESSO_MAPA';
  if (
    typeof assertCanRead === 'function' &&
    escopo &&
    escopo !== scopeLiberadoAcesso
  ) {
    assertCanRead(`Leitura de cache ${escopo}`);
  }

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
  const afetaDashboardFinanceiro =
    aba === 'ESTOQUE' ||
    aba === 'COMPRAS' ||
    aba === 'VENDAS' ||
    aba === 'DESPESAS_GERAIS' ||
    aba === 'PAGAMENTOS' ||
    aba === 'PARCELAS_FINANCEIRAS' ||
    aba.startsWith('PRODUCAO') ||
    aba.startsWith('PRODUTOS');

  if (aba === 'ESTOQUE' && typeof limparCacheEstoque === 'function') {
    limparCacheEstoque();
  }

  if (aba === 'COMPRAS' && typeof limparCacheCompras === 'function') {
    limparCacheCompras();
  }

  if (aba === 'VENDAS' && typeof limparCacheVendas === 'function') {
    limparCacheVendas();
  }

  if (aba === 'PRODUTOS' && typeof limparCacheProdutos === 'function') {
    limparCacheProdutos();
  }

  if (aba === 'PRODUCAO' && typeof limparCacheProducao === 'function') {
    limparCacheProducao();
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
    if (typeof limparCacheVendas === 'function') {
      limparCacheVendas();
    }
  }

  if (aba === 'PARCELAS_FINANCEIRAS' && typeof limparCacheDashboardFinanceiro === 'function') {
    limparCacheDashboardFinanceiro();
  }

  if ((aba === 'VALIDACAO' || aba === 'VALIDACAO_TIPO_CATEGORIA') && typeof limparCacheValidacoes === 'function') {
    limparCacheValidacoes();
  }

  if (aba === 'USUARIOS' && typeof limparCacheUsuariosAcesso === 'function') {
    limparCacheUsuariosAcesso();
  }

  if (aba.startsWith('PRODUTOS') && typeof limparCacheProdutos === 'function') {
    limparCacheProdutos();
  }

  if (aba.startsWith('PRODUCAO') && typeof limparCacheProducao === 'function') {
    limparCacheProducao();
  }

  if (afetaDashboardFinanceiro && typeof limparCacheDashboardFinanceiro === 'function') {
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

function buildHeaderMapFromHeaders(headers) {
  const map = {};
  (Array.isArray(headers) ? headers : []).forEach((h, i) => {
    map[h] = i;
  });
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

function criarContextoLeituraRows() {
  return { __rowsToObjectsCache: {} };
}

function getRowsToObjectsCacheBucket(contexto) {
  if (!contexto || typeof contexto !== 'object') return null;
  if (!contexto.__rowsToObjectsCache || typeof contexto.__rowsToObjectsCache !== 'object') {
    contexto.__rowsToObjectsCache = {};
  }
  return contexto.__rowsToObjectsCache;
}

function getRowsToObjectsCacheKey(sheet) {
  if (!sheet) return '';
  const planilha = sheet.getParent();
  const planilhaId = planilha && typeof planilha.getId === 'function'
    ? String(planilha.getId() || '').trim()
    : '';
  const aba = typeof sheet.getName === 'function'
    ? String(sheet.getName() || '').trim()
    : '';
  if (!aba) return '';
  return `${planilhaId || 'SEM_ID'}:${aba}`;
}

function cloneRowsToObjects(rows) {
  return (Array.isArray(rows) ? rows : []).map(row => ({ ...row }));
}

function rowsToObjects(sheet, opcoes) {
  if (!sheet) return [];

  const opts = opcoes || {};
  const contexto = (opts.context && typeof opts.context === 'object')
    ? opts.context
    : null;
  const cacheBucket = getRowsToObjectsCacheBucket(contexto);
  const cacheKey = cacheBucket ? getRowsToObjectsCacheKey(sheet) : '';

  if (cacheBucket && cacheKey && Array.isArray(cacheBucket[cacheKey])) {
    return cloneRowsToObjects(cacheBucket[cacheKey]);
  }

  const data = sheet.getDataRange().getValues();
  if (!Array.isArray(data) || data.length === 0) {
    return [];
  }

  const headers = Array.isArray(data[0]) ? data[0] : [];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });

  if (cacheBucket && cacheKey) {
    cacheBucket[cacheKey] = rows;
  }

  return cloneRowsToObjects(rows);
}

function normalizeIdValue(value) {
  return String(value ?? '').trim();
}

function getRowLookupCacheKey(sheetName, idField, idNorm) {
  const planilhaId = (typeof getDataSpreadsheetIdAtivo === 'function')
    ? String(getDataSpreadsheetIdAtivo({ skipAccessCheck: true }) || '').trim()
    : String((typeof DATA_SPREADSHEET_ID === 'string' ? DATA_SPREADSHEET_ID : '') || '').trim();
  const aba = String(sheetName || '').trim().toUpperCase();
  const campo = String(idField || '').trim().toUpperCase();
  const id = normalizeIdValue(idNorm);
  return `ROW_LOOKUP:${planilhaId || 'SEM_ID'}:${ROW_LOOKUP_CACHE_VERSION}:${aba}:${campo}:${id}`;
}

function lerRowLookupCache(sheetName, idField, idNorm) {
  const key = getRowLookupCacheKey(sheetName, idField, idNorm);
  try {
    const raw = CacheService.getScriptCache().get(key);
    if (!raw) return 0;
    const row = Number(raw);
    if (!Number.isFinite(row) || row < 2) return 0;
    return Math.floor(row);
  } catch (error) {
    return 0;
  }
}

function salvarRowLookupCache(sheetName, idField, idNorm, rowIndex) {
  const key = getRowLookupCacheKey(sheetName, idField, idNorm);
  const row = Number(rowIndex);
  if (!Number.isFinite(row) || row < 2) return;
  try {
    CacheService.getScriptCache().put(
      key,
      String(Math.floor(row)),
      ROW_LOOKUP_CACHE_TTL_SEC
    );
  } catch (error) {
    // sem acao
  }
}

function limparRowLookupCache(sheetName, idField, idNorm) {
  const key = getRowLookupCacheKey(sheetName, idField, idNorm);
  try {
    CacheService.getScriptCache().remove(key);
  } catch (error) {
    // sem acao
  }
}

function encontrarLinhaPorId(sheet, sheetName, idField, idCol, idNorm) {
  const idAlvo = normalizeIdValue(idNorm);
  const lastRow = sheet.getLastRow();
  if (!idAlvo || lastRow < 2 || idCol < 0) return 0;

  const rowCache = lerRowLookupCache(sheetName, idField, idAlvo);
  if (rowCache >= 2 && rowCache <= lastRow) {
    const valorLinhaCache = normalizeIdValue(sheet.getRange(rowCache, idCol + 1, 1, 1).getValue());
    if (valorLinhaCache === idAlvo) {
      return rowCache;
    }
    limparRowLookupCache(sheetName, idField, idAlvo);
  }

  const totalLinhasDados = lastRow - 1;
  if (totalLinhasDados <= 0) return 0;

  try {
    const colunaIds = sheet.getRange(2, idCol + 1, totalLinhasDados, 1);
    const encontrado = colunaIds
      .createTextFinder(idAlvo)
      .matchEntireCell(true)
      .matchCase(true)
      .findNext();
    if (encontrado) {
      const rowIndex = encontrado.getRow();
      salvarRowLookupCache(sheetName, idField, idAlvo, rowIndex);
      return rowIndex;
    }
  } catch (error) {
    // fallback para varredura manual abaixo
  }

  const valoresColunaId = sheet.getRange(2, idCol + 1, totalLinhasDados, 1).getValues();
  for (let i = 0; i < valoresColunaId.length; i++) {
    const valorLinha = normalizeIdValue(valoresColunaId[i][0]);
    if (valorLinha === idAlvo) {
      const rowIndex = i + 2;
      salvarRowLookupCache(sheetName, idField, idAlvo, rowIndex);
      return rowIndex;
    }
  }

  return 0;
}

function insertMany(sheetName, payloads, schema) {
  if (typeof assertCanWrite === 'function') {
    assertCanWrite(`Criacao em lote na aba ${String(sheetName || '').trim().toUpperCase() || 'SEM_NOME'}`);
  }

  const lista = Array.isArray(payloads) ? payloads : [];
  if (lista.length === 0) return 0;

  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  ensureSchema(sheet, schema);
  const headerMap = getHeaderMap(sheet);
  const totalColunas = sheet.getLastColumn();
  if (totalColunas <= 0) return 0;

  const rows = lista.map(payload => {
    const row = Array(totalColunas).fill('');
    const item = payload || {};
    schema.forEach(key => {
      if (key in headerMap) {
        row[headerMap[key]] = item[key] ?? '';
      }
    });
    return row;
  });

  if (rows.length === 0) return 0;

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, totalColunas).setValues(rows);
  invalidarCachesRelacionadosAba(sheetName);
  return rows.length;
}

function insert(sheetName, payload, schema) {
  if (typeof assertCanWrite === 'function') {
    assertCanWrite(`Criacao na aba ${String(sheetName || '').trim().toUpperCase() || 'SEM_NOME'}`);
  }

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
  if (typeof assertCanWrite === 'function') {
    assertCanWrite(`Atualizacao na aba ${String(sheetName || '').trim().toUpperCase() || 'SEM_NOME'}`);
  }

  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  ensureSchema(sheet, schema);
  const lastCol = sheet.getLastColumn();
  if (lastCol <= 0) return false;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headerMap = buildHeaderMapFromHeaders(headers);
  const idCol = headers.indexOf(idField);
  const idAlvo = normalizeIdValue(id);

  if (idCol === -1 || !idAlvo) return false;

  const rowIndex = encontrarLinhaPorId(sheet, sheetName, idField, idCol, idAlvo);
  if (rowIndex < 2) return false;

  const row = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  const idAnterior = normalizeIdValue(row[idCol]);
  let alterou = false;

  Object.keys(payload || {}).forEach(key => {
    if (key in headerMap) {
      row[headerMap[key]] = payload[key];
      alterou = true;
    }
  });

  if (alterou) {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  }

  const idNovo = normalizeIdValue(row[idCol]);
  if (idAnterior && idAnterior !== idNovo) {
    limparRowLookupCache(sheetName, idField, idAnterior);
  }
  if (idNovo) {
    salvarRowLookupCache(sheetName, idField, idNovo, rowIndex);
  }

  invalidarCachesRelacionadosAba(sheetName);
  return true;
}

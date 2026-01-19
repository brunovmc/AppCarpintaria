function getSheet(nome) {
  return SpreadsheetApp.getActive().getSheetByName(nome);
}

function getHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => map[h] = i);
  return map;
}

function ensureSchema(sheet, schema) {
  if (!sheet || !Array.isArray(schema)) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
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
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  ensureSchema(sheet, schema);
  const headerMap = getHeaderMap(sheet);
  const row = Array(sheet.getLastColumn()).fill('');

  schema.forEach(key => {
    if (key in headerMap) {
      row[headerMap[key]] = payload[key] ?? '';
    }
  });

  sheet.appendRow(row);
  return true;
}


function updateById(sheetName, idField, id, payload, schema) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  ensureSchema(sheet, schema);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf(idField);

  if (idCol === -1) return false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) {
      const headerMap = getHeaderMap(sheet);

      Object.keys(payload).forEach(key => {
        if (key in headerMap) {
          sheet.getRange(i + 1, headerMap[key] + 1).setValue(payload[key]);
        }
      });

      return true;
    }
  }
  return false;
}

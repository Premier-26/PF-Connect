/**
 * PF Connect — Google Sheets backend
 * Paste this entire file into Extensions > Apps Script > Code.gs
 * Then Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone
 */

var PHOTO_FOLDER_NAME = 'PF Connect Photos';

function doGet(e) {
  try {
    var action = e.parameter.action || 'list';
    var sheetName = e.parameter.sheet;
    if (action === 'list') {
      return jsonOut(listRows(sheetName));
    }
    return jsonOut({ error: 'Unknown action' });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var sheetName = body.sheet;

    if (action === 'add') {
      return jsonOut(addRow(sheetName, body.data));
    }
    if (action === 'upsert') {
      return jsonOut(upsertRow(sheetName, body.id, body.data));
    }
    if (action === 'bulkUpsert') {
      return jsonOut(bulkUpsert(sheetName, body.items));
    }
    if (action === 'update') {
      return jsonOut(updateRow(sheetName, body.id, body.data));
    }
    if (action === 'delete') {
      return jsonOut(deleteRow(sheetName, body.id));
    }
    if (action === 'uploadPhoto') {
      return jsonOut(uploadPhoto(body.filename, body.mimeType, body.base64));
    }
    return jsonOut({ error: 'Unknown action' });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

/* ---------- sheet helpers ---------- */
function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}

function getHeaders(sh) {
  var lastCol = sh.getLastColumn();
  return sh.getRange(1, 1, 1, lastCol).getValues()[0];
}

function rowToObject(headers, row) {
  var obj = {};
  for (var i = 0; i < headers.length; i++) obj[headers[i]] = row[i];
  return obj;
}

function listRows(sheetName) {
  var sh = getSheet(sheetName);
  var headers = getHeaders(sh);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var out = [];
  for (var r = 0; r < values.length; r++) {
    if (values[r][0] === '' || values[r][0] === null) continue; // skip blank rows
    out.push(rowToObject(headers, values[r]));
  }
  return out;
}

function findRowIndexById(sh, headers, id) {
  var idCol = headers.indexOf('ID');
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sh.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // sheet row number
  }
  return -1;
}

/** Insert if the ID doesn't exist yet, otherwise overwrite that row. Never creates a duplicate. */
function upsertRow(sheetName, id, data) {
  var sh = getSheet(sheetName);
  var headers = getHeaders(sh);
  var rowNum = findRowIndexById(sh, headers, id);
  var merged = Object.assign({ ID: id }, data);
  if (!merged.CreatedAt) merged.CreatedAt = new Date().getTime();
  var row = headers.map(function(h) { return (merged[h] !== undefined ? merged[h] : ''); });
  if (rowNum === -1) {
    sh.appendRow(row);
  } else {
    sh.getRange(rowNum, 1, 1, headers.length).setValues([row]);
  }
  return merged;
}

/** Same as upsertRow but for many records in one call — much faster for a full export. */
function bulkUpsert(sheetName, items) {
  var sh = getSheet(sheetName);
  var headers = getHeaders(sh);
  var idCol = headers.indexOf('ID');
  var lastRow = sh.getLastRow();
  var idToRow = {};
  if (lastRow >= 2) {
    var ids = sh.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] !== '' && ids[i][0] !== null) idToRow[String(ids[i][0])] = i + 2;
    }
  }
  var toAppend = [];
  items.forEach(function(item) {
    var merged = Object.assign({ ID: item.id }, item.data);
    if (!merged.CreatedAt) merged.CreatedAt = new Date().getTime();
    var row = headers.map(function(h) { return (merged[h] !== undefined ? merged[h] : ''); });
    var existingRow = idToRow[String(item.id)];
    if (existingRow) {
      sh.getRange(existingRow, 1, 1, headers.length).setValues([row]);
    } else {
      toAppend.push(row);
    }
  });
  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, headers.length).setValues(toAppend);
  }
  return { updated: items.length - toAppend.length, added: toAppend.length };
}

function addRow(sheetName, data) {
  var sh = getSheet(sheetName);
  var headers = getHeaders(sh);
  var id = Utilities.getUuid();
  data.ID = id;
  data.CreatedAt = new Date().getTime();
  var row = headers.map(function(h) { return (data[h] !== undefined ? data[h] : ''); });
  sh.appendRow(row);
  return data;
}

function updateRow(sheetName, id, data) {
  var sh = getSheet(sheetName);
  var headers = getHeaders(sh);
  var rowNum = findRowIndexById(sh, headers, id);
  if (rowNum === -1) throw new Error('Row not found: ' + id);
  var current = rowToObject(headers, sh.getRange(rowNum, 1, 1, headers.length).getValues()[0]);
  var merged = Object.assign({}, current, data);
  var row = headers.map(function(h) { return (merged[h] !== undefined ? merged[h] : ''); });
  sh.getRange(rowNum, 1, 1, headers.length).setValues([row]);
  return merged;
}

function deleteRow(sheetName, id) {
  var sh = getSheet(sheetName);
  var headers = getHeaders(sh);
  var rowNum = findRowIndexById(sh, headers, id);
  if (rowNum === -1) return { deleted: false };
  sh.deleteRow(rowNum);
  return { deleted: true };
}

/* ---------- photo upload ---------- */
function getPhotoFolder() {
  var folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function uploadPhoto(filename, mimeType, base64) {
  var folder = getPhotoFolder();
  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, mimeType, filename);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var directUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();
  return { id: file.getId(), url: directUrl };
}

/* ---------- output ---------- */
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

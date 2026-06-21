const SHEET_ID = '1ODcMNnguu-pHW6BK7zbOGSMRYkgwQiDac5YPsjS7Smc';
const SHEET_NAME = 'Songs';
const REQUIRED_COLUMNS = ['id', 'title', 'key', 'body'];
const OPTIONAL_COLUMNS = ['youtubeurl'];

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || 'list').trim().toLowerCase();
  try {
    if (action === 'list') return jsonOutput({ ok: true, songs: listSongs_() });
    if (action === 'editor-status') return jsonOutput(getEditorStatus_());
    return jsonOutput({ ok: false, error: `Acción GET no soportada: ${action}` });
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message || String(err) });
  }
}

function doPost(e) {
  try {
    const payload = parseRequestBody_(e);
    const action = String(payload.action || '').trim().toLowerCase();
    if (action !== 'update') {
      return jsonOutput({ ok: false, error: `Acción POST no soportada: ${action || '(vacía)'}` });
    }
    const song = sanitizeSongPayload_(payload.song || payload);
    const updated = updateSong_(song);
    return jsonOutput({ ok: true, song: updated });
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message || String(err) });
  }
}

function parseRequestBody_(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  return JSON.parse(raw);
}

function getSpreadsheetContext_() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`No existe la pestaña "${SHEET_NAME}".`);
  const file = DriveApp.getFileById(SHEET_ID);
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('La hoja Songs está vacía.');

  const headers = values[0].map(header => String(header || '').trim().toLowerCase());
  const columnIndex = {};
  headers.forEach((header, index) => { if (header) columnIndex[header] = index; });
  REQUIRED_COLUMNS.forEach(column => {
    if (columnIndex[column] === undefined) throw new Error(`Falta la columna requerida "${column}".`);
  });

  return { spreadsheet, sheet, file, values, headers, columnIndex };
}

function getCurrentUserEmail_() {
  return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
}

function canCurrentUserEdit_(file) {
  const email = getCurrentUserEmail_();
  if (!email) return false;
  const ownerEmail = file.getOwner() ? String(file.getOwner().getEmail() || '').trim().toLowerCase() : '';
  if (ownerEmail && ownerEmail === email) return true;
  return file.getEditors().some(user => String(user.getEmail() || '').trim().toLowerCase() === email);
}

function getEditorStatus_() {
  const ctx = getSpreadsheetContext_();
  const currentUserEmail = getCurrentUserEmail_();
  return {
    ok: true,
    canEdit: canCurrentUserEdit_(ctx.file),
    currentUserEmail: currentUserEmail
  };
}

function listSongs_() {
  const ctx = getSpreadsheetContext_();
  const rows = ctx.values.slice(1);
  return rows.map((row, rowIndex) => {
    const visibleIndex = ctx.columnIndex.visible;
    const visibleValue = visibleIndex === undefined ? 'TRUE' : String(row[visibleIndex] || 'TRUE').trim().toUpperCase();
    if (visibleValue === 'FALSE' || visibleValue === 'NO' || visibleValue === '0') return null;
    return sanitizeSongPayload_({
      id: row[ctx.columnIndex.id],
      title: row[ctx.columnIndex.title],
      key: row[ctx.columnIndex.key],
      body: row[ctx.columnIndex.body],
      youtubeUrl: ctx.columnIndex.youtubeurl === undefined ? '' : row[ctx.columnIndex.youtubeurl]
    }, rowIndex);
  }).filter(Boolean);
}

function sanitizeSongPayload_(song, fallbackIndex) {
  const title = String(song && song.title || '').trim();
  const key = String(song && song.key || '').trim();
  const body = String(song && song.body || '').replace(/\r\n/g, '\n').trim();
  const youtubeUrl = String(song && (song.youtubeUrl !== undefined ? song.youtubeUrl : song.youtubeurl) || '').trim();
  let id = String(song && song.id || title || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
  if (!id && typeof fallbackIndex === 'number') id = `song-${fallbackIndex + 1}`;
  if (!id || !title || !key || !body) throw new Error('La canción debe incluir id, title, key y body.');
  return { id, title, key, body, youtubeUrl };
}

function updateSong_(song) {
  const ctx = getSpreadsheetContext_();
  if (!canCurrentUserEdit_(ctx.file)) {
    throw new Error('Tu cuenta no tiene permiso de edición sobre esta Google Sheet.');
  }

  const rowIndex = findSongRowIndex_(ctx, song.id);
  if (rowIndex < 0) throw new Error(`No encontré una canción con id "${song.id}".`);

  setCellValue_(ctx.sheet, rowIndex, ctx.columnIndex.title, song.title);
  setCellValue_(ctx.sheet, rowIndex, ctx.columnIndex.key, song.key);
  setCellValue_(ctx.sheet, rowIndex, ctx.columnIndex.body, song.body);
  if (ctx.columnIndex.youtubeurl !== undefined) {
    setCellValue_(ctx.sheet, rowIndex, ctx.columnIndex.youtubeurl, song.youtubeUrl);
  }

  SpreadsheetApp.flush();
  return song;
}

function findSongRowIndex_(ctx, songId) {
  const normalizedId = String(songId || '').trim().toLowerCase();
  for (let row = 1; row < ctx.values.length; row++) {
    const currentId = String(ctx.values[row][ctx.columnIndex.id] || '').trim().toLowerCase();
    if (currentId === normalizedId) return row + 1;
  }
  return -1;
}

function setCellValue_(sheet, rowIndex, columnIndex, value) {
  if (columnIndex === undefined) return;
  sheet.getRange(rowIndex, columnIndex + 1).setValue(value);
}

function jsonOutput(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

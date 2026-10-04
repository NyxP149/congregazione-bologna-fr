const SHEET_ID = '1pgvSXZz7VgPbxhAsVpnMQVL1z_aW28yAFWV286jdxHw';

// ── Récupère le mot de passe depuis Script Properties ────────
function getEditorPassword() {
  const pwd = PropertiesService.getScriptProperties().getProperty('EDITOR_PWD');
  if (!pwd) throw new Error('EDITOR_PWD non configurato nelle Script Properties');
  return pwd;
}

// ── Hash SHA-256 via Utilities.computeDigest ─────────────────
function sha256(str) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    str,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// Compare les hash sans sortie anticipée (temps constant)
function checkPassword(pwd) {
  const a = sha256(String(pwd || ''));
  const b = sha256(getEditorPassword());
  let diff = a.length ^ b.length;
  for (let i = 0; i < a.length && i < b.length; i++) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  const ok = diff === 0;
  if (!ok) Utilities.sleep(1500); // ralentit les essais de force brute
  return ok;
}

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === 'Mesi')    sheet.getRange(1,1,1,2).setValues([['mese','data_json']]);
    if (name === 'Annuale') sheet.getRange(1,1,1,1).setValues([['data_json']]);
    if (name === 'Meta')    sheet.getRange(1,1,1,2).setValues([['key','value']]);
  }
  return sheet;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET : uniquement un ping, aucune donnée, aucun mot de passe ──
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  if (action === 'ping') {
    return jsonResponse({ ok: true, message: 'Script attivo v10 ✓' });
  }
  return jsonResponse({ ok: false, error: 'Usa POST' });
}

// ── POST : toutes les actions, mot de passe dans le corps ────
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try { payload = JSON.parse(e.postData.contents); } catch (err) {}
    }
    const action = payload.action || '';

    if (!checkPassword(payload.pwd)) {
      if (action === 'verifyPassword') {
        return jsonResponse({ ok: true, authorized: false });
      }
      return jsonResponse({ ok: false, error: 'Non autorizzato' });
    }

    if (action === 'verifyPassword') return jsonResponse({ ok: true, authorized: true });
    if (action === 'loadAll')        return loadAll();
    if (action === 'saveAll')        return saveAll(payload);

    return jsonResponse({ ok: false, error: 'Azione sconosciuta: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ── HELPER: résout la clé "mese" même si Sheets l'a converti en Date ──
function resolveMeseKey(cellVal, parsed, tz) {
  if (parsed && typeof parsed === 'object' &&
      typeof parsed.mese === 'string' && /^\d{4}-\d{2}$/.test(parsed.mese.trim())) {
    return parsed.mese.trim();
  }
  if (typeof cellVal === 'string' && /^\d{4}-\d{2}$/.test(cellVal.trim())) {
    return cellVal.trim();
  }
  if (cellVal instanceof Date) {
    return Utilities.formatDate(cellVal, tz, 'yyyy-MM');
  }
  if (typeof cellVal === 'string') {
    const d = new Date(cellVal);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, tz, 'yyyy-MM');
    }
  }
  return null;
}

// ── LOAD ALL ─────────────────────────────────────────────────
function loadAll() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const sheetMesi    = getOrCreateSheet('Mesi');
  const sheetAnnuale = getOrCreateSheet('Annuale');

  const months = {};
  const lastRowMesi = sheetMesi.getLastRow();
  if (lastRowMesi >= 2) {
    const rows = sheetMesi.getRange(2, 1, lastRowMesi - 1, 2).getValues();
    rows.forEach(function(row) {
      const json = row[1] ? String(row[1]).trim() : '';
      if (!json) return;
      let parsed;
      try { parsed = JSON.parse(json); } catch(e) { return; }
      const mese = resolveMeseKey(row[0], parsed, tz);
      if (!mese) return;
      if (parsed && typeof parsed === 'object') parsed.mese = mese;
      months[mese] = parsed;
    });
  }

  let annual = [];
  const lastRowAnn = sheetAnnuale.getLastRow();
  if (lastRowAnn >= 2) {
    const val = sheetAnnuale.getRange(2, 1).getValue();
    if (val) { try { annual = JSON.parse(String(val)); } catch(e) {} }
  }

  return jsonResponse({ ok: true, months: months, annual: annual });
}

// ── SAVE ALL ─────────────────────────────────────────────────
function saveAll(payload) {
  const months = payload.months || {};
  const annual = payload.annual || [];

  const sheetMesi = getOrCreateSheet('Mesi');
  sheetMesi.getRange(1,1,1,2).setValues([['mese','data_json']]);
  const maxRows = sheetMesi.getMaxRows();
  if (maxRows >= 2) {
    sheetMesi.getRange(2,1,maxRows-1,2).clearContent();
    sheetMesi.getRange(2,1,maxRows-1,1).setNumberFormat('@');
  }
  const rows = Object.keys(months).sort().map(function(mese) {
    return [mese, JSON.stringify(months[mese])];
  });
  if (rows.length > 0) {
    sheetMesi.getRange(2,1,rows.length,1).setNumberFormat('@');
    sheetMesi.getRange(2,1,rows.length,2).setValues(rows);
  }

  const sheetAnnuale = getOrCreateSheet('Annuale');
  sheetAnnuale.getRange(1,1).setValue('data_json');
  const lastRowAnn = sheetAnnuale.getLastRow();
  if (lastRowAnn >= 2) sheetAnnuale.getRange(2,1,lastRowAnn-1,1).clearContent();
  if (annual.length > 0) sheetAnnuale.getRange(2,1).setValue(JSON.stringify(annual));

  return jsonResponse({ ok: true, saved: Object.keys(months).length + ' mesi salvati' });
}

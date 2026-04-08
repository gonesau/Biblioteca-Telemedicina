// ============================================================
//  Setup.gs — Biblioteca de Telemedicina
//  Crea todas las hojas necesarias con sus encabezados.
//  Ejecuta setupSheets() una sola vez antes de usar el sistema.
// ============================================================

function setupSheets() {
  var ss = getSpreadsheet_();

  for (var key in SHEET_NAMES) {
    var name    = SHEET_NAMES[key];
    var headers = SHEET_HEADERS[key];
    if (!headers) continue;

    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    ensureHeaders_(sheet, headers);
  }

  SpreadsheetApp.getActive().toast('✓ Hojas inicializadas correctamente', 'Biblioteca Telemedicina', 5);
  Logger.log('setupSheets completado: ' + Object.keys(SHEET_NAMES).length + ' hojas creadas/verificadas.');
}

// ============================================================
//  Utils.gs — funciones utilitarias compartidas
// ============================================================

function ensureHeaders_(sheet, headers) {
  if (!sheet || !headers || !headers.length) return;
  var first = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var needsHeaders = first.every(function(v) { return !v; });
  if (needsHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function getRows_(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1); // omite la fila de encabezados
}

function getHeaderIndexMap_(sheet, headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[headers[i]] = i;
  }
  return map;
}

function findRowIndexByKey_(sheet, keyName, keyValue, map) {
  var rows = getRows_(sheet);
  var idx  = map[keyName];
  for (var i = 0; i < rows.length; i++) {
    if ((rows[i][idx] || '').toString() === keyValue.toString()) {
      return i + 2; // +1 por header, +1 por índice base 1
    }
  }
  return -1;
}

function getOrCreateSheetByName_(name) {
  var ss    = getSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function uuid_() {
  return Utilities.getUuid();
}

function now_() {
  return new Date();
}

function toIsoString_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function slugify_(text) {
  return (text || '')
    .toString()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')  // quitar tildes
    .replace(/[^a-z0-9\s-]/g, '')
    .trim()
    .replace(/[\s]+/g, '-')
    .replace(/-+/g, '-');
}

function ensureUniqueSlug_(slug) {
  var sheet   = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLES);
  var hdrs    = SHEET_HEADERS.ARTICLES;
  var map     = getHeaderIndexMap_(sheet, hdrs);
  var rows    = getRows_(sheet);
  var slugIdx = map['slug'];
  var existing = {};
  for (var i = 0; i < rows.length; i++) {
    var s = rows[i][slugIdx];
    if (s) existing[s] = true;
  }
  if (!existing[slug]) return slug;
  var counter = 2;
  while (existing[slug + '-' + counter]) counter++;
  return slug + '-' + counter;
}

function sanitizeHtml_(html) {
  // Permite tags seguros básicos; elimina scripts y atributos peligrosos
  return (html || '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/on\w+="[^"]*"/gi, '')
    .replace(/on\w+='[^']*'/gi, '');
}

function extractDriveFileIdFromAny_(input) {
  if (!input) return '';
  var str = input.toString().trim();
  // Si ya es un ID limpio (sin slashes ni params)
  if (/^[a-zA-Z0-9_-]{10,}$/.test(str)) return str;
  // Extraer de URL de Drive
  var match = str.match(/[-\w]{25,}/);
  return match ? match[0] : str;
}

function getActiveUserEmail_() {
  return Session.getActiveUser().getEmail();
}

function isEmailAllowed_(email) {
  if (!email) return false;
  var cfg = getAccessConfig();
  if (cfg.allowedEmails && cfg.allowedEmails.indexOf(email) !== -1) return true;
  if (cfg.allowedDomains) {
    for (var i = 0; i < cfg.allowedDomains.length; i++) {
      if (email.endsWith(cfg.allowedDomains[i])) return true;
    }
  }
  return false;
}

function getCachedJSON_(key) {
  var cache = CacheService.getScriptCache();
  var raw   = cache.get(key);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function setCachedJSON_(key, value, ttlSeconds) {
  var cache = CacheService.getScriptCache();
  try { cache.put(key, JSON.stringify(value), ttlSeconds || 300); } catch (e) {}
}

function invalidateArticlesCache_() {
  var cache = CacheService.getScriptCache();
  try {
    cache.remove(CACHE_KEY_ARTICLES + '_all');
    cache.remove(CACHE_KEY_ARTICLES + '_pub');
  } catch (e) {}
}

function tryWithAudit_(fn, context) {
  try {
    return fn();
  } catch (e) {
    try {
      logAudit_(
        context.entityType || 'unknown',
        context.entityId   || '',
        'ERROR',
        getActiveUserEmail_(),
        { message: e && e.message ? e.message : String(e) }
      );
    } catch (_) {}
    throw e;
  }
}



function runImport() {
  importFromDriveFolder(
    '1ZsTIovrGb8IrfGv3MxVYJorbHvpW7XDc'
  );
}

// Constantes de caché
var CACHE_KEY_ARTICLES = 'articles_list';
var CACHE_TTL_SECONDS  = 300;
// ============================================================
//  Utils.gs — Biblioteca de Telemedicina
//  FUENTE ÚNICA de funciones utilitarias. No duplicar en otros archivos.
// ============================================================

// --- Constantes de caché ---
var CACHE_KEY_ARTICLES = 'wiki_articles_v2';
var CACHE_TTL_SECONDS  = 300; // 5 minutos

// --- Date / Time ---
function now_() {
  return new Date();
}

function toIsoString_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd'T'HH:mm:ss"
  );
}

// --- Identificadores ---
function uuid_() {
  return Utilities.getUuid();
}

// --- Texto ---
function slugify_(text) {
  if (!text) return '';
  var slug = text.toString().toLowerCase();
  var accents   = 'áéíóúýñäëïöüâêîôûàèìòùãõ';
  var noAccents = 'aeiouynaeiouaeiouaeiouao';
  for (var i = 0; i < accents.length; i++) {
    slug = slug.replace(new RegExp(accents.charAt(i), 'g'), noAccents.charAt(i));
  }
  return slug
    .replace(/[^a-z0-9\s-]/g, '')
    .trim()
    .replace(/[\s_]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function sanitizeHtml_(html) {
  if (!html) return '';
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/ on[a-z]+\s*=\s*"[^"]*"/gi, '')
    .replace(/ on[a-z]+\s*=\s*'[^']*'/gi, '')
    .replace(/ on[a-z]+\s*=\s*[^\s>]+/gi, '');
}

function extractDriveFileIdFromAny_(input) {
  if (!input) return '';
  var trimmed = String(input).trim();
  // Si ya es un ID limpio (sin slashes ni protocolo)
  if (!/[\/:]/.test(trimmed)) return trimmed;
  // /file/d/{id} o /document|spreadsheets|presentation/d/{id}
  var m = trimmed.match(/\/(file|document|spreadsheets|presentation)\/d\/([^\/\?#]+)/);
  if (m) return m[2];
  // open?id={id}
  var mid = trimmed.match(/[?&]id=([^&#]+)/);
  if (mid) return mid[1];
  return trimmed;
}

// --- Sesión / Acceso ---
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

// --- Hojas de cálculo ---
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
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function getHeaderIndexMap_(sheet, headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[headers[i]] = i;
  }
  return map;
}

function findRowIndexByKey_(sheet, keyHeader, keyValue, headerMap) {
  var rows = getRows_(sheet);
  var idx = headerMap[keyHeader];
  for (var i = 0; i < rows.length; i++) {
    if ((rows[i][idx] || '').toString() === (keyValue || '').toString()) {
      return i + 2; // +1 encabezado, +1 base-1
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

function ensureUniqueSlug_(slugBase) {
  var sheet   = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLES);
  var hdrs    = SHEET_HEADERS.ARTICLES;
  var map     = getHeaderIndexMap_(sheet, hdrs);
  var rows    = getRows_(sheet);
  var slugIdx = map['slug'];
  var existing = {};
  for (var i = 0; i < rows.length; i++) {
    var s = (rows[i][slugIdx] || '').toString();
    if (s) existing[s] = true;
  }
  if (!existing[slugBase]) return slugBase;
  var n = 2;
  while (existing[slugBase + '-' + n]) n++;
  return slugBase + '-' + n;
}

// --- Error handling con auditoría ---
function tryWithAudit_(fn, context) {
  try {
    return fn();
  } catch (err) {
    try {
      logAudit_(
        (context && context.entityType) || 'unknown',
        (context && context.entityId)   || '',
        'ERROR',
        getActiveUserEmail_(),
        { message: err && err.message, stack: err && err.stack }
      );
    } catch (_) {}
    throw err;
  }
}

// --- CacheService (chunked para payloads grandes > 100KB) ---
function getCachedJSON_(key) {
  try {
    var cache  = CacheService.getScriptCache();
    var chunk0 = cache.get(key + '_0');
    if (!chunk0) return null;
    var chunks = [chunk0];
    for (var i = 1; i < 20; i++) {
      var c = cache.get(key + '_' + i);
      if (!c) break;
      chunks.push(c);
    }
    return JSON.parse(chunks.join(''));
  } catch (e) {
    return null;
  }
}

function setCachedJSON_(key, obj, ttl) {
  try {
    var cache     = CacheService.getScriptCache();
    var json      = JSON.stringify(obj);
    var chunkSize = 90000; // límite de CacheService ~100KB por clave
    var numChunks = Math.ceil(json.length / chunkSize);
    var pairs     = {};
    for (var i = 0; i < numChunks; i++) {
      pairs[key + '_' + i] = json.substring(i * chunkSize, (i + 1) * chunkSize);
    }
    cache.putAll(pairs, ttl || CACHE_TTL_SECONDS);
  } catch (e) {
    Logger.log('setCachedJSON_ error: ' + (e && e.message));
  }
}

function invalidateArticlesCache_() {
  try {
    var cache = CacheService.getScriptCache();
    var keys  = [];
    for (var i = 0; i < 20; i++) {
      keys.push(CACHE_KEY_ARTICLES + '_all_' + i);
      keys.push(CACHE_KEY_ARTICLES + '_pub_' + i);
    }
    cache.removeAll(keys);
  } catch (e) {}
}
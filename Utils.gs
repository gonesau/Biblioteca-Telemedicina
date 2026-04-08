function now_() {
  return new Date();
}

function uuid_() {
  return Utilities.getUuid();
}

function slugify_(text) {
  if (!text) return '';
  var slug = text.toString().toLowerCase();
  
  // Reemplazo manual de tildes compatible con el motor antiguo de Google (Rhino)
  var accents = 'áéíóúýñäëïöüâêîôû';
  var noAccents = 'aeiouynaeiouaeiou';
  for (var i = 0; i < accents.length; i++) {
    slug = slug.replace(new RegExp(accents.charAt(i), 'g'), noAccents.charAt(i));
  }

  slug = slug
    .replace(/[^a-z0-9\s-]/g, '')
    .trim()
    .replace(/[\s_-]+/g, '-')
    .replace(/^-+|-+$/g, '');
  return slug;
}

function getActiveUserEmail_() {
  return Session.getActiveUser().getEmail();
}

function isEmailAllowed_(email) {
  var cfg = getAccessConfig();
  if (!email) return false;
  if (cfg.allowedEmails.indexOf(email) !== -1) return true;
  for (var i = 0; i < cfg.allowedDomains.length; i++) {
    if (email.endsWith(cfg.allowedDomains[i])) return true;
  }
  return false;
}

function sanitizeHtml_(html) {
  if (!html) return '';
  // Remove script/style tags and event handlers as a baseline server-side sanitation
  var sanitized = html
    .replace(/<\/(?:script|style)>/gi, '')
    .replace(/<(script|style)[\s\S]*?>[\s\S]*?<\/\1>/gi, '')
    .replace(/ on[a-z]+\s*=\s*"[^"]*"/gi, '')
    .replace(/ on[a-z]+\s*=\s*'[^']*'/gi, '')
    .replace(/ on[a-z]+\s*=\s*[^\s>]+/gi, '');
  return sanitized;
}

function extractDriveFileIdFromAny_(input) {
  if (!input) return '';
  var trimmed = String(input).trim();
  // If looks like plain ID (no slash or scheme), return as-is
  if (!/[\/:]/.test(trimmed)) return trimmed;
  // /file/d/{id}/... or /document|spreadsheets|presentation/d/{id}/...
  var m = trimmed.match(/\/(file|document|spreadsheets|presentation)\/d\/([^\/\?]+)/);
  if (m) return m[2];
  // open?id={id}
  var mid = trimmed.match(/[?&]id=([^&]+)/);
  if (mid) return mid[1];
  return trimmed;
}

function getHeaderIndexMap_(sheet, expectedHeaders) {
  var header = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];
  var map = {};
  for (var i = 0; i < expectedHeaders.length; i++) {
    map[expectedHeaders[i]] = i;
  }
  return map;
}

function getRows_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function findRowIndexByKey_(sheet, keyHeader, keyValue, headerMap) {
  var rows = getRows_(sheet);
  var idx = headerMap[keyHeader];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][idx] === keyValue) return i + 2; // offset for header
  }
  return -1;
}

function ensureUniqueSlug_(slugBase) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLES);
  var hdrs = SHEET_HEADERS.ARTICLES;
  var map = getHeaderIndexMap_(sheet, hdrs);
  var slugIdx = map['slug'];
  var rows = getRows_(sheet);
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

function toIsoString_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function getBasePublicUrl_() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty('BASE_PUBLIC_URL') || '';
}

function tryWithAudit_(fn, context) {
  try {
    return fn();
  } catch (err) {
    try {
      logAudit_((context && context.entityType) || 'unknown', (context && context.entityId) || '', 'ERROR', getActiveUserEmail_(), {
        message: err && err.message,
        stack: err && err.stack,
        context: context || {}
      });
    } catch (ignored) {}
    throw err;
  }
}

// --- CacheService helpers ---
var CACHE_TTL_SECONDS = 300; // 5 minutes
var CACHE_KEY_ARTICLES = 'wiki_articles_v2';

function getCachedJSON_(key) {
  try {
    var cache = CacheService.getScriptCache();
    var chunks = [];
    var raw = cache.get(key + '_0');
    if (!raw) return null;
    chunks.push(raw);
    for (var i = 1; i < 20; i++) {
      var chunk = cache.get(key + '_' + i);
      if (!chunk) break;
      chunks.push(chunk);
    }
    return JSON.parse(chunks.join(''));
  } catch (e) {
    return null;
  }
}

function setCachedJSON_(key, obj, ttl) {
  try {
    var cache = CacheService.getScriptCache();
    var json = JSON.stringify(obj);
    // CacheService max value size is ~100KB, split into chunks
    var chunkSize = 90000;
    var numChunks = Math.ceil(json.length / chunkSize);
    var pairs = {};
    for (var i = 0; i < numChunks; i++) {
      pairs[key + '_' + i] = json.substring(i * chunkSize, (i + 1) * chunkSize);
    }
    cache.putAll(pairs, ttl || CACHE_TTL_SECONDS);
  } catch (e) {}
}

function invalidateArticlesCache_() {
  try {
    var cache = CacheService.getScriptCache();
    var keys = [];
    for (var i = 0; i < 20; i++) {
      keys.push(CACHE_KEY_ARTICLES + '_all_' + i);
      keys.push(CACHE_KEY_ARTICLES + '_pub_' + i);
    }
    cache.removeAll(keys);
  } catch (e) {}
}




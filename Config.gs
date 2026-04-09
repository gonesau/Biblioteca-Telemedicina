// ============================================================
//  Config.gs — Biblioteca de Telemedicina
// ============================================================

var DEFAULT_ACCESS_CONFIG = {
  adminEmails:       [],
  allowedDomains:    ['@telesalud.gob.sv', '@goes.gob.sv'],
  allowedEmails:     [],
  guestEditorEmails: []
};

var BIBLIOTECA_PUBLICA_URL = 'https://sites.google.com/telesalud.gob.sv/bibliotecatelemedicina/inicio';

function getAccessConfig() {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ACCESS_CONFIG);
  if (!sheet) return DEFAULT_ACCESS_CONFIG;

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return DEFAULT_ACCESS_CONFIG;

  var header = values[0];
  var row    = values[1];
  var idx    = {};
  for (var i = 0; i < header.length; i++) idx[header[i]] = i;

  function splitEmails(raw) {
    return (raw || '').toString().split(',')
      .map(function(s) { return s.trim(); })
      .filter(Boolean);
  }

  var adminEmails       = splitEmails(row[idx['adminEmail']]);
  var allowedDomains    = splitEmails(row[idx['allowedDomains']]);
  var allowedEmails     = splitEmails(row[idx['allowedEmails']]);
  var guestEditorEmails = splitEmails(row[idx['guestEditorEmails']]);

  return {
    adminEmails:       adminEmails.length       ? adminEmails       : DEFAULT_ACCESS_CONFIG.adminEmails,
    allowedDomains:    allowedDomains.length    ? allowedDomains    : DEFAULT_ACCESS_CONFIG.allowedDomains,
    allowedEmails:     allowedEmails.length     ? allowedEmails     : DEFAULT_ACCESS_CONFIG.allowedEmails,
    guestEditorEmails: guestEditorEmails.length ? guestEditorEmails : DEFAULT_ACCESS_CONFIG.guestEditorEmails
  };
}

function getNotificationsConfig() {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.NOTIFICATIONS_CONFIG);
  if (!sheet) return { recipients: [], replyTo: '', senderAlias: 'Biblioteca Telemedicina' };

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return { recipients: [], replyTo: '', senderAlias: 'Biblioteca Telemedicina' };

  var header = values[0];
  var idx    = {};
  for (var i = 0; i < header.length; i++) idx[header[i]] = i;

  var recipients  = [];
  var replyTo     = '';
  var senderAlias = 'Biblioteca Telemedicina';

  for (var r = 1; r < values.length; r++) {
    var row     = values[r];
    var enabled = (row[idx['enabled']] || '').toString().toUpperCase() === 'Y';
    if (enabled) {
      recipients.push((row[idx['recipientEmail']] || '').toString().trim());
      if (!replyTo)     replyTo     = (row[idx['replyTo']]     || '').toString().trim();
      if (!senderAlias) senderAlias = (row[idx['senderAlias']] || '').toString().trim();
    }
  }

  return {
    recipients:  recipients.filter(Boolean),
    replyTo:     replyTo,
    senderAlias: senderAlias || 'Biblioteca Telemedicina'
  };
}

/**
 * URL pública base del Web App.
 * Lee primero de la hoja SiteConfig (clave 'basePublicUrl'),
 * con fallback a BIBLIOTECA_PUBLICA_URL.
 * NOTA: Esta es la ÚNICA definición de getBasePublicUrl_() en el proyecto.
 */
function getBasePublicUrl_() {
  try {
    var ss    = getSpreadsheet_();
    var sheet = ss.getSheetByName(SHEET_NAMES.SITE_CONFIG);
    if (!sheet) return BIBLIOTECA_PUBLICA_URL;
    var hdr  = SHEET_HEADERS.SITE_CONFIG;
    var map  = getHeaderIndexMap_(sheet, hdr);
    var rows = getRows_(sheet);
    for (var i = 0; i < rows.length; i++) {
      if ((rows[i][map['key']] || '') === 'basePublicUrl') {
        return rows[i][map['value']] || BIBLIOTECA_PUBLICA_URL;
      }
    }
  } catch (e) {}
  return BIBLIOTECA_PUBLICA_URL;
}

function getEnvironment() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty('ENV') || 'production';
}

function isDev() {
  return getEnvironment() === 'development';
}
// Default access configuration. Can be overridden via ACCESS_CONFIG sheet.

var DEFAULT_ACCESS_CONFIG = {
  adminEmail: '', // set in sheet
  allowedDomains: ['@telesalud.gob.sv', '@goes.gob.sv'],
  allowedEmails: ['rodolfovargasoff@gmail.com', 'ia.rodolfovargas@ufg.edu.sv'],
  guestEditorEmails: []
};

function getAccessConfig() {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ACCESS_CONFIG);
  if (!sheet) return DEFAULT_ACCESS_CONFIG;

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return DEFAULT_ACCESS_CONFIG;
  var header = values[0];
  var row = values[1];
  var idx = {};
  for (var i = 0; i < header.length; i++) idx[header[i]] = i;

  var adminEmails = (row[idx['adminEmail']] || '').toString().split(',').map(function(s){return s.trim();}).filter(Boolean);
  var allowedDomains = (row[idx['allowedDomains']] || '').toString().split(',').map(function(s){return s.trim();}).filter(Boolean);
  var allowedEmails = (row[idx['allowedEmails']] || '').toString().split(',').map(function(s){return s.trim();}).filter(Boolean);
  var guestEditorEmails = (row[idx['guestEditorEmails']] || '').toString().split(',').map(function(s){return s.trim();}).filter(Boolean);

  return {
    adminEmails: adminEmails.length ? adminEmails : [DEFAULT_ACCESS_CONFIG.adminEmail],
    allowedDomains: allowedDomains.length ? allowedDomains : DEFAULT_ACCESS_CONFIG.allowedDomains,
    allowedEmails: allowedEmails.length ? allowedEmails : DEFAULT_ACCESS_CONFIG.allowedEmails,
    guestEditorEmails: guestEditorEmails.length ? guestEditorEmails : DEFAULT_ACCESS_CONFIG.guestEditorEmails
  };
}

function getNotificationsConfig() {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.NOTIFICATIONS_CONFIG);
  if (!sheet) return { recipients: [], replyTo: '', senderAlias: '' };
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return { recipients: [], replyTo: '', senderAlias: '' };
  var header = values[0];
  var idx = {};
  for (var i = 0; i < header.length; i++) idx[header[i]] = i;
  var recipients = [];
  var replyTo = '';
  var senderAlias = '';
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var enabled = (row[idx['enabled']] || '').toString().toUpperCase() === 'Y';
    if (enabled) {
      recipients.push((row[idx['recipientEmail']] || '').toString().trim());
      if (!replyTo) replyTo = (row[idx['replyTo']] || '').toString().trim();
      if (!senderAlias) senderAlias = (row[idx['senderAlias']] || '').toString().trim();
    }
  }
  return { recipients: recipients.filter(Boolean), replyTo: replyTo, senderAlias: senderAlias };
}



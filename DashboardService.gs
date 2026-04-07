function getDashboardStats() {
  // Solo administradores
  assertAccessOrThrow_(true);
  var ss = getSpreadsheet_();
  var articlesSheet = ss.getSheetByName(SHEET_NAMES.ARTICLES);
  var versionsSheet = ss.getSheetByName(SHEET_NAMES.ARTICLE_VERSIONS);
  var totalArticles = 0;
  if (articlesSheet) {
    var rows = getRows_(articlesSheet);
    totalArticles = rows.length;
  }
  var now = new Date();
  var sevenDaysAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  var recentEdits7d = 0;
  if (versionsSheet) {
    var hdrs = SHEET_HEADERS.ARTICLE_VERSIONS;
    var map = getHeaderIndexMap_(versionsSheet, hdrs);
    var vrows = getRows_(versionsSheet);
    for (var i = 0; i < vrows.length; i++) {
      var changedAtStr = vrows[i][map['changedAt']];
      var changeType = (vrows[i][map['changeType']] || '').toString();
      var changedAt = changedAtStr ? new Date(changedAtStr) : null;
      if (changedAt && changedAt >= sevenDaysAgo && changeType !== 'delete') {
        recentEdits7d++;
      }
    }
  }
  return {
    totalArticles: totalArticles,
    recentEdits7d: recentEdits7d
  };
}

function getRecentActivity(limit) {
  // Solo administradores
  assertAccessOrThrow_(true);
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);
  if (!sheet) return [];
  var hdrs = SHEET_HEADERS.AUDIT_LOG;
  var map = getHeaderIndexMap_(sheet, hdrs);
  var rows = getRows_(sheet);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var metaStr = r[map['metadata']];
    var meta = {};
    try { meta = metaStr ? JSON.parse(metaStr) : {}; } catch (e) { meta = {}; }
    out.push({
      at: r[map['at']],
      action: r[map['action']],
      entityType: r[map['entityType']],
      entityId: r[map['entityId']],
      actor: r[map['actor']],
      title: meta && meta.title || '',
      slug: meta && meta.slug || ''
    });
  }
  // Orden descendente por fecha (asumimos ISO); si no, invertimos por append
  out.sort(function(a, b) {
    var da = a.at ? new Date(a.at).getTime() : 0;
    var db = b.at ? new Date(b.at).getTime() : 0;
    return db - da;
  });
  var n = typeof limit === 'number' && limit > 0 ? limit : 10;
  return out.slice(0, n);
}

// --- Site metrics/config ---
function getSiteMetric(key) {
  return tryWithAudit_(function() {
    assertAccessOrThrow_(false);
    var k = (key || '').toString().trim();
    if (!k) return '';
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(SHEET_NAMES.SITE_CONFIG);
    if (!sheet) {
      // Crear hoja y headers si no existe aún (lectura resiliente)
      sheet = getOrCreateSheetByName_(SHEET_NAMES.SITE_CONFIG);
      ensureHeaders_(sheet, SHEET_HEADERS.SITE_CONFIG);
    }
    var hdr = SHEET_HEADERS.SITE_CONFIG;
    var map = getHeaderIndexMap_(sheet, hdr);
    var rows = getRows_(sheet);
    for (var i = 0; i < rows.length; i++) {
      if ((rows[i][map['key']] || '') === k) return rows[i][map['value']] || '';
    }
    return '';
  }, { entityType: 'site_config', entityId: key || '' });
}

function updateSiteMetric(payload) {
  return tryWithAudit_(function() {
    var auth = assertAccessOrThrow_(true);
    var k = (payload && payload.key || '').toString().trim();
    var v = (payload && payload.value || '').toString().trim();
    if (!k) throw new Error('key requerido');
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(SHEET_NAMES.SITE_CONFIG);
    if (!sheet) {
      sheet = getOrCreateSheetByName_(SHEET_NAMES.SITE_CONFIG);
      ensureHeaders_(sheet, SHEET_HEADERS.SITE_CONFIG);
    }
    var hdr = SHEET_HEADERS.SITE_CONFIG;
    var map = getHeaderIndexMap_(sheet, hdr);
    var rows = getRows_(sheet);
    var found = -1;
    for (var i = 0; i < rows.length; i++) { if ((rows[i][map['key']] || '') === k) { found = i + 2; break; } }
    var row = [];
    row[map['key']] = k;
    row[map['value']] = v;
    row[map['updatedAt']] = toIsoString_(now_());
    row[map['updatedBy']] = auth.email;
    if (found > 0) {
      sheet.getRange(found, 1, 1, hdr.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    logAudit_('site_config', k, 'UPDATE', auth.email, { value: v });
    return { key: k, value: v };
  }, { entityType: 'site_config', entityId: (payload && payload.key) || '' });
}



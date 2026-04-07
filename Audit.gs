function logAudit_(entityType, entityId, action, actor, metadata) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.AUDIT_LOG);
  var row = [uuid_(), entityType, entityId, action, actor, toIsoString_(now_()), JSON.stringify(metadata || {})];
  sheet.appendRow(row);
}

function appendVersion_(articleRowObject) {
  if (!SHEET_NAMES || !SHEET_NAMES.ARTICLE_VERSIONS) {
    return;
  }
  var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLE_VERSIONS);
  if (!sheet) return;
  var versionNumber = getNextVersionNumber_(articleRowObject.articleId);
  var row = [
    uuid_(),
    articleRowObject.articleId,
    versionNumber,
    articleRowObject.title,
    articleRowObject.descriptionHtmlSanitized,
    articleRowObject.driveFileId,
    articleRowObject.driveMimeType,
    toIsoString_(now_()),
    articleRowObject.updatedBy || articleRowObject.createdBy,
    articleRowObject.changeType || 'update',
    articleRowObject.changesSummary || '',
    articleRowObject.status || 'PUBLISHED'
  ];
  sheet.appendRow(row);
}

function getNextVersionNumber_(articleId) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLE_VERSIONS);
  var rows = getRows_(sheet);
  var hdrs = SHEET_HEADERS.ARTICLE_VERSIONS;
  var map = getHeaderIndexMap_(sheet, hdrs);
  var articleIdIdx = map['articleId'];
  var versionNumberIdx = map['versionNumber'];
  var max = 0;
  for (var i = 0; i < rows.length; i++) {
    if ((rows[i][articleIdIdx] || '') === articleId) {
      var v = parseInt(rows[i][versionNumberIdx], 10) || 0;
      if (v > max) max = v;
    }
  }
  return max + 1;
}



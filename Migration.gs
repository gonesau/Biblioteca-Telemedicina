function createMissingSheetsOnly() {
  var ss = getSpreadsheet_();
  var requiredSheets = [
    SHEET_NAMES.ARTICLE_TYPES,
    SHEET_NAMES.TAGS,
    SHEET_NAMES.ARTICLE_TAGS
  ];
  
  for (var i = 0; i < requiredSheets.length; i++) {
    var name = requiredSheets[i];
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    // Asegurar encabezados para estas nuevas hojas
    var keyName = '';
    for (var k in SHEET_NAMES) {
      if (SHEET_NAMES[k] === name) {
        keyName = k;
        break;
      }
    }
    var headers = SHEET_HEADERS[keyName];
    ensureHeaders_(sheet, headers);
  }
}

function migrate_v1_types_tags() {
  assertAccessOrThrow_(true);
  var ss = getSpreadsheet_();
  var articles = ss.getSheetByName(SHEET_NAMES.ARTICLES);
  var types = ss.getSheetByName(SHEET_NAMES.ARTICLE_TYPES);
  var tags = ss.getSheetByName(SHEET_NAMES.TAGS);
  var articleTags = ss.getSheetByName(SHEET_NAMES.ARTICLE_TAGS);
  var aHdr = SHEET_HEADERS.ARTICLES;
  var aMap = getHeaderIndexMap_(articles, aHdr);

  // Índices de lookup para tipos y tags existentes
  var typeMapByName = {}; // name(lower) -> typeId
  var tRows = getRows_(types);
  var tHdr = SHEET_HEADERS.ARTICLE_TYPES;
  var tMap = getHeaderIndexMap_(types, tHdr);
  for (var i = 0; i < tRows.length; i++) {
    var tr = tRows[i];
    var name = (tr[tMap['name']] || '').toString().trim();
    var tid = tr[tMap['typeId']];
    if (name && tid) typeMapByName[name.toLowerCase()] = tid;
  }

  var tagIdBySlug = {}; // slug -> tagId
  var tgRows = getRows_(tags);
  var tgHdr = SHEET_HEADERS.TAGS;
  var tgMap = getHeaderIndexMap_(tags, tgHdr);
  for (var j = 0; j < tgRows.length; j++) {
    var gr = tgRows[j];
    var slug = (gr[tgMap['slug']] || '').toString();
    var gid = gr[tgMap['tagId']];
    if (slug && gid) tagIdBySlug[slug] = gid;
  }

  var aRows = getRows_(articles);
  for (var r = 0; r < aRows.length; r++) {
    var row = aRows[r];
    var articleId = row[aMap['articleId']];
    var title = row[aMap['title']];
    var desc = row[aMap['descriptionHtmlSanitized']] || '';
    var mimeAsTypeName = (row[aMap['driveMimeType']] || 'General').toString().trim();

    // Asignar typeId basado en mimeAsTypeName
    var typeId = row[aMap['typeId']];
    if (!typeId) {
      var key = mimeAsTypeName.toLowerCase();
      if (!typeMapByName[key]) {
        var newTypeId = uuid_();
        var name = mimeAsTypeName || 'General';
        var color = '#94a3b8';
        var order = 99;
        types.appendRow([newTypeId, name, slugify_(name), color, order]);
        typeMapByName[key] = newTypeId;
      }
      typeId = typeMapByName[key];
      // Escribir typeId a la fila
      row[aMap['typeId']] = typeId;
      articles.getRange(r + 2, 1, 1, aHdr.length).setValues([row]);
    }

    // Extraer tags tipo #palabra del HTML plano
    var plain = desc.toString().replace(/<[^>]+>/g, ' ').toLowerCase();
    var found = plain.match(/#([a-z0-9áéíóúñ_-]{2,})/gi) || [];
    var slugs = {};
    for (var k = 0; k < found.length; k++) {
      var t = found[k].replace('#', '').trim();
      if (!t) continue;
      var slug = slugify_(t);
      if (slug) slugs[slug] = t;
    }
    var slugList = Object.keys(slugs);
    for (var s = 0; s < slugList.length; s++) {
      var slug = slugList[s];
      var tagId = tagIdBySlug[slug];
      if (!tagId) {
        tagId = uuid_();
        tags.appendRow([tagId, slugs[slug], slug]);
        tagIdBySlug[slug] = tagId;
      }
      // Vincular si no existe
      var exists = false;
      var atRows = getRows_(articleTags);
      var atHdr = SHEET_HEADERS.ARTICLE_TAGS;
      var atMap = getHeaderIndexMap_(articleTags, atHdr);
      for (var x = 0; x < atRows.length; x++) {
        var ar = atRows[x];
        if (ar[atMap['articleId']] === articleId && ar[atMap['tagId']] === tagId) { exists = true; break; }
      }
      if (!exists) articleTags.appendRow([articleId, tagId]);
    }
  }

  logAudit_('migration', 'v1', 'TYPES_TAGS_DONE', getActiveUserEmail_(), {});
  return 'OK';
}

function migrate_v1_drive_restrictions(batchSize) {
  assertAccessOrThrow_(true);
  var size = (batchSize && batchSize > 0 ? batchSize : 50);
  var ss = getSpreadsheet_();
  var articles = ss.getSheetByName(SHEET_NAMES.ARTICLES);
  var aHdr = SHEET_HEADERS.ARTICLES;
  var aMap = getHeaderIndexMap_(articles, aHdr);
  var rows = getRows_(articles);

  var props = PropertiesService.getScriptProperties();
  var cursor = parseInt(props.getProperty('MIG_V1_DRIVE_CURSOR') || '0', 10) || 0;
  var end = Math.min(rows.length, cursor + size);
  for (var i = cursor; i < end; i++) {
    var r = rows[i];
    var article = {
      articleId: r[aMap['articleId']],
      driveFileId: r[aMap['driveFileId']]
    };
    try { ensureRestrictionsForArticle_(article); } catch (e) {}
  }
  var next = end >= rows.length ? 0 : end;
  props.setProperty('MIG_V1_DRIVE_CURSOR', String(next));
  logAudit_('migration', 'v1', 'DRIVE_RESTRICTIONS_BATCH', getActiveUserEmail_(), { from: cursor, to: end, next: next });
  return { processed: end - cursor, nextCursor: next, total: rows.length };
}



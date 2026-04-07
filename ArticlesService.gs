function assertAccessOrThrow_(requireAdmin) {
  var email = getActiveUserEmail_();
  var cfg = getAccessConfig();
  var isAdmin = cfg.adminEmails && cfg.adminEmails.indexOf(email) !== -1;
  var isGuestEditor = cfg.guestEditorEmails && cfg.guestEditorEmails.indexOf(email) !== -1;
  var allowed = isAdmin || isGuestEditor || isEmailAllowed_(email);
  if (!allowed) {
    logAudit_('article', '', 'READ', email, { denied: true, reason: 'not allowed' });
    throw new Error('Acceso denegado');
  }
  if (requireAdmin && !isAdmin) {
    logAudit_('article', '', 'READ', email, { denied: true, reason: 'admin required' });
    throw new Error('Permisos insuficientes');
  }
  return { email: email, isAdmin: isAdmin, isGuestEditor: isGuestEditor };
}

function readArticles() {
  return tryWithAudit_(function() {
    var auth = assertAccessOrThrow_(false);
    var canSeeReview = auth.isAdmin || auth.isGuestEditor;
    
    // Try cache first
    var cacheKey = CACHE_KEY_ARTICLES + (canSeeReview ? '_all' : '_pub');
    var cached = getCachedJSON_(cacheKey);
    if (cached) return cached;
    
    var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLES);
    var hdrs = SHEET_HEADERS.ARTICLES;
    var map = getHeaderIndexMap_(sheet, hdrs);
    var rows = getRows_(sheet);
    var typesSheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLE_TYPES);
    var typesMap = {};
    if (typesSheet) {
      var th = SHEET_HEADERS.ARTICLE_TYPES;
      var tm = getHeaderIndexMap_(typesSheet, th);
      var trows = getRows_(typesSheet);
      for (var ti = 0; ti < trows.length; ti++) {
        var tr = trows[ti];
        typesMap[tr[tm['typeId']]] = {
          typeId: tr[tm['typeId']],
          name: tr[tm['name']],
          slug: tr[tm['slug']],
          color: tr[tm['color']],
          order: tr[tm['order']]
        };
      }
    }
    var tagsByArticle = {};
    var tagsSheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.TAGS);
    var atSheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLE_TAGS);
    if (tagsSheet && atSheet) {
      var tgHdr = SHEET_HEADERS.TAGS;
      var tgMap = getHeaderIndexMap_(tagsSheet, tgHdr);
      var tgRows = getRows_(tagsSheet);
      var tagById = {};
      for (var gj = 0; gj < tgRows.length; gj++) {
        var gr = tgRows[gj];
        tagById[gr[tgMap['tagId']]] = { tagId: gr[tgMap['tagId']], name: gr[tgMap['name']], slug: gr[tgMap['slug']] };
      }
      var atHdr = SHEET_HEADERS.ARTICLE_TAGS;
      var atMap = getHeaderIndexMap_(atSheet, atHdr);
      var atRows = getRows_(atSheet);
      for (var ai = 0; ai < atRows.length; ai++) {
        var ar = atRows[ai];
        var aid = ar[atMap['articleId']];
        var tid = ar[atMap['tagId']];
        if (!tagsByArticle[aid]) tagsByArticle[aid] = [];
        if (tagById[tid]) tagsByArticle[aid].push(tagById[tid]);
      }
    }
    var list = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var status = r[map['status']] || 'PUBLISHED';
      if (!canSeeReview && status === 'NEEDS_REVIEW') continue;

      var typeObj = typesMap[r[map['typeId']]] || null;
      var descHtml = r[map['descriptionHtmlSanitized']] || '';
      var descPlain = descHtml.replace(/<[^>]+>/g, '').trim();
      var descShort = descPlain.length > 200 ? descPlain.substring(0, 200) + '...' : descPlain;
      list.push({
        articleId: r[map['articleId']],
        title: r[map['title']],
        slug: r[map['slug']],
        typeId: r[map['typeId']],
        typeName: typeObj ? typeObj.name : '',
        typeColor: typeObj ? typeObj.color : '',
        descriptionShort: descShort,
        driveFileId: r[map['driveFileId']],
        driveMimeType: r[map['driveMimeType']],
        createdAt: r[map['createdAt']],
        updatedAt: r[map['updatedAt']],
        createdBy: r[map['createdBy']],
        updatedBy: r[map['updatedBy']],
        status: status,
        tags: tagsByArticle[r[map['articleId']]] || []
      });
    }
    setCachedJSON_(cacheKey, list, CACHE_TTL_SECONDS);
    return list;
  }, { entityType: 'article', entityId: '' });
}

function readArticleBySlug(slug) {
  return tryWithAudit_(function() {
    var auth = assertAccessOrThrow_(false);
    var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLES);
    var hdrs = SHEET_HEADERS.ARTICLES;
    var map = getHeaderIndexMap_(sheet, hdrs);
    var rows = getRows_(sheet);
    // Prepara lookups de types y tags
    var typesSheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLE_TYPES);
    var typesMap = {};
    if (typesSheet) {
      var th = SHEET_HEADERS.ARTICLE_TYPES;
      var tm = getHeaderIndexMap_(typesSheet, th);
      var trows = getRows_(typesSheet);
      for (var ti = 0; ti < trows.length; ti++) {
        var tr = trows[ti];
        typesMap[tr[tm['typeId']]] = {
          typeId: tr[tm['typeId']], name: tr[tm['name']], slug: tr[tm['slug']], color: tr[tm['color']], order: tr[tm['order']]
        };
      }
    }
    var tagsSheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.TAGS);
    var atSheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLE_TAGS);
    var tagsForArticle = function(articleId) {
      var out = [];
      if (!tagsSheet || !atSheet) return out;
      var tgHdr = SHEET_HEADERS.TAGS;
      var tgMap = getHeaderIndexMap_(tagsSheet, tgHdr);
      var atHdr = SHEET_HEADERS.ARTICLE_TAGS;
      var atMap = getHeaderIndexMap_(atSheet, atHdr);
      var tgRows = getRows_(tagsSheet);
      var atRows = getRows_(atSheet);
      var tagById = {};
      for (var gj = 0; gj < tgRows.length; gj++) tagById[tgRows[gj][tgMap['tagId']]] = { tagId: tgRows[gj][tgMap['tagId']], name: tgRows[gj][tgMap['name']], slug: tgRows[gj][tgMap['slug']] };
      for (var ai = 0; ai < atRows.length; ai++) {
        if (atRows[ai][atMap['articleId']] === articleId) {
          var tid = atRows[ai][atMap['tagId']];
          if (tagById[tid]) out.push(tagById[tid]);
        }
      }
      return out;
    };
    var slugIdx = map['slug'];
    var canSeeReview = auth.isAdmin || auth.isGuestEditor;
    for (var i = 0; i < rows.length; i++) {
      if ((rows[i][slugIdx] || '') === slug) {
        var r = rows[i];
        var status = r[map['status']] || 'PUBLISHED';
        if (!canSeeReview && status === 'NEEDS_REVIEW') return null;

        var typeObj = typesMap[r[map['typeId']]] || null;
        return {
          articleId: r[map['articleId']],
          title: r[map['title']],
          slug: r[map['slug']],
          typeId: r[map['typeId']],
          typeName: typeObj ? typeObj.name : '',
          typeColor: typeObj ? typeObj.color : '',
          descriptionHtmlSanitized: r[map['descriptionHtmlSanitized']],
          driveFileId: r[map['driveFileId']],
          driveMimeType: r[map['driveMimeType']],
          createdAt: r[map['createdAt']],
          updatedAt: r[map['updatedAt']],
          createdBy: r[map['createdBy']],
          updatedBy: r[map['updatedBy']],
          status: status,
          tags: tagsForArticle(r[map['articleId']])
        };
      }
    }
    return null;
  }, { entityType: 'article', entityId: slug });
}

/**
 * Artículo completo para edición (incluye descriptionHtmlSanitized).
 * El listado readArticles() solo envía descriptionShort para aligerar el payload.
 */
function readArticleByArticleId(articleId) {
  return tryWithAudit_(function() {
    var auth = assertAccessOrThrow_(false);
    var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLES);
    var hdrs = SHEET_HEADERS.ARTICLES;
    var map = getHeaderIndexMap_(sheet, hdrs);
    var rows = getRows_(sheet);
    var typesSheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLE_TYPES);
    var typesMap = {};
    if (typesSheet) {
      var th = SHEET_HEADERS.ARTICLE_TYPES;
      var tm = getHeaderIndexMap_(typesSheet, th);
      var trows = getRows_(typesSheet);
      for (var ti = 0; ti < trows.length; ti++) {
        var tr = trows[ti];
        typesMap[tr[tm['typeId']]] = {
          typeId: tr[tm['typeId']], name: tr[tm['name']], slug: tr[tm['slug']], color: tr[tm['color']], order: tr[tm['order']]
        };
      }
    }
    var tagsSheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.TAGS);
    var atSheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLE_TAGS);
    var tagsForArticle = function(aid) {
      var out = [];
      if (!tagsSheet || !atSheet) return out;
      var tgHdr = SHEET_HEADERS.TAGS;
      var tgMap = getHeaderIndexMap_(tagsSheet, tgHdr);
      var atHdr = SHEET_HEADERS.ARTICLE_TAGS;
      var atMap = getHeaderIndexMap_(atSheet, atHdr);
      var tgRows = getRows_(tagsSheet);
      var atRows = getRows_(atSheet);
      var tagById = {};
      for (var gj = 0; gj < tgRows.length; gj++) tagById[tgRows[gj][tgMap['tagId']]] = { tagId: tgRows[gj][tgMap['tagId']], name: tgRows[gj][tgMap['name']], slug: tgRows[gj][tgMap['slug']] };
      for (var ai = 0; ai < atRows.length; ai++) {
        if (atRows[ai][atMap['articleId']] === aid) {
          var tid = atRows[ai][atMap['tagId']];
          if (tagById[tid]) out.push(tagById[tid]);
        }
      }
      return out;
    };
    var idIdx = map['articleId'];
    var wantId = (articleId || '').toString();
    var canSeeReview = auth.isAdmin || auth.isGuestEditor;
    for (var i = 0; i < rows.length; i++) {
      if ((rows[i][idIdx] || '').toString() === wantId) {
        var r = rows[i];
        var status = r[map['status']] || 'PUBLISHED';
        if (!canSeeReview && status === 'NEEDS_REVIEW') return null;
        var typeObj = typesMap[r[map['typeId']]] || null;
        return {
          articleId: r[map['articleId']],
          title: r[map['title']],
          slug: r[map['slug']],
          typeId: r[map['typeId']],
          typeName: typeObj ? typeObj.name : '',
          typeColor: typeObj ? typeObj.color : '',
          descriptionHtmlSanitized: r[map['descriptionHtmlSanitized']],
          driveFileId: r[map['driveFileId']],
          driveMimeType: r[map['driveMimeType']],
          createdAt: r[map['createdAt']],
          updatedAt: r[map['updatedAt']],
          createdBy: r[map['createdBy']],
          updatedBy: r[map['updatedBy']],
          status: status,
          tags: tagsForArticle(r[map['articleId']])
        };
      }
    }
    return null;
  }, { entityType: 'article', entityId: articleId });
}

function createArticle(payload) {
  return tryWithAudit_(function() {
    var auth = assertAccessOrThrow_(true);
    var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLES);
    var hdrs = SHEET_HEADERS.ARTICLES;
    var map = getHeaderIndexMap_(sheet, hdrs);

    var title = (payload.title || '').toString().trim();
    var description = sanitizeHtml_((payload.descriptionHtml || '').toString());
    var driveFileId = extractDriveFileIdFromAny_((payload.driveFileId || '').toString().trim());
    var driveMimeType = (payload.driveMimeType || '').toString().trim();
    var typeId = (payload.typeId || '').toString().trim();
    var status = (payload.status || 'PUBLISHED').toString().trim();
    if (!typeId && driveMimeType) {
      try { typeId = getOrCreateTypeIdByName_(driveMimeType); } catch (_) {}
    }
    if (!title) throw new Error('Título requerido');

    var slug = ensureUniqueSlug_(slugify_(title));
    var articleId = uuid_();
    var ts = toIsoString_(now_());
    var row = [];
    row[map['articleId']] = articleId;
    row[map['title']] = title;
    row[map['slug']] = slug;
    row[map['typeId']] = typeId;
    row[map['descriptionHtmlSanitized']] = description;
    row[map['driveFileId']] = driveFileId;
    row[map['driveMimeType']] = driveMimeType;
    row[map['createdAt']] = ts;
    row[map['updatedAt']] = ts;
    row[map['createdBy']] = auth.email;
    row[map['updatedBy']] = auth.email;
    row[map['status']] = status;

    sheet.appendRow(row);

    var articleObj = {
      articleId: articleId,
      title: title,
      slug: slug,
      typeId: typeId,
      descriptionHtmlSanitized: description,
      driveFileId: driveFileId,
      driveMimeType: driveMimeType,
      createdBy: auth.email,
      updatedBy: auth.email,
      status: status,
      changeType: 'create',
      changesSummary: 'Artículo creado'
    };

    appendVersion_(articleObj);
    logAudit_('article', articleId, 'CREATE', auth.email, { title: title, slug: slug });
    // Aplicar restricciones de Drive si hay archivo
    try { ensureRestrictionsForArticle_(articleObj); } catch (ignored) {}
    // Guardar tags (tagsText separadas por comas)
    try { upsertArticleTagsForArticle_(articleId, (payload && payload.tagsText) || ''); } catch (e) { logAudit_('tag', articleId, 'ERROR', auth.email, { message: e && e.message }); }
    if (status !== 'NEEDS_REVIEW') {
      try { notifyChange_('create', articleObj); } catch (mailErr) { logAudit_('article', articleId, 'ERROR', auth.email, { notificationFailed: true, message: mailErr && mailErr.message }); }
    }
    invalidateArticlesCache_();
    return { articleId: articleId, slug: slug };
  }, { entityType: 'article', entityId: '' });
}

function updateArticle(payload) {
  return tryWithAudit_(function() {
    var auth = assertAccessOrThrow_(true);
    var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLES);
    var hdrs = SHEET_HEADERS.ARTICLES;
    var map = getHeaderIndexMap_(sheet, hdrs);
    var id = (payload.articleId || '').toString().trim();
    if (!id) throw new Error('articleId requerido');
    var rowIndex = findRowIndexByKey_(sheet, 'articleId', id, map);
    if (rowIndex < 0) throw new Error('Artículo no encontrado');

    var row = sheet.getRange(rowIndex, 1, 1, hdrs.length).getValues()[0];
    var before = {
      title: row[map['title']],
      descriptionHtmlSanitized: row[map['descriptionHtmlSanitized']],
      driveFileId: row[map['driveFileId']],
      driveMimeType: row[map['driveMimeType']],
      typeId: row[map['typeId']],
      status: row[map['status']]
    };

    var title = (payload.title != null ? payload.title : before.title).toString().trim();
    var description = (payload.descriptionHtml != null ? sanitizeHtml_(payload.descriptionHtml) : before.descriptionHtmlSanitized);
    var driveFileId = (payload.driveFileId != null ? extractDriveFileIdFromAny_(payload.driveFileId) : before.driveFileId).toString().trim();
    var driveMimeType = (payload.driveMimeType != null ? payload.driveMimeType : before.driveMimeType).toString().trim();
    var typeId = (payload.typeId != null ? payload.typeId : before.typeId).toString().trim();
    var status = (payload.status != null ? payload.status : (before.status || 'PUBLISHED')).toString().trim();
    // si cambiaron la categoría (driveMimeType) recalcular/asegurar typeId
    if ((payload.driveMimeType != null) && driveMimeType) {
      try { typeId = getOrCreateTypeIdByName_(driveMimeType); } catch (_) {}
    }

    row[map['title']] = title;
    // Slug inmutable por defecto; si el título cambió, mantenemos slug.
    row[map['typeId']] = typeId;
    row[map['descriptionHtmlSanitized']] = description;
    row[map['driveFileId']] = driveFileId;
    row[map['driveMimeType']] = driveMimeType;
    row[map['status']] = status;
    row[map['updatedAt']] = toIsoString_(now_());
    row[map['updatedBy']] = auth.email;
    sheet.getRange(rowIndex, 1, 1, hdrs.length).setValues([row]);

    var articleObj = {
      articleId: id,
      title: title,
      slug: row[map['slug']],
      typeId: typeId,
      descriptionHtmlSanitized: description,
      driveFileId: driveFileId,
      driveMimeType: driveMimeType,
      updatedBy: auth.email,
      status: status,
      changeType: 'update',
      changesSummary: 'Artículo actualizado'
    };

    appendVersion_(articleObj);
    logAudit_('article', id, 'UPDATE', auth.email, { before: before, after: { title: title } });
    // Reaplicar restricciones de Drive en caso de cambio de archivo o por seguridad
    try { ensureRestrictionsForArticle_(articleObj); } catch (ignored) {}
    // Actualizar tags del artículo
    try { upsertArticleTagsForArticle_(id, (payload && payload.tagsText) || ''); } catch (e) { logAudit_('tag', id, 'ERROR', auth.email, { message: e && e.message }); }
    if (status !== 'NEEDS_REVIEW') {
      try { notifyChange_('update', articleObj); } catch (mailErr) { logAudit_('article', id, 'ERROR', auth.email, { notificationFailed: true, message: mailErr && mailErr.message }); }
    }
    invalidateArticlesCache_();
    return { articleId: id, slug: row[map['slug']] };
  }, { entityType: 'article', entityId: (payload && payload.articleId) || '' });
}

// --- Comments helpers ---
function getArticleComments(articleId) {
  return tryWithAudit_(function() {
    var auth = assertAccessOrThrow_(false);
    var canSeeReview = auth.isAdmin || auth.isGuestEditor;
    if (!canSeeReview) return []; // Only admins and guest editors can see comments

    var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLE_COMMENTS);
    if (!sheet) return [];
    var hdrs = SHEET_HEADERS.ARTICLE_COMMENTS;
    var map = getHeaderIndexMap_(sheet, hdrs);
    var rows = getRows_(sheet);
    var comments = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r[map['articleId']] === articleId) {
        comments.push({
          commentId: r[map['commentId']],
          articleId: r[map['articleId']],
          text: r[map['text']],
          createdAt: r[map['createdAt']],
          createdBy: r[map['createdBy']]
        });
      }
    }
    // Sort by createdAt desc
    comments.sort(function(a, b) {
      return new Date(b.createdAt) - new Date(a.createdAt);
    });
    return comments;
  }, { entityType: 'comment', entityId: articleId });
}

function addArticleComment(payload) {
  return tryWithAudit_(function() {
    var auth = assertAccessOrThrow_(false);
    var canSeeReview = auth.isAdmin || auth.isGuestEditor;
    if (!canSeeReview) throw new Error('Permisos insuficientes para comentar');

    var articleId = (payload.articleId || '').toString().trim();
    var text = (payload.text || '').toString().trim();
    if (!articleId || !text) throw new Error('articleId y text requeridos');

    var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLE_COMMENTS);
    var hdrs = SHEET_HEADERS.ARTICLE_COMMENTS;
    var map = getHeaderIndexMap_(sheet, hdrs);

    var commentId = uuid_();
    var ts = toIsoString_(now_());
    var row = [];
    row[map['commentId']] = commentId;
    row[map['articleId']] = articleId;
    row[map['text']] = text;
    row[map['createdAt']] = ts;
    row[map['createdBy']] = auth.email;

    sheet.appendRow(row);
    logAudit_('comment', commentId, 'CREATE', auth.email, { articleId: articleId });

    // Enviar notificación
    try { notifyComment_(articleId, text, auth.email); } catch (mailErr) { logAudit_('comment', commentId, 'ERROR', auth.email, { notificationFailed: true, message: mailErr && mailErr.message }); }

    return { commentId: commentId, articleId: articleId, text: text, createdAt: ts, createdBy: auth.email };
  }, { entityType: 'comment', entityId: '' });
}

function notifyComment_(articleId, text, author) {
  var cfg = getNotificationsConfig();
  if (!cfg.recipients.length) return;
  
  // Look up article by ID directly from the sheet
  var article = null;
  var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLES);
  var hdrs = SHEET_HEADERS.ARTICLES;
  var map = getHeaderIndexMap_(sheet, hdrs);
  var rows = getRows_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][map['articleId']] === articleId) {
      article = { title: rows[i][map['title']], slug: rows[i][map['slug']] };
      break;
    }
  }

  var base = getBasePublicUrl_();
  var url = base ? (base + '#/articulo/' + (article ? article.slug : articleId)) : ('#/articulo/' + (article ? article.slug : articleId));
  var title = article ? article.title : articleId;

  var subject = 'Biblioteca Telemedicina: Nuevo comentario en revisión - ' + title;
  
  var actionColor = '#F59E0B'; // Orange for review comments
  
  var htmlBody = '<!DOCTYPE html>' +
    '<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">' +
    '<style>body{margin:0;padding:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;background-color:#f5f7fa;}' +
    '.container{max-width:600px;margin:0 auto;background-color:#ffffff;}' +
    '.header{background:linear-gradient(135deg,#04013C 0%,#032D6B 100%);padding:32px 24px;text-align:center;}' +
    '.logo{width:60px;height:60px;background:url("https://i.pinimg.com/originals/57/9b/90/579b90f5e64d7631fc81e55ba716ba9f.png") center/contain no-repeat;border-radius:12px;margin:0 auto 16px;}' +
    '.header-title{color:#ffffff;font-size:24px;font-weight:700;margin:0;}' +
    '.badge{display:inline-block;background:' + actionColor + ';color:#fff;padding:8px 16px;border-radius:20px;font-size:14px;font-weight:600;margin-top:12px;text-transform:uppercase;}' +
    '.content{padding:32px 24px;}' +
    '.article-title{color:#04013C;font-size:22px;font-weight:700;margin:0 0 12px;line-height:1.3;}' +
    '.meta{color:#6b7280;font-size:14px;margin-bottom:20px;display:flex;flex-wrap:wrap;gap:16px;}' +
    '.meta-item{display:inline-flex;align-items:center;gap:6px;}' +
    '.meta-label{font-weight:600;}' +
    '.description{color:#384254;font-size:15px;line-height:1.6;margin-bottom:24px;padding:16px;background:#f8fafc;border-left:4px solid ' + actionColor + ';border-radius:4px;white-space:pre-wrap;}' +
    '.cta{text-align:center;margin:24px 0;}' +
    '.btn{display:inline-block;background:' + actionColor + ';color:#ffffff;text-decoration:none;padding:14px 32px;border-radius:8px;font-weight:600;font-size:16px;}' +
    '.btn:hover{opacity:0.9;}' +
    '.footer{background:#f8fafc;padding:24px;text-align:center;color:#6b7280;font-size:13px;border-top:1px solid #e5e7eb;}' +
    '.footer-text{margin:0 0 8px;}' +
    '@media only screen and (max-width:600px){.container{width:100%!important;}.content{padding:24px 16px!important;}.header{padding:24px 16px!important;}}</style>' +
    '</head><body>' +
    '<div class="container">' +
    '<div class="header">' +
    '<div class="logo"></div>' +
    '<h1 class="header-title">Biblioteca Telemedicina</h1>' +
    '<div class="badge">Nuevo Comentario de Revisión</div>' +
    '</div>' +
    '<div class="content">' +
    '<h2 class="article-title">' + title + '</h2>' +
    '<div class="meta">' +
    '<span class="meta-item"><span class="meta-label">Por:</span> ' + author + '</span>' +
    '</div>' +
    '<div class="description">' + text + '</div>' +
    '<div class="cta">' +
    '<a href="' + url + '" class="btn">Ver Artículo y Comentarios</a>' +
    '</div>' +
    '</div>' +
    '<div class="footer">' +
    '<p class="footer-text">Este es un mensaje automático de la Biblioteca de Telemedicina.</p>' +
    '</div>' +
    '</div>' +
    '</body></html>';

  var plainBody = 'Nuevo comentario de ' + author + ' en el artículo "' + title + '":\n\n' +
                  text + '\n\n' +
                  'Acceso: ' + url;
  
  var options = {
    htmlBody: htmlBody
  };
  if (cfg.replyTo) options.replyTo = cfg.replyTo;
  if (cfg.senderAlias) options.name = cfg.senderAlias;
  
  for (var i = 0; i < cfg.recipients.length; i++) {
    MailApp.sendEmail(cfg.recipients[i], subject, plainBody, options);
  }
}

// --- Version history ---
function getArticleVersions(articleId) {
  return tryWithAudit_(function() {
    var auth = assertAccessOrThrow_(true);
    if (!articleId) return [];
    var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLE_VERSIONS);
    if (!sheet) return [];
    var hdrs = SHEET_HEADERS.ARTICLE_VERSIONS;
    var map = getHeaderIndexMap_(sheet, hdrs);
    var rows = getRows_(sheet);
    var versions = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r[map['articleId']] === articleId) {
        versions.push({
          versionId: r[map['versionId']],
          versionNumber: r[map['versionNumber']],
          title: r[map['title']],
          changedAt: r[map['changedAt']],
          changedBy: r[map['changedBy']],
          changeType: r[map['changeType']],
          changesSummary: r[map['changesSummary']],
          status: r[map['status']]
        });
      }
    }
    versions.sort(function(a, b) {
      return (parseInt(b.versionNumber, 10) || 0) - (parseInt(a.versionNumber, 10) || 0);
    });
    return versions;
  }, { entityType: 'version', entityId: articleId });
}

// --- Tags helpers ---
function upsertArticleTagsForArticle_(articleId, tagsInput) {
  try {
    var ss = getSpreadsheet_();
    var tagsSheet = ss.getSheetByName(SHEET_NAMES.TAGS);
    var atSheet = ss.getSheetByName(SHEET_NAMES.ARTICLE_TAGS);
    if (!tagsSheet || !atSheet) return;
    var tgHdr = SHEET_HEADERS.TAGS, tgMap = getHeaderIndexMap_(tagsSheet, tgHdr);
    var atHdr = SHEET_HEADERS.ARTICLE_TAGS, atMap = getHeaderIndexMap_(atSheet, atHdr);

    // Normalizar entrada → lista única de nombres
    var list = Array.isArray(tagsInput) ? tagsInput : String(tagsInput || '').split(',');
    var names = [];
    var seen = {};
    for (var i = 0; i < list.length; i++) {
      var n = (list[i] || '').toString().trim();
      if (!n) continue;
      var key = n.toLowerCase();
      if (!seen[key]) { seen[key] = true; names.push(n); }
    }

    // Borrar relaciones actuales del artículo
    var atRows = getRows_(atSheet);
    for (var r = atRows.length - 1; r >= 0; r--) {
      if (atRows[r][atMap['articleId']] === articleId) {
        atSheet.deleteRow(r + 2);
      }
    }
    if (!names.length) return;

    // Indexar tags existentes por slug
    var tgRows = getRows_(tagsSheet);
    var bySlug = {};
    for (var t = 0; t < tgRows.length; t++) {
      var slug = (tgRows[t][tgMap['slug']] || '').toString();
      if (slug) bySlug[slug] = tgRows[t][tgMap['tagId']];
    }

    // Asegurar tags y crear relaciones
    for (var j = 0; j < names.length; j++) {
      var name = names[j];
      var slugName = slugify_(name);
      var tagId = bySlug[slugName];
      if (!tagId) {
        tagId = uuid_();
        tagsSheet.appendRow([tagId, name, slugName]);
        bySlug[slugName] = tagId;
      }
      atSheet.appendRow([articleId, tagId]);
    }
  } catch (e) {
    throw e;
  }
}

// Asegura y retorna un typeId para un nombre de tipo (case-insensitive). Crea el tipo si no existe.
function getOrCreateTypeIdByName_(name) {
  var typeName = (name || '').toString().trim();
  if (!typeName) return '';
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_NAMES.ARTICLE_TYPES);
  if (!sheet) return '';
  var hdr = SHEET_HEADERS.ARTICLE_TYPES;
  var map = getHeaderIndexMap_(sheet, hdr);
  var rows = getRows_(sheet);
  var lower = typeName.toLowerCase();
  for (var i = 0; i < rows.length; i++) {
    var nm = (rows[i][map['name']] || '').toString().trim().toLowerCase();
    if (nm === lower) return rows[i][map['typeId']];
  }
  var id = uuid_();
  var slug = slugify_(typeName);
  var color = '#94a3b8';
  var order = 99;
  var row = [];
  row[map['typeId']] = id;
  row[map['name']] = typeName;
  row[map['slug']] = slug;
  row[map['color']] = color;
  row[map['order']] = order;
  sheet.appendRow(row);
  return id;
}

function listArticleTypes() {
  return tryWithAudit_(function() {
    assertAccessOrThrow_(false);
    var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLE_TYPES);
    if (!sheet) return [];
    var hdr = SHEET_HEADERS.ARTICLE_TYPES;
    var map = getHeaderIndexMap_(sheet, hdr);
    var rows = getRows_(sheet);
    var out = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      out.push({ typeId: r[map['typeId']], name: r[map['name']], slug: r[map['slug']], color: r[map['color']], order: r[map['order']] });
    }
    // Orden por order asc, luego name
    out.sort(function(a,b){ var ao = parseInt(a.order,10)||0; var bo = parseInt(b.order,10)||0; if (ao!==bo) return ao-bo; return (a.name||'').localeCompare(b.name||''); });
    return out;
  }, { entityType: 'type', entityId: '' });
}

function listTags() {
  return tryWithAudit_(function() {
    assertAccessOrThrow_(false);
    var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.TAGS);
    if (!sheet) return [];
    var hdr = SHEET_HEADERS.TAGS;
    var map = getHeaderIndexMap_(sheet, hdr);
    var rows = getRows_(sheet);
    var out = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      out.push({ tagId: r[map['tagId']], name: r[map['name']], slug: r[map['slug']] });
    }
    out.sort(function(a,b){ return (a.name||'').localeCompare(b.name||''); });
    return out;
  }, { entityType: 'tag', entityId: '' });
}

function logPreview(payload) {
  return tryWithAudit_(function() {
    var auth = assertAccessOrThrow_(false);
    var id = payload && payload.articleId || '';
    logAudit_('article', id, 'PREVIEW', auth.email, {});
    return 'OK';
  }, { entityType: 'article', entityId: (payload && payload.articleId) || '' });
}

function deleteArticle(payload) {
  return tryWithAudit_(function() {
    var auth = assertAccessOrThrow_(true);
    var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLES);
    var hdrs = SHEET_HEADERS.ARTICLES;
    var map = getHeaderIndexMap_(sheet, hdrs);
    var id = (payload.articleId || '').toString().trim();
    if (!id) throw new Error('articleId requerido');
    var rowIndex = findRowIndexByKey_(sheet, 'articleId', id, map);
    if (rowIndex < 0) throw new Error('Artículo no encontrado');

    var row = sheet.getRange(rowIndex, 1, 1, hdrs.length).getValues()[0];
    var articleObj = {
      articleId: id,
      title: row[map['title']],
      slug: row[map['slug']],
      descriptionHtmlSanitized: row[map['descriptionHtmlSanitized']],
      driveFileId: row[map['driveFileId']],
      driveMimeType: row[map['driveMimeType']],
      updatedBy: auth.email,
      changeType: 'delete',
      changesSummary: 'Artículo eliminado'
    };
    // Registrar versión antes de eliminar para histórico
    appendVersion_(articleObj);
    sheet.deleteRow(rowIndex);
    logAudit_('article', id, 'DELETE', auth.email, { title: articleObj.title });
    try { notifyChange_('delete', articleObj); } catch (mailErr) { logAudit_('article', id, 'ERROR', auth.email, { notificationFailed: true, message: mailErr && mailErr.message }); }
    invalidateArticlesCache_();
    return { articleId: id, slug: articleObj.slug };
  }, { entityType: 'article', entityId: (payload && payload.articleId) || '' });
}

function notifyChange_(action, articleObj) {
  // Enviar notificación a Google Chat (si está configurado)
  try { sendChatNotification_(action, articleObj); } catch (e) { logAudit_('chat', (articleObj && articleObj.articleId) || '', 'ERROR', getActiveUserEmail_(), { message: e && e.message ? e.message : String(e) }); }

  var cfg = getNotificationsConfig();
  if (!cfg.recipients.length) return;
  var base = getBasePublicUrl_();
  var url = base ? (base + '#/articulo/' + articleObj.slug) : ('#/articulo/' + articleObj.slug);
  
  var actionText = action === 'create' ? 'creado' : action === 'update' ? 'actualizado' : 'eliminado';
  var actionColor = action === 'create' ? '#34A853' : action === 'update' ? '#4285F4' : '#EA4335';
  var subject = 'Biblioteca Telemedicina: ' + (action === 'create' ? 'Nuevo Artículo' : action === 'update' ? 'Actualización' : 'Eliminación') + ' - ' + (articleObj.title || '');
  
  var htmlBody = buildEmailTemplate_(action, actionText, actionColor, articleObj, url);
  var plainBody = 'El artículo "' + (articleObj.title || '') + '" ha sido ' + actionText + ' por ' + (articleObj.updatedBy || articleObj.createdBy) + '\n' +
                  'Acceso: ' + url + '\n' +
                  'Fecha: ' + toIsoString_(now_());
  
  var options = {
    htmlBody: htmlBody
  };
  if (cfg.replyTo) options.replyTo = cfg.replyTo;
  if (cfg.senderAlias) options.name = cfg.senderAlias;
  
  for (var i = 0; i < cfg.recipients.length; i++) {
    MailApp.sendEmail(cfg.recipients[i], subject, plainBody, options);
  }
}

function buildEmailTemplate_(action, actionText, actionColor, article, url) {
  var descText = (article.descriptionHtmlSanitized || '').replace(/<[^>]+>/g, '').trim();
  var descShort = descText.length > 200 ? descText.slice(0, 200) + '...' : descText;
  var author = article.updatedBy || article.createdBy || 'Sistema';
  var date = toIsoString_(now_()).slice(0, 19).replace('T', ' ');
  
  return '<!DOCTYPE html>' +
    '<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">' +
    '<style>body{margin:0;padding:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;background-color:#f5f7fa;}' +
    '.container{max-width:600px;margin:0 auto;background-color:#ffffff;}' +
    '.header{background:linear-gradient(135deg,#04013C 0%,#032D6B 100%);padding:32px 24px;text-align:center;}' +
    '.logo{width:60px;height:60px;background:url("https://i.pinimg.com/originals/57/9b/90/579b90f5e64d7631fc81e55ba716ba9f.png") center/contain no-repeat;border-radius:12px;margin:0 auto 16px;}' +
    '.header-title{color:#ffffff;font-size:24px;font-weight:700;margin:0;}' +
    '.badge{display:inline-block;background:' + actionColor + ';color:#fff;padding:8px 16px;border-radius:20px;font-size:14px;font-weight:600;margin-top:12px;text-transform:uppercase;}' +
    '.content{padding:32px 24px;}' +
    '.article-title{color:#04013C;font-size:22px;font-weight:700;margin:0 0 12px;line-height:1.3;}' +
    '.meta{color:#6b7280;font-size:14px;margin-bottom:20px;display:flex;flex-wrap:wrap;gap:16px;}' +
    '.meta-item{display:inline-flex;align-items:center;gap:6px;}' +
    '.meta-label{font-weight:600;}' +
    '.description{color:#384254;font-size:15px;line-height:1.6;margin-bottom:24px;padding:16px;background:#f8fafc;border-left:4px solid ' + actionColor + ';border-radius:4px;}' +
    '.cta{text-align:center;margin:24px 0;}' +
    '.btn{display:inline-block;background:' + actionColor + ';color:#ffffff;text-decoration:none;padding:14px 32px;border-radius:8px;font-weight:600;font-size:16px;}' +
    '.btn:hover{opacity:0.9;}' +
    '.footer{background:#f8fafc;padding:24px;text-align:center;color:#6b7280;font-size:13px;border-top:1px solid #e5e7eb;}' +
    '.footer-text{margin:0 0 8px;}' +
    '@media only screen and (max-width:600px){.container{width:100%!important;}.content{padding:24px 16px!important;}.header{padding:24px 16px!important;}}</style>' +
    '</head><body>' +
    '<div class="container">' +
    '<div class="header">' +
    '<div class="logo"></div>' +
    '<h1 class="header-title">Biblioteca Telemedicina</h1>' +
    '<div class="badge">' + (action === 'create' ? 'Nuevo Artículo' : action === 'update' ? 'Artículo Actualizado' : 'Artículo Eliminado') + '</div>' +
    '</div>' +
    '<div class="content">' +
    '<h2 class="article-title">' + (article.title || 'Sin título') + '</h2>' +
    '<div class="meta">' +
    '<span class="meta-item"><span class="meta-label">Fecha:</span> ' + date + '</span>' +
    (article.driveMimeType ? '<span class="meta-item"><span class="meta-label"> Tipo:</span> ' + article.driveMimeType + '</span>' : '') +
    '</div>' +
    (descShort ? '<div class="description">' + descShort + '</div>' : '') +
    '<div class="cta">' +
    '<a href="' + url + '" class="btn">Ver Biblioteca de Telemedicina</a>' +
    '</div>' +
    '</div>' +
    '<div class="footer">' +
    '<p class="footer-text">Este es un mensaje automático de la Biblioteca de Telemedicina.</p>' +
    '<p class="footer-text">Para ver todos los artículos, accede al sistema.</p>' +
    '</div>' +
    '</div>' +
    '</body></html>';
}



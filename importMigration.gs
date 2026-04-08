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


/**
 * Creates a template spreadsheet in the user's Drive for bulk import.
 * Returns the URL so the user can download/open it.
 */
function createImportTemplate() {
  assertAccessOrThrow_(true);
  var ss = SpreadsheetApp.create('Plantilla Importación - Biblioteca Telemedicina');
  var sheet = ss.getActiveSheet();
  sheet.setName('Artículos');
  
  var headers = ['title', 'category', 'description', 'driveFileId', 'tags', 'status'];
  var headerLabels = ['Título *', 'Categoría', 'Descripción', 'Drive File ID o URL', 'Tags (separadas por coma)', 'Estado (PUBLISHED o NEEDS_REVIEW)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headerLabels]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f5f9');
  
  // Example row
  sheet.getRange(2, 1, 1, headers.length).setValues([['Guía de ejemplo', 'protocolo', 'Descripción del artículo', 'https://drive.google.com/file/d/ABC123/view', 'diabetes, tamizaje', 'PUBLISHED']]);
  sheet.getRange(2, 1, 1, headers.length).setFontColor('#6b7280').setFontStyle('italic');
  
  // Auto-resize columns
  for (var i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Add a note sheet with instructions
  var instrSheet = ss.insertSheet('Instrucciones');
  instrSheet.getRange('A1').setValue('Instrucciones para importar artículos').setFontSize(16).setFontWeight('bold');
  instrSheet.getRange('A3').setValue('1. Llena la hoja "Artículos" con los datos de cada artículo.');
  instrSheet.getRange('A4').setValue('2. El campo "Título" es obligatorio.');
  instrSheet.getRange('A5').setValue('3. "Categoría" es el tipo de documento (ej: protocolo, tamizaje, guía).');
  instrSheet.getRange('A6').setValue('4. "Drive File ID o URL" puede ser el ID o la URL completa del archivo en Google Drive.');
  instrSheet.getRange('A7').setValue('5. "Tags" son etiquetas separadas por coma.');
  instrSheet.getRange('A8').setValue('6. "Estado" puede ser PUBLISHED (publicado) o NEEDS_REVIEW (falta revisión). Si se deja vacío, será PUBLISHED.');
  instrSheet.getRange('A9').setValue('7. Elimina la fila de ejemplo antes de importar.');
  instrSheet.getRange('A10').setValue('8. Copia el ID de esta hoja de cálculo y pégalo en la wiki para importar.');
  instrSheet.autoResizeColumn(1);
  
  var url = ss.getUrl();
  var id = ss.getId();
  logAudit_('import', id, 'TEMPLATE_CREATED', getActiveUserEmail_(), { url: url });
  return { url: url, spreadsheetId: id };
}

/**
 * Imports articles from a template spreadsheet created by createImportTemplate().
 */
function importFromTemplate(spreadsheetId, skipDuplicates) {
  assertAccessOrThrow_(true);
  if (!spreadsheetId) throw new Error('Se requiere el ID de la hoja de cálculo');
  
  var source;
  try {
    source = SpreadsheetApp.openById(spreadsheetId);
  } catch (ex) {
    throw new Error('No se pudo abrir la hoja. Verifica que el ID es correcto y que tienes acceso.');
  }
  var sheet = source.getSheetByName('Artículos');
  if (!sheet) throw new Error('No se encontró la hoja "Artículos" en la hoja de cálculo.');
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('La hoja no tiene datos (solo encabezados o vacía).');
  
  var articlesData = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var title = (row[0] || '').toString().trim();
    if (!title) continue;
    articlesData.push({
      title: title,
      driveMimeType: (row[1] || '').toString().trim(),
      descriptionHtml: (row[2] || '').toString().trim(),
      driveFileId: (row[3] || '').toString().trim(),
      tagsText: (row[4] || '').toString().trim(),
      status: (row[5] || 'PUBLISHED').toString().trim()
    });
  }
  
  if (!articlesData.length) throw new Error('No se encontraron artículos válidos.');
  
  var result = importArticlesFromData(articlesData, skipDuplicates || false);
  
  // Also process tags for each imported article
  try {
    var artSheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLES);
    var artMap = getHeaderIndexMap_(artSheet, SHEET_HEADERS.ARTICLES);
    var artRows = getRows_(artSheet);
    for (var i = 0; i < articlesData.length; i++) {
      if (articlesData[i].tagsText) {
        // Find the articleId by title match
        var title = articlesData[i].title.toLowerCase();
        for (var j = artRows.length - 1; j >= 0; j--) {
          if ((artRows[j][artMap['title']] || '').toString().trim().toLowerCase() === title) {
            var aid = artRows[j][artMap['articleId']];
            try { upsertArticleTagsForArticle_(aid, articlesData[i].tagsText); } catch (tagErr) {}
            break;
          }
        }
      }
    }
  } catch (tagEx) {}
  
  invalidateArticlesCache_();
  return result;
}

// Funciones para migrar artículos desde un proyecto anterior

/**
 * Importa artículos desde una hoja de cálculo externa o desde datos en formato array
 * @param {Array} articlesData - Array de objetos con los datos de artículos: {title, descriptionHtml, driveFileId, driveMimeType, createdAt, createdBy}
 * @param {Boolean} skipDuplicates - Si es true, omite artículos con títulos duplicados
 */
function importArticlesFromData(articlesData, skipDuplicates) {
  assertAccessOrThrow_(true);
  if (!articlesData || !Array.isArray(articlesData) || articlesData.length === 0) {
    throw new Error('No se proporcionaron datos válidos para importar');
  }
  
  var ss = getSpreadsheet_();
  var articlesSheet = ss.getSheetByName(SHEET_NAMES.ARTICLES);
  var hdrs = SHEET_HEADERS.ARTICLES;
  var map = getHeaderIndexMap_(articlesSheet, hdrs);
  
  // Obtener títulos existentes si skipDuplicates
  var existingTitles = {};
  if (skipDuplicates) {
    var existingRows = getRows_(articlesSheet);
    for (var i = 0; i < existingRows.length; i++) {
      var title = (existingRows[i][map['title']] || '').toString().trim().toLowerCase();
      if (title) existingTitles[title] = true;
    }
  }
  
  var imported = 0;
  var skipped = 0;
  var errors = [];
  var auth = { email: getActiveUserEmail_() };
  
  for (var i = 0; i < articlesData.length; i++) {
    try {
      var data = articlesData[i];
      var title = (data.title || '').toString().trim();
      if (!title) {
        errors.push('Artículo ' + (i + 1) + ': título requerido');
        continue;
      }
      
      // Verificar duplicados
      if (skipDuplicates && existingTitles[title.toLowerCase()]) {
        skipped++;
        continue;
      }
      
      var articleId = data.articleId || uuid_();
      var slug = ensureUniqueSlug_(slugify_(title));
      var description = sanitizeHtml_((data.descriptionHtml || data.description || '').toString());
      var driveFileId = extractDriveFileIdFromAny_((data.driveFileId || '').toString().trim());
      var driveMimeType = (data.driveMimeType || data.category || '').toString().trim();
      var typeId = (data.typeId || '').toString().trim();
      if (!typeId && driveMimeType) {
        try { typeId = getOrCreateTypeIdByName_(driveMimeType); } catch (_) {}
      }
      
      var createdAt = data.createdAt || toIsoString_(now_());
      var updatedAt = data.updatedAt || createdAt;
      var createdBy = data.createdBy || auth.email;
      var updatedBy = data.updatedBy || createdBy;
      var status = (data.status || 'PUBLISHED').toString().trim();
      
      var row = [];
      row[map['articleId']] = articleId;
      row[map['title']] = title;
      row[map['slug']] = slug;
      row[map['typeId']] = typeId;
      row[map['descriptionHtmlSanitized']] = description;
      row[map['driveFileId']] = driveFileId;
      row[map['driveMimeType']] = driveMimeType;
      row[map['createdAt']] = createdAt;
      row[map['updatedAt']] = updatedAt;
      row[map['createdBy']] = createdBy;
      row[map['updatedBy']] = updatedBy;
      row[map['status']] = status;
      
      articlesSheet.appendRow(row);
      imported++;
      
      // Aplicar restricciones si hay driveFileId
      if (driveFileId) {
        try {
          ensureRestrictionsForArticle_({ articleId: articleId, driveFileId: driveFileId });
        } catch (e) {
          errors.push('Artículo "' + title + '": error aplicando restricciones Drive: ' + (e && e.message || e));
        }
      }
      
      logAudit_('article', articleId, 'IMPORT', auth.email, { title: title, source: 'import' });
      
    } catch (e) {
      errors.push('Artículo ' + (i + 1) + ': ' + (e && e.message || e));
    }
  }
  
  return {
    imported: imported,
    skipped: skipped,
    errors: errors,
    message: 'Importación completada: ' + imported + ' importados, ' + skipped + ' omitidos, ' + errors.length + ' errores'
  };
}

/**
 * Importa artículos desde otra hoja de Google Sheets
 * @param {String} sourceSpreadsheetId - ID de la hoja de cálculo fuente
 * @param {String} sourceSheetName - Nombre de la hoja fuente (por defecto 'Articles')
 * @param {Object} columnMapping - Mapeo de columnas: {title: 'A', descriptionHtml: 'B', driveFileId: 'C', ...}
 * @param {Boolean} skipDuplicates - Si es true, omite duplicados
 */
function importArticlesFromSheet(sourceSpreadsheetId, sourceSheetName, columnMapping, skipDuplicates) {
  assertAccessOrThrow_(true);
  
  try {
    var sourceSheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    var sourceDataSheet = sourceSheet.getSheetByName(sourceSheetName || 'Articles');
    if (!sourceDataSheet) {
      throw new Error('No se encontró la hoja ' + (sourceSheetName || 'Articles') + ' en la hoja de cálculo fuente');
    }
    
    var sourceRows = sourceDataSheet.getDataRange().getValues();
    if (sourceRows.length < 2) {
      throw new Error('La hoja fuente no tiene datos (solo encabezados o está vacía)');
    }
    
    var headers = sourceRows[0];
    var headerMap = {};
    for (var h = 0; h < headers.length; h++) {
      headerMap[headers[h].toString().toLowerCase().trim()] = h;
    }
    
    // Si no hay mapeo personalizado, intentar mapeo automático
    if (!columnMapping) {
      columnMapping = {};
      var expected = ['title', 'description', 'drivefileid', 'drivemimetype', 'createdat', 'createdby'];
      for (var e = 0; e < expected.length; e++) {
        var key = expected[e];
        if (headerMap[key]) {
          columnMapping[key] = headerMap[key];
        }
      }
    }
    
    var articlesData = [];
    for (var r = 1; r < sourceRows.length; r++) {
      var row = sourceRows[r];
      var article = {};
      
      // Mapear columnas
      if (columnMapping.title !== undefined && row[columnMapping.title]) {
        article.title = row[columnMapping.title].toString().trim();
      }
      if (columnMapping.descriptionHtml !== undefined && row[columnMapping.descriptionHtml]) {
        article.descriptionHtml = row[columnMapping.descriptionHtml].toString();
      } else if (columnMapping.description !== undefined && row[columnMapping.description]) {
        article.descriptionHtml = row[columnMapping.description].toString();
      }
      if (columnMapping.driveFileId !== undefined && row[columnMapping.driveFileId]) {
        article.driveFileId = row[columnMapping.driveFileId].toString().trim();
      }
      if (columnMapping.driveMimeType !== undefined && row[columnMapping.driveMimeType]) {
        article.driveMimeType = row[columnMapping.driveMimeType].toString().trim();
      } else if (columnMapping.category !== undefined && row[columnMapping.category]) {
        article.driveMimeType = row[columnMapping.category].toString().trim();
      }
      if (columnMapping.createdAt !== undefined && row[columnMapping.createdAt]) {
        article.createdAt = row[columnMapping.createdAt];
      }
      if (columnMapping.createdBy !== undefined && row[columnMapping.createdBy]) {
        article.createdBy = row[columnMapping.createdBy].toString().trim();
      }
      if (columnMapping.articleId !== undefined && row[columnMapping.articleId]) {
        article.articleId = row[columnMapping.articleId].toString().trim();
      }
      
      if (article.title) {
        articlesData.push(article);
      }
    }
    
    if (articlesData.length === 0) {
      throw new Error('No se encontraron artículos válidos para importar');
    }
    
    return importArticlesFromData(articlesData, skipDuplicates || false);
    
  } catch (e) {
    logAudit_('import', '', 'ERROR', getActiveUserEmail_(), { message: e && e.message || e });
    throw e;
  }
}


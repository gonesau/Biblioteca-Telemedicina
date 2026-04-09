// ============================================================
//  DriveImport.gs — Biblioteca de Telemedicina
//
//  Importa automáticamente los archivos de la carpeta Drive
//  hacia la hoja Articles, creando tipos a partir de las
//  subcarpetas y artículos por cada archivo encontrado.
//
//  USO:
//    1. Ejecuta setupSheets() primero.
//    2. Llama importFromDriveFolder(FOLDER_ID) desde el editor.
//    3. Para reimportar solo novedades usa syncDriveFolder(FOLDER_ID).
// ============================================================

// ID de la carpeta raíz de Telemedicina (del link compartido).
// Puedes sobreescribir este valor al llamar las funciones, o dejarlo hardcodeado.
var DEFAULT_FOLDER_ID = '1ZsTIovrGb8IrfGv3MxVYJorbHvpW7XDc';

// Tipos MIME de Drive que se consideran "archivos" (no carpetas).
var IMPORTABLE_MIME_TYPES = {
  'application/vnd.google-apps.document':     'Documento',
  'application/vnd.google-apps.presentation': 'Presentación',
  'application/vnd.google-apps.spreadsheet':  'Hoja de cálculo',
  'application/pdf':                          'PDF',
  'video/mp4':                                'Video',
  'video/quicktime':                          'Video',
  'image/jpeg':                               'Imagen',
  'image/png':                                'Imagen',
  'audio/mpeg':                               'Audio',
  'application/zip':                          'Archivo comprimido'
};

// Mapeo de nombres de carpeta → categoría legible (ajusta según tu estructura).
var FOLDER_TYPE_OVERRIDES = {
  'videos':                                    'Video',
  'anuncios':                                  'Anuncio',
  'carrete de imágenes':                       'Imagen',
  'fondos de pantalla':                        'Imagen',
  'plataforma llamadas telefónicas':           'Protocolo',
  'scripts':                                   'Script',
  'programacion de turnos para servicios profesionales': 'Programación',
  'apoyo para la atencion medica':             'Apoyo Médico',
  'referencias':                               'Referencia',
  'flujos de atencion y procesos':             'Flujo de Atención'
};

// ----------------------------------------------------------------
// FUNCIÓN PRINCIPAL: importación completa desde la carpeta raíz
// ----------------------------------------------------------------
function importFromDriveFolder(folderId) {
  assertAccessOrThrow_(true);
  var id = folderId || DEFAULT_FOLDER_ID;
  var stats = { created: 0, skipped: 0, errors: 0 };

  try {
    var folder = DriveApp.getFolderById(id);
    _processFolder(folder, null, stats);
  } catch (e) {
    logAudit_('import', id, 'ERROR', getActiveUserEmail_(), { message: e.message });
    throw new Error('Error accediendo a la carpeta: ' + e.message);
  }

  logAudit_('import', id, 'IMPORT_COMPLETE', getActiveUserEmail_(), stats);
  invalidateArticlesCache_();
  return stats;
}

// ----------------------------------------------------------------
// FUNCIÓN DE SINCRONIZACIÓN: solo importa archivos nuevos
// (no existentes en la hoja Articles por driveFileId)
// ----------------------------------------------------------------
function syncDriveFolder(folderId) {
  assertAccessOrThrow_(true);
  var id = folderId || DEFAULT_FOLDER_ID;

  // Construir índice de fileIds ya importados
  var existing = _getExistingFileIds_();
  var stats = { created: 0, skipped: 0, errors: 0 };

  try {
    var folder = DriveApp.getFolderById(id);
    _processFolder(folder, null, stats, existing);
  } catch (e) {
    logAudit_('import', id, 'ERROR', getActiveUserEmail_(), { message: e.message });
    throw new Error('Error durante sincronización: ' + e.message);
  }

  logAudit_('import', id, 'SYNC_COMPLETE', getActiveUserEmail_(), stats);
  invalidateArticlesCache_();
  return stats;
}

// ----------------------------------------------------------------
// Recorre recursivamente la carpeta procesando subcarpetas y archivos
// ----------------------------------------------------------------
function _processFolder(folder, parentTypeName, stats, existingIds) {
  var folderName  = folder.getName();
  var typeName    = _resolveTypeName_(folderName, parentTypeName);

  // Procesar archivos directos en esta carpeta
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    try {
      _importFile_(file, typeName, stats, existingIds);
    } catch (e) {
      stats.errors++;
      logAudit_('import', file.getId(), 'FILE_ERROR', getActiveUserEmail_(), { message: e.message, name: file.getName() });
    }
  }

  // Recursar en subcarpetas
  var subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    var sub = subFolders.next();
    _processFolder(sub, typeName, stats, existingIds);
  }
}

// ----------------------------------------------------------------
// Importa un único archivo como artículo
// ----------------------------------------------------------------
function _importFile_(file, typeName, stats, existingIds) {
  var fileId   = file.getId();
  var mimeType = file.getMimeType();
  var name     = file.getName();

  // Saltar si ya existe
  if (existingIds && existingIds[fileId]) {
    stats.skipped++;
    return;
  }

  // Saltar carpetas (por si DriveApp las incluye)
  if (mimeType === 'application/vnd.google-apps.folder') {
    stats.skipped++;
    return;
  }

  // Determinar typeId
  var resolvedType = typeName || IMPORTABLE_MIME_TYPES[mimeType] || 'General';
  var typeId       = getOrCreateTypeIdByName_(resolvedType);

  // Construir artículo
  var sheet  = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLES);
  var hdrs   = SHEET_HEADERS.ARTICLES;
  var map    = getHeaderIndexMap_(sheet, hdrs);

  var articleId = uuid_();
  var title     = _cleanFileName_(name);
  var slug      = ensureUniqueSlug_(slugify_(title));
  var ts        = toIsoString_(now_());
  var actor     = getActiveUserEmail_();

  // Descripción automática desde metadatos del archivo
  var descHtml = _buildFileDescription_(file, resolvedType);

  var row = [];
  row[map['articleId']]               = articleId;
  row[map['title']]                   = title;
  row[map['slug']]                    = slug;
  row[map['typeId']]                  = typeId;
  row[map['descriptionHtmlSanitized']] = descHtml;
  row[map['driveFileId']]             = fileId;
  row[map['driveMimeType']]           = resolvedType;
  row[map['createdAt']]               = ts;
  row[map['updatedAt']]               = ts;
  row[map['createdBy']]               = actor;
  row[map['updatedBy']]               = actor;
  row[map['status']]                  = 'PUBLISHED';

  sheet.appendRow(row);

  // Registrar versión inicial
  appendVersion_({
    articleId:                articleId,
    title:                    title,
    slug:                     slug,
    typeId:                   typeId,
    descriptionHtmlSanitized: descHtml,
    driveFileId:              fileId,
    driveMimeType:            resolvedType,
    createdBy:                actor,
    updatedBy:                actor,
    status:                   'PUBLISHED',
    changeType:               'create',
    changesSummary:           'Importado desde Drive: ' + name
  });

  logAudit_('import', articleId, 'CREATE', actor, { driveFileId: fileId, title: title, type: resolvedType });
  stats.created++;
}

// ----------------------------------------------------------------
// Helpers privados
// ----------------------------------------------------------------

/** Retorna el nombre de tipo basado en el nombre de la carpeta o hereda del padre. */
function _resolveTypeName_(folderName, parentTypeName) {
  var key = (folderName || '').toLowerCase().trim();
  if (FOLDER_TYPE_OVERRIDES[key]) return FOLDER_TYPE_OVERRIDES[key];
  // Si el nombre no coincide con ningún override, usar el nombre de la carpeta capitalizado
  // o heredar del padre si estamos en un nivel profundo
  return folderName ? _capitalize_(folderName) : (parentTypeName || 'General');
}

/** Limpia la extensión del nombre de archivo para usarlo como título. */
function _cleanFileName_(name) {
  return (name || '').replace(/\.(mp4|pdf|doc|docx|xls|xlsx|ppt|pptx|jpg|jpeg|png|gif|mp3|zip)$/i, '').trim() || name;
}

/** Capitaliza primera letra. */
function _capitalize_(str) {
  if (!str) return str;
  return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}

/** Construye un HTML de descripción básico con metadatos del archivo. */
function _buildFileDescription_(file, typeName) {
  var lines = [];
  lines.push('<p><strong>Tipo:</strong> ' + escapeHtml_(typeName) + '</p>');

  var size = file.getSize();
  if (size > 0) {
    var sizeStr = size > 1048576
      ? (Math.round(size / 1048576 * 10) / 10) + ' MB'
      : (Math.round(size / 1024)) + ' KB';
    lines.push('<p><strong>Tamaño:</strong> ' + sizeStr + '</p>');
  }

  var modified = file.getLastUpdated();
  if (modified) {
    lines.push('<p><strong>Última actualización en Drive:</strong> ' + toIsoString_(modified).slice(0, 10) + '</p>');
  }

  return lines.join('\n');
}

/** Devuelve un objeto { fileId: true } de todos los driveFileId ya registrados. */
// ----------------------------------------------------------------
// Asegura filas en ArticleTypes para cada carpeta (sin importar archivos)
// ----------------------------------------------------------------
function _ensureTypesInFolderTree_(folder, parentTypeName) {
  var folderName = folder.getName();
  var typeName = _resolveTypeName_(folderName, parentTypeName);
  if (typeName) {
    try {
      getOrCreateTypeIdByName_(typeName);
    } catch (e) {
      logAudit_('import', folder.getId(), 'TYPE_ENSURE_ERROR', getActiveUserEmail_(), { message: e && e.message, typeName: typeName });
    }
  }
  var subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    var sub = subFolders.next();
    _ensureTypesInFolderTree_(sub, typeName);
  }
}

function ensureArticleTypesFromDriveFolder(folderId) {
  assertAccessOrThrow_(true);
  var id = folderId || DEFAULT_FOLDER_ID;
  try {
    var folder = DriveApp.getFolderById(id);
    _ensureTypesInFolderTree_(folder, null);
  } catch (e) {
    logAudit_('import', id, 'ERROR', getActiveUserEmail_(), { message: e.message, phase: 'ensureTypes' });
    throw new Error('Error recorriendo carpetas: ' + e.message);
  }
  logAudit_('import', id, 'ENSURE_TYPES_COMPLETE', getActiveUserEmail_(), {});
  invalidateArticlesCache_();
  return { ok: true };
}

function _getExistingFileIds_() {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.ARTICLES);
  var hdrs  = SHEET_HEADERS.ARTICLES;
  var map   = getHeaderIndexMap_(sheet, hdrs);
  var rows  = getRows_(sheet);
  var index = {};
  for (var i = 0; i < rows.length; i++) {
    var fid = rows[i][map['driveFileId']];
    if (fid) index[fid] = true;
  }
  return index;
}

// ----------------------------------------------------------------
// Utilidad de menú: lanzar import con confirmación (llamada desde Menu.gs)
// ----------------------------------------------------------------
function menu_importDriveFolder_() {
  var ui     = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Importar desde Drive',
    'Ingresa el ID de la carpeta de Drive (deja vacío para usar la carpeta por defecto de Telemedicina):',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  var inputId = result.getResponseText().trim() || DEFAULT_FOLDER_ID;
  try {
    var stats = importFromDriveFolder(inputId);
    ui.alert('Importación completada ✓\n\nCreados: ' + stats.created + '\nOmitidos: ' + stats.skipped + '\nErrores: ' + stats.errors);
  } catch (e) {
    ui.alert('Error durante la importación:\n' + e.message);
  }
}

function menu_syncDriveFolder_() {
  var ui = SpreadsheetApp.getUi();
  try {
    var stats = syncDriveFolder(DEFAULT_FOLDER_ID);
    ui.alert('Sincronización completada ✓\n\nNuevos: ' + stats.created + '\nYa existentes: ' + stats.skipped + '\nErrores: ' + stats.errors);
  } catch (e) {
    ui.alert('Error durante la sincronización:\n' + e.message);
  }
}

function menu_ensureArticleTypesFromDrive_() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Categorías desde carpetas',
    'ID de carpeta raíz en Drive (vacío = carpeta por defecto de Telemedicina):',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;
  var inputId = result.getResponseText().trim() || DEFAULT_FOLDER_ID;
  try {
    ensureArticleTypesFromDriveFolder(inputId);
    ui.alert('Listo: se crearon o confirmaron categorías en ArticleTypes según la jerarquía de carpetas.');
  } catch (e) {
    ui.alert('Error:\n' + e.message);
  }
}
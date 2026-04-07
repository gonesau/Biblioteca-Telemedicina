// Sheet names and headers for normalized data model

var SHEET_NAMES = {
  ARTICLES: 'Articles',
  ARTICLE_VERSIONS: 'ArticleVersions',
  AUDIT_LOG: 'AuditLog',
  NOTIFICATIONS_CONFIG: 'NotificationsConfig',
  ACCESS_CONFIG: 'AccessConfig',
  ARTICLE_TYPES: 'ArticleTypes',
  TAGS: 'Tags',
  ARTICLE_TAGS: 'ArticleTags',
  SITE_CONFIG: 'SiteConfig',
  ARTICLE_COMMENTS: 'ArticleComments'
};

var SHEET_HEADERS = {
  ARTICLES: [
    'articleId', 'title', 'slug', 'typeId', 'descriptionHtmlSanitized',
    'driveFileId', 'driveMimeType', 'createdAt', 'updatedAt',
    'createdBy', 'updatedBy', 'status'
  ],
  ARTICLE_VERSIONS: [
    'versionId', 'articleId', 'versionNumber', 'title',
    'descriptionHtmlSanitized', 'driveFileId', 'driveMimeType',
    'changedAt', 'changedBy', 'changeType', 'changesSummary', 'status'
  ],
  AUDIT_LOG: [
    'auditId', 'entityType', 'entityId', 'action', 'actor', 'at', 'metadata'
  ],
  NOTIFICATIONS_CONFIG: [
    'recipientEmail', 'enabled', 'replyTo', 'senderAlias'
  ],
  ACCESS_CONFIG: [
    'adminEmail', 'allowedDomains', 'allowedEmails', 'guestEditorEmails'
  ],
  ARTICLE_TYPES: [
    'typeId', 'name', 'slug', 'color', 'order'
  ],
  TAGS: [
    'tagId', 'name', 'slug'
  ],
  ARTICLE_TAGS: [
    'articleId', 'tagId'
  ],
  SITE_CONFIG: [
    'key', 'value', 'updatedAt', 'updatedBy'
  ],
  ARTICLE_COMMENTS: [
    'commentId', 'articleId', 'text', 'createdAt', 'createdBy'
  ]
};

function getSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  // Fallback para proyectos no vinculados: usar SPREADSHEET_ID desde propiedades de script
  var props = PropertiesService.getScriptProperties();
  var id = props && props.getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  throw new Error('No se pudo determinar la hoja. Configure SPREADSHEET_ID en Propiedades del script o vincule el proyecto a una hoja.');
}



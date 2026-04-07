function applyDriveRestrictions_(fileId) {
  if (!fileId) return;
  try {
    // Drive Advanced Service v2: aplicar settings de copia/descarga
    // - viewersCanCopyContent: false → evita copiar contenido a lectores/comentaristas
    // - copyRequiresWriterPermission: true → requiere permiso de editor para descargar/imprimir/copiar
    Drive.Files.update({ viewersCanCopyContent: false, copyRequiresWriterPermission: true }, fileId);
  } catch (e) {
    try { logAudit_('drive', fileId, 'RESTRICT_ERROR', getActiveUserEmail_(), { message: e && e.message }); } catch (ignored) {}
    throw e;
  }
}

function ensureRestrictionsForArticle_(article) {
  if (!article || !article.driveFileId) return;
  try {
    applyDriveRestrictions_(article.driveFileId);
    logAudit_('drive', article.articleId || '', 'RESTRICT_APPLIED', getActiveUserEmail_(), { fileId: article.driveFileId });
  } catch (e) {
    // No rethrow to avoid blocking main flow silently; we already audited the error in applyDriveRestrictions_
  }
}



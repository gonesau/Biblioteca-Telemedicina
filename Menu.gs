// ============================================================
//  Menu.gs — Biblioteca de Telemedicina
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Biblioteca Telemedicina')
    .addItem('🔧 Inicializar hojas', 'setupSheets')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Drive')
      .addItem('Importar desde carpeta Drive...', 'menu_importDriveFolder_')
      .addItem('Sincronizar novedades (solo nuevos)', 'menu_syncDriveFolder_')
      .addItem('Asegurar categorías desde carpetas (sin archivos)', 'menu_ensureArticleTypesFromDrive_'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Migraciones')
      .addItem('Migrar tipos y tags (v1)', 'migrate_v1_types_tags')
      .addItem('Aplicar restricciones Drive (batch)', 'menu_apply_drive_restrictions_batch_'))
    .addToUi();
}

function menu_apply_drive_restrictions_batch_() {
  var res = migrate_v1_drive_restrictions(50);
  SpreadsheetApp.getActive().toast(
    'Procesado: ' + res.processed + ' | Próximo: ' + res.nextCursor + ' / ' + res.total
  );
}
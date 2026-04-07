function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Artículos')
    .addItem('Inicializar hojas', 'setupSheets')
    .addSeparator()
    .addItem('Migrar tipos y tags (v1)', 'migrate_v1_types_tags')
    .addItem('Aplicar restricciones Drive (batch)', 'menu_apply_drive_restrictions_batch_')
    .addSeparator()
    .addItem('Importar artículos desde otra hoja...', 'menu_import_from_sheet_')
    .addToUi();
}

function menu_import_from_sheet_() {
  var html = HtmlService.createHtmlOutputFromFile('ImportDialog').setWidth(560).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Importar artículos');
}

function menu_apply_drive_restrictions_batch_() {
  var res = migrate_v1_drive_restrictions(50);
  SpreadsheetApp.getActive().toast('Procesado: ' + res.processed + ' | Próximo cursor: ' + res.nextCursor + ' / ' + res.total);
}



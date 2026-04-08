// ============================================================
//  Setup.gs — Biblioteca de Telemedicina
//  Inicializa las hojas. Ejecutar UNA sola vez.
// ============================================================

function setupSheets() {
  var ss = getSpreadsheet_();

  for (var key in SHEET_NAMES) {
    var name    = SHEET_NAMES[key];
    var headers = SHEET_HEADERS[key];
    if (!headers) continue;

    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    ensureHeaders_(sheet, headers);
  }

  try {
    ss.toast('✓ Hojas inicializadas correctamente', 'Biblioteca Telemedicina', 5);
  } catch (_) {
    Logger.log('setupSheets completado correctamente.');
  }
}
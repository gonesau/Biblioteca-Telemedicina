// ============================================================
//  WebApp.gs — Biblioteca de Telemedicina
//  Punto de entrada de la Web App.
// ============================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Biblioteca Telemedicina')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Lanzador manual desde el editor de scripts
function runImport() {
  importFromDriveFolder('1ZsTIovrGb8IrfGv3MxVYJorbHvpW7XDc');
}
/**
 * Erstellt das benutzerdefinierte Menü beim Öffnen der Tabellenkalkulation
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(' Update Sheet')
    .addItem('Update', 'main')
    .addToUi();
}

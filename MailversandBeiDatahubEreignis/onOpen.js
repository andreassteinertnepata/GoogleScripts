function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Update Sheet') // So heißt das Menü ganz oben
    .addItem('Update Lagerbestand', 'lagerbestandAktualisieren')
    .addItem('Verkäufe prüfen', 'verkaeufeAktionPruefenUndSenden')
    .addItem('Update all', 'main')
    .addToUi();
  
}
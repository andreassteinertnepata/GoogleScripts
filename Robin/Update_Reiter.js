function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Update Sheet') // So heißt das Menü ganz oben
    .addItem('Update All', 'main')
    .addToUi();
  
}
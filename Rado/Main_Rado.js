function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    showStatusBox("Update Running", "Please wait");
    Utilities.sleep(500); 

    fetchOffenePosten();

    showStatusBox("Update Running", "Loading open orders...");
    fetchVorgaenge();

    showStatusBox("Update Running", "Loading item data...");
    fetchArticles();
    fetchStuecklisten() 
    
    showStatusBox("Update Running", "Loading customer data...");
    fetchCustomers() 

    showStatusBox("Update Running", "Loading history data...");
    importArchivKomplett();
    
    showStatusBox("Translating", "Cells and headers...");
    translateHeadersFromExternalSheet();
    translateValuesFromExternalSheet();

    showStatusBox("Finishing", "Layout and Filters...");
    formatGoogleSheet();

    closeStatusBox();
    ss.toast("Everything is up to date.", "Finished!", 5);

  } catch (e) {
    closeStatusBox();
    console.error("ERROR: " + e.toString());
  }
}


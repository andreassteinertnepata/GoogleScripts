function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vtrNr = String(CONFIG.VERTRETER_NR).trim();
  const isEnglish = (CONFIG.ENGLISH_REPS || []).includes(vtrNr);

  try {
    showStatusBox("Update Running", "Please wait");
    Utilities.sleep(500); 

    showStatusBox("Update Running", "Loading outstanding invoices...");
    fetchOffenePosten();

    showStatusBox("Update Running", "Loading open orders...");
    fetchVorgaenge();

    showStatusBox("Update Running", "Loading item data...");
    fetchArticles();
    fetchStuecklisten();

    showStatusBox("Update Running", "Loading history data...");
    importArchivKomplett();
    
    showStatusBox("Update Running", "Loading customer data...");
    fetchCustomers();
    
    // Übersetzungen nur triggern, wenn der Vertriebler in ENGLISH_REPS definiert ist
    if (isEnglish) {
      showStatusBox("Translating", "Cells and headers...");
      if (typeof translateHeadersFromExternalSheet === "function") {
        translateHeadersFromExternalSheet();
      }
      if (typeof translateValuesFromExternalSheet === "function") {
        translateValuesFromExternalSheet();
      }
    }

    showStatusBox("Finishing", "Layout and Filters...");
    formatGoogleSheet();

    closeStatusBox();
    ss.toast("Everything is up to date.", "Finished!", 5);

  } catch (e) {
    closeStatusBox();
    console.error("ERROR: " + e.toString());
    ss.toast("Fehler beim Update: " + e.message, "Fehler", 10);
  }
}
function showStatusBox(title, message) {
  try {
    // Prüft, ob eine Benutzeroberfläche vorhanden ist (ist nur der Fall, wenn ein Benutzer aktiv im Sheet ist)
    const ui = SpreadsheetApp.getUi(); 
    
    const htmlContent = `
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <style>
            body { font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; margin: 0; padding: 16px; background-color: #f8f9fa; color: #2c3e50; display: flex; flex-direction: column; align-items: center; justify-content: center; }
            .card { background: #ffffff; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); padding: 20px; width: 100%; box-sizing: border-box; text-align: center; }
            .spinner { border: 4px solid #f3f3f3; border-top: 4px solid #3498db; border-radius: 50%; width: 32px; height: 32px; animation: spin 1s linear infinite; margin: 0 auto 15px auto; }
            @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
            h3 { margin: 0 0 8px 0; font-size: 16px; }
            p { margin: 0; font-size: 13px; color: #7f8c8d; }
          </style>
        </head>
        <body>
          <div class="card">
            <div class="spinner"></div>
            <h3>${title}</h3>
            <p>${message}</p>
          </div>
        </body>
      </html>
    `;
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(320).setHeight(150);
    ui.showModelessDialog(htmlOutput, " ");
  } catch (e) {
    // Falls das Skript über den zeitgesteuerten Trigger läuft, gibt getUi() einen Fehler aus.
    // Wir fangen ihn hier ab und protokollieren stattdessen einfach im Log!
    Logger.log("Hintergrundausführung: " + title + " - " + message);
  }
}

function closeStatusBox() {
  try {
    const ui = SpreadsheetApp.getUi();
    const closeScript = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>").setWidth(10).setHeight(10);
    ui.showModelessDialog(closeScript, " ");
  } catch (e) {
    // Im Hintergrund-Trigger passiert nichts und es gibt keinen Fehler
  }
}
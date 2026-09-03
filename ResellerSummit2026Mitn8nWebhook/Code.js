// 1. Erstellt das benutzerdefinierte Menü beim Öffnen der Datei
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Update')
    .addItem('An n8n senden', 'triggerN8nWebhook')
    .addToUi();
}

// 2. Liest die Tabellendaten aus und sendet sie an den n8n Webhook
function triggerN8nWebhook() {
  // DEINE N8N WEBHOOK-URL HIER EINFÜGEN:
  const webhookUrl = "https://nepata.app.n8n.cloud/webhook/5c1e6a08-54d6-444e-a9dd-3d8369329b13";
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // Abbruch, falls die Tabelle leer ist
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("Keine Daten zum Senden vorhanden.");
    return;
  }
  
  // Erste Zeile als Spaltenüberschriften (Keys) nutzen
  const headers = data[0];
  const payloadData = [];
  
  // Zeilen ab Zeile 2 in JSON-Objekte umwandeln
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowObject = {};
    for (let j = 0; j < headers.length; j++) {
      rowObject[headers[j]] = row[j];
    }
    payloadData.push(rowObject);
  }
  
  // HTTP POST Request konfigurieren
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      sheetName: sheet.getName(),
      data: payloadData
    }),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    const responseCode = response.getResponseCode();
    
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Fehler: ${error.toString()}`);
  }
  formatSheet()
}
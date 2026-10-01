// ==========================================
// FILE: EmailTemplateService.gs
// ==========================================
function buildEmailHtml(lang, repName, vtrNr, docs) {
  const isEn = lang === "EN";
  
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.5; }
        .section-title { color: #2c3e50; font-size: 16px; font-weight: bold; margin-top: 25px; margin-bottom: 10px; border-bottom: 2px solid #2c3e50; padding-bottom: 5px; }
        table.styled-table { width: 100%; border-collapse: collapse; margin-top: 8px; margin-bottom: 20px; font-size: 13px; }
        table.styled-table th { background-color: #2c3e50; color: #ffffff; padding: 8px 10px; text-align: left; font-weight: bold; border: 1px solid #1a252f; }
        table.styled-table td { padding: 8px 10px; border: 1px solid #e2e8f0; vertical-align: middle; }
        table.styled-table tr:nth-child(even) { background-color: #f8fafc; }
        .order-card { background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px; padding: 12px; margin-bottom: 20px; }
        .order-card-header { background-color: #f1f5f9; padding: 8px 12px; margin: -12px -12px 10px -12px; border-bottom: 1px solid #cbd5e1; font-weight: bold; color: #1e293b; border-radius: 5px 5px 0 0; }
        .text-right { text-align: right; }
        .text-center { text-align: center; }
        .badge { background-color: #e2e8f0; color: #334155; padding: 2px 6px; border-radius: 4px; font-size: 12px; font-weight: bold; }
      </style>
    </head>
    <body>
      <p>${isEn ? `Hello ${repName},` : `Hallo ${repName},`}</p>
      <p>${isEn 
        ? `The following prepayment (Vorkasse) orders are older than 3 months and are still open.<br><strong>Should these be cancelled?</strong>` 
        : `die folgenden Vorkasse-Vorgänge sind bereits älter als 3 Monate und noch offen.<br><strong>Sollen diese storniert werden?</strong>`}</p>
      
      <!-- 1. KURZÜBERSICHT -->
      <div class="section-title">${isEn ? 'SUMMARY / SHORT FORM' : 'ÜBERSICHT / KURZFORM'}</div>
      
      <table class="styled-table">
        <thead>
          <tr>
            <th>${isEn ? 'Doc No' : 'Beleg-Nr'}</th>
            <th>${isEn ? 'Order No' : 'Auftrags-Nr'}</th>
            <th class="text-center">${isEn ? 'Type' : 'Art'}</th>
            <th class="text-center">${isEn ? 'Rep ID' : 'Vertreter'}</th>
            <th>${isEn ? 'Customer' : 'Kunde'}</th>
            <th class="text-center">${isEn ? 'Doc Date' : 'Belegdatum'}</th>
            <th class="text-center">${isEn ? 'Delivery Date' : 'Lieferdatum'}</th>
            <th class="text-right">${isEn ? 'TOTAL Qty' : 'TOTAL Menge'}</th>
            <th class="text-right">${isEn ? 'TOTAL Net' : 'TOTAL Netto'}</th>
          </tr>
        </thead>
        <tbody>
  `;

  docs.forEach(doc => {
    const custName = `${doc.reNa2}${doc.reNa3 ? ' (' + doc.reNa3 + ')' : ''}`;
    html += `
      <tr>
        <td><strong>${doc.belegNr}</strong></td>
        <td>${doc.auftrNr || '-'}</td>
        <td class="text-center"><span class="badge">${doc.art}</span></td>
        <td class="text-center">${doc.vtrNr || '-'}</td>
        <td>${doc.adrNr} - ${custName}</td>
        <td class="text-center">${formatDateIso(doc.docDate)}</td>
        <td class="text-center">${formatDateIso(doc.liefDate)}</td>
        <td class="text-right"><strong>${formatNumber(doc.totalMge)}</strong></td>
        <td class="text-right"><strong>${formatMoney(doc.totalNetto)}</strong></td>
      </tr>
    `;
  });

  html += `
        </tbody>
      </table>

      <!-- 2. AUFTRAGSDETAILS JE BELEG -->
      <div class="section-title">${isEn ? 'ORDER DETAILS (POSITIONS)' : 'AUFTRAGSDETAILS (POSITIONEN)'}</div>
  `;

  docs.forEach(doc => {
    const custName = `${doc.reNa2}${doc.reNa3 ? ' (' + doc.reNa3 + ')' : ''}`;
    html += `
      <div class="order-card">
        <div class="order-card-header">
          ${isEn ? 'Doc No' : 'Beleg-Nr'}: ${doc.belegNr} &nbsp;|&nbsp; ${isEn ? 'Order No' : 'Auftrags-Nr'}: ${doc.auftrNr || '-'} &nbsp;|&nbsp; ${isEn ? 'Customer' : 'Kunde'}: ${doc.adrNr} - ${custName}
        </div>
        <table class="styled-table" style="margin-bottom: 0;">
          <thead>
            <tr>
              <th>${isEn ? 'Item No' : 'Art.-Nr'}</th>
              <th>${isEn ? 'Description' : 'Artikelbezeichnung'}</th>
              <th class="text-right">${isEn ? 'Qty' : 'Menge'}</th>
              <th class="text-right">${isEn ? 'Unit Price' : 'Einzelpreis'}</th>
              <th class="text-right">${isEn ? 'Total Net' : 'Gesamt Netto'}</th>
            </tr>
          </thead>
          <tbody>
    `;

    if (doc.positions.length > 0) {
      doc.positions.forEach(p => {
        const desc = [p.kuBez1, p.kuBez3].filter(Boolean).join(' / ') || '-';
        html += `
          <tr>
            <td>${p.artNr || '-'}</td>
            <td>${desc}</td>
            <td class="text-right">${formatNumber(p.mge)}</td>
            <td class="text-right">${formatMoney(p.eprNt)}</td>
            <td class="text-right">${formatMoney(p.lineNetto)}</td>
          </tr>
        `;
      });
    } else {
      html += `
        <tr>
          <td colspan="5" class="text-center" style="color: #888888;">${isEn ? '(No position items found)' : '(Keine Positionen vorhanden)'}</td>
        </tr>
      `;
    }

    html += `
          </tbody>
        </table>
      </div>
    `;
  });

  html += `
      <p style="margin-top: 20px;">${isEn ? 'Please check and process/cancel them in büro+ accordingly.' : 'Bitte prüfen und ggf. in büro+ entsprechend bearbeiten/stornieren.'}</p>
      <p>${isEn ? 'Best regards,<br>This email has been automatically generated' : 'Viele Grüße,<br>Diese Mail wurde automatisch generiert'}</p>
    </body>
    </html>
  `;

  return html;
}

function buildEmailText(lang, repName, vtrNr, docs) {
  let body = "";
  const isEn = lang === "EN";

  if (isEn) {
    body += `Hello ${repName},\n\n`;
    body += `The following prepayment (Vorkasse) orders are older than 3 months and are still open.\nShould these be cancelled?\n\n`;
    body += `SUMMARY / SHORT FORM\n--------------------------------------------------\n`;
    docs.forEach((doc, idx) => {
      body += `${idx + 1}) Doc No: ${doc.belegNr} | Order No: ${doc.auftrNr} | Type: ${doc.art} | Rep ID: ${doc.vtrNr || "-"}\n`;
      body += `   Customer: ${doc.adrNr} - ${doc.reNa2}\n`;
      body += `   Doc Date: ${formatDateIso(doc.docDate)} | Delivery Date: ${formatDateIso(doc.liefDate)}\n`;
      body += `   -> TOTAL Qty: ${formatNumber(doc.totalMge)} | TOTAL Net: ${formatMoney(doc.totalNetto)}\n\n`;
    });
    body += `ORDER DETAILS\n--------------------------------------------------\n`;
    docs.forEach(doc => {
      body += `[DOC NO: ${doc.belegNr} | Order: ${doc.auftrNr} | Customer: ${doc.adrNr}]\n`;
      doc.positions.forEach(p => {
        body += `  * Item: ${p.artNr || "N/A"} | Qty: ${formatNumber(p.mge)} | Price: ${formatMoney(p.eprNt)} | Total: ${formatMoney(p.lineNetto)}\n`;
      });
      body += `\n`;
    });
    body += `Please check and process/cancel them in büro+ accordingly.\n\nBest regards,\nThis email has been automatically generated`;
  } else {
    body += `Hallo ${repName},\n\n`;
    body += `die folgenden Vorkasse-Vorgänge sind bereits älter als 3 Monate und noch offen.\nSollen diese storniert werden?\n\n`;
    body += `ÜBERSICHT / KURZFORM\n--------------------------------------------------\n`;
    docs.forEach((doc, idx) => {
      body += `${idx + 1}) Beleg-Nr: ${doc.belegNr} | Auftrags-Nr: ${doc.auftrNr} | Art: ${doc.art} | Vertreter: ${doc.vtrNr || "-"}\n`;
      body += `   Kunde: ${doc.adrNr} - ${doc.reNa2}\n`;
      body += `   Belegdatum: ${formatDateIso(doc.docDate)} | Lief.Datum: ${formatDateIso(doc.liefDate)}\n`;
      body += `   -> TOTAL Menge: ${formatNumber(doc.totalMge)} | TOTAL Netto: ${formatMoney(doc.totalNetto)}\n\n`;
    });
    body += `AUFTRAGSDETAILS\n--------------------------------------------------\n`;
    docs.forEach(doc => {
      body += `[BELEG-NR: ${doc.belegNr} | Auftrag: ${doc.auftrNr} | Kunde: ${doc.adrNr}]\n`;
      doc.positions.forEach(p => {
        body += `  * Art.-Nr: ${p.artNr || "keine"} | Menge: ${formatNumber(p.mge)} | Einzelpreis: ${formatMoney(p.eprNt)} | Gesamt: ${formatMoney(p.lineNetto)}\n`;
      });
      body += `\n`;
    });
    body += `Bitte prüfen und ggf. in büro+ entsprechend bearbeiten/stornieren.\n\nViele Grüße\nDiese Mail wurde automatisch generiert`;
  }

  return body;
}
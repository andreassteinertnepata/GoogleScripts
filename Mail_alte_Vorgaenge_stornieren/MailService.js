// ==========================================
// FILE: MailService.gs
// ==========================================
function sendVorkasseMails(ss, belegeNachVertreter, totalFoundDocs) {
  const now = new Date();

  // Log-Blatt vor jedem Mail-Versand einmalig leeren & vorbereiten
  initLogSheet(ss);

  for (const [vtrNr, docs] of Object.entries(belegeNachVertreter)) {
    let toEmail = "";
    let lang = "DE";
    let repName = "";

    if (vtrNr === "FALLBACK" || !REPS[vtrNr]) {
      toEmail = CONFIG.MAIL_FALLBACK;
      lang = "DE";
      repName = "Team (Unzugeordnet)";
    } else {
      toEmail = REPS[vtrNr].email;
      lang = REPS[vtrNr].lang;
      repName = REPS[vtrNr].name;
    }

    const subject = lang === "EN"
      ? `Cancellation Check Required: Open Prepayment Orders (>3 Months)`
      : `Stornoprüfung erforderlich: Offene Vorkasse-Belege (>3 Monate)`;

    const htmlBody = buildEmailHtml(lang, repName, vtrNr, docs);
    const plainTextBody = buildEmailText(lang, repName, vtrNr, docs);

    try {
      MailApp.sendEmail({
        to: toEmail,
        subject: subject,
        body: plainTextBody,
        htmlBody: htmlBody
      });
      Logger.log(`E-Mail gesendet an: ${toEmail} | Belege: ${docs.length}`);

      const belegNrListe = docs.map(d => d.belegNr);
      writeLogEntry(ss, now, vtrNr, repName, toEmail, belegNrListe);

    } catch (e) {
      Logger.log(`Fehler beim Senden der E-Mail an ${toEmail}: ${e.message}`);
    }
  }

  safeToast(`Erfolgreich Mails für ${totalFoundDocs} Belege versendet & protokolliert.`, "Abgeschlossen");
}
// ==========================================
// FILE: Config.gs
// ==========================================
const CONFIG = {
  API_URL: "https://datahub.launchpad.nepata.cloud/v2/nepata_vertrieb/graphql",
  API_TOKEN: "e12Bfv!@Ss#asrpPFjucm8a8",
  MAIL_FALLBACK: "andreas.steinert+test@nepata.de", // Fallback-Empfänger für unzugeordnete Belege
  VORGANGSARTEN: ["106", "112", "118", "120", "164"], // Feste Vorkasse-Vorgangsarten
  TAB_DETAILS: "Auftragsdetails",
  TAB_LOG: "Mail_Log",
  MAX_EXECUTION_TIME_MS: 5.5 * 60 * 1000 // 5,5 Minuten Time-Out-Bremse
};

// Vertreter-Konfiguration mit Sprache (DE/EN)
const REPS = {
  "28": { email: "andreas.steinert+test@nepata.de", lang: "DE", name: "Alexis" },
  "36": { email: "andreas.steinert+test@nepata.de", lang: "DE", name: "Imad" },
  "43": { email: "andreas.steinert+test@nepata.de", lang: "DE", name: "Andreas" },
  "46": { email: "andreas.steinert+test@nepata.de", lang: "DE", name: "Stefanie" },
  "56": { email: "andreas.steinert+test@nepata.de", lang: "EN", name: "Robin" },
  "59": { email: "andreas.steinert+test@nepata.de", lang: "DE", name: "Aneta" },
  "60": { email: "andreas.steinert+test@nepata.de", lang: "EN", name: "Rado" }
};
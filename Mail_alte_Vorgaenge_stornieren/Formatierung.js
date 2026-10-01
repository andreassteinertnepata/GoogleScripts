// ==========================================
// FILE: Utils.gs
// ==========================================
function safeToast(msg, title) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, title);
  } catch (e) {
    Logger.log(`${title}: ${msg}`);
  }
}

function formatDateIso(d) {
  if (!d) return "-";
  if (typeof d === "string") d = new Date(d);
  if (isNaN(d.getTime())) return "-";
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function formatNumber(num) {
  const n = Number(num) || 0;
  return n.toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatMoney(num) {
  return formatNumber(num) + " €";
}
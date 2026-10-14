/**
 * JUNIOR 40 – THE JOURNEY
 * Enregistre chaque réservation du site dans ce Google Sheet.
 *
 * Installation : voir google-apps-script/INSTALLATION.md
 */

const SHEET_NAME = "Réservations";
const HEADERS = [
  "Date de réponse", "Référence", "Prénom", "Nom", "Présence",
  "Passagers", "Siège", "Téléphone WhatsApp", "Message"
];

/** Reçoit les réservations envoyées par le site. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents || "{}");
    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      clean_(d.reference),
      clean_(d.prenom),
      clean_(d.nom),
      clean_(d.presence),
      Number(d.passagers) || 0,
      clean_(d.siege),
      clean_(d.telephone),
      clean_(d.message)
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Permet de vérifier que l'application Web répond (ouvrir l'URL /exec). */
function doGet() {
  return ContentService.createTextOutput("JUNIOR 40 – enregistrement des réservations : OK");
}

/** À lancer une fois depuis l'éditeur pour créer l'onglet et le récapitulatif. */
function setup() {
  getSheet_();
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight("bold").setBackground("#0b1530").setFontColor("#efe5cf");
    sheet.setFrozenRows(1);
    sheet.getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm");
    sheet.getRange("H:H").setNumberFormat("@");

    // Récapitulatif pour le traiteur
    sheet.getRange("K1:L1").setValues([["RÉCAPITULATIF", ""]]).setFontWeight("bold");
    sheet.getRange("K2:K5").setValues([["Réponses reçues"], ["Réponses « Oui »"], ["Réponses « Non »"], ["Total passagers confirmés"]]);
    sheet.getRange("L2").setFormula('=COUNTA(B2:B)');
    sheet.getRange("L3").setFormula('=COUNTIF(E2:E,"Oui")');
    sheet.getRange("L4").setFormula('=COUNTIF(E2:E,"Non")');
    sheet.getRange("L5").setFormula('=SUMIF(E2:E,"Oui",F2:F)');
    sheet.getRange("K5:L5").setFontWeight("bold").setBackground("#efe5cf");
    sheet.autoResizeColumns(1, 12);
  }
  return sheet;
}

/** Évite qu'un texte saisi soit interprété comme une formule. */
function clean_(v) {
  const s = String(v == null ? "" : v).slice(0, 1000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

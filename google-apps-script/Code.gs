/**
 * JUNIOR 40 – THE JOURNEY
 * Enregistre chaque réservation du site dans ce Google Sheet.
 *
 * Installation : voir google-apps-script/INSTALLATION.md
 */

// Facultatif : l'identifiant du Google Sheet (la partie entre /d/ et /edit dans son adresse).
// À remplir uniquement si le script n'a PAS été créé depuis le Sheet (Extensions → Apps Script).
const SHEET_ID = "";

const SHEET_NAME = "Réservations";
const HEADERS = [
  "Date de réponse", "Référence", "Prénom", "Nom", "Présence",
  "Passagers", "Siège", "Téléphone WhatsApp", "Message", "Accompagnants"
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
      // Le site recopie les accompagnants dans le message (sécurité) : on les retire ici
      clean_(String(d.message || "").replace(/\s*\|?\s*Accompagnants : .*$/, "")),
      clean_(d.accompagnants)
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Vérification depuis le navigateur :
 *  - …/exec          → affiche OK si l'application Web répond
 *  - …/exec?test=1   → ajoute une ligne « TEST » dans l'onglet Réservations
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.test) {
    try {
      const sheet = getSheet_();
      sheet.appendRow([new Date(), "TEST", "Test", "Navigateur", "Oui", 2, "", "", "Ligne de test, à supprimer", "Accompagnant Test"]);
      return ContentService.createTextOutput(
        "✅ Test réussi : une ligne TEST a été ajoutée dans l'onglet « " + SHEET_NAME +
        " » du fichier « " + sheet.getParent().getName() + " ».");
    } catch (err) {
      return ContentService.createTextOutput("❌ Erreur : " + err);
    }
  }
  return ContentService.createTextOutput("JUNIOR 40 – enregistrement des réservations : OK (version 2, avec accompagnants)");
}

/** À lancer une fois depuis l'éditeur pour créer l'onglet et le récapitulatif. */
function setup() {
  getSheet_();
}

function getSheet_() {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Aucun Google Sheet relié : remplissez SHEET_ID en haut du script.");
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight("bold").setBackground("#0b1530").setFontColor("#efe5cf");
    sheet.setFrozenRows(1);
    sheet.getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm");
    sheet.getRange("H:H").setNumberFormat("@");

    // Récapitulatif pour le traiteur
    sheet.getRange("L1:M1").setValues([["RÉCAPITULATIF", ""]]).setFontWeight("bold");
    sheet.getRange("L2:L5").setValues([["Réponses reçues"], ["Réponses « Oui »"], ["Réponses « Non »"], ["Total passagers confirmés"]]);
    sheet.getRange("M2").setFormula('=COUNTA(B2:B)');
    sheet.getRange("M3").setFormula('=COUNTIF(E2:E,"Oui")');
    sheet.getRange("M4").setFormula('=COUNTIF(E2:E,"Non")');
    sheet.getRange("M5").setFormula('=SUMIF(E2:E,"Oui",F2:F)');
    sheet.getRange("L5:M5").setFontWeight("bold").setBackground("#efe5cf");
    sheet.autoResizeColumns(1, 13);
  }
  // Ajoute la colonne « Accompagnants » à un onglet créé avec une ancienne version
  if (sheet.getRange(1, HEADERS.length).getValue() !== HEADERS[HEADERS.length - 1]) {
    sheet.getRange(1, HEADERS.length).setValue(HEADERS[HEADERS.length - 1])
      .setFontWeight("bold").setBackground("#0b1530").setFontColor("#efe5cf");
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

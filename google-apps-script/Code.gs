/**
 * JUNIOR 40 – THE JOURNEY
 * Enregistre chaque réservation du site dans ce Google Sheet.
 * Version 3 : une ligne par personne (titulaire + accompagnants).
 *
 * Installation : voir google-apps-script/INSTALLATION.md
 */

// Facultatif : l'identifiant du Google Sheet (la partie entre /d/ et /edit dans son adresse).
// À remplir uniquement si le script n'a PAS été créé depuis le Sheet (Extensions → Apps Script).
const SHEET_ID = "";

const SHEET_NAME = "Réservations";
const VERSION = "version 3, une ligne par personne";
const HEADERS = [
  "Date de réponse", "Référence", "Prénom", "Nom", "Présence",
  "Passagers", "Siège", "Téléphone WhatsApp", "Message", "Rôle"
];
const ROLE_COL = HEADERS.length;          // colonne J
const HEAD_STYLE = { bg: "#0b1530", fg: "#efe5cf" };

/** Reçoit les réservations envoyées par le site. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents || "{}");
    const sheet = getSheet_();
    const now = new Date();
    const yes = String(d.presence) === "Oui";

    // 1) Titulaire de la réservation
    const rows = [[
      now,
      clean_(d.reference),
      clean_(d.prenom),
      clean_(d.nom),
      clean_(d.presence),
      Number(d.passagers) || 0,
      clean_(d.siege),
      clean_(d.telephone),
      // Le site recopie les accompagnants dans le message (sécurité) : on les retire ici
      clean_(String(d.message || "").replace(/\s*\|?\s*Accompagnants : .*$/, "")),
      "Titulaire"
    ]];

    // 2) Une ligne par accompagnant, juste en dessous, avec la même référence
    if (yes) {
      companions_(d).forEach(function (c) {
        rows.push([now, clean_(d.reference), clean_(c.prenom), clean_(c.nom), "Oui", "", "", "", "", "Accompagnant"]);
      });
    }

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    return json_({ ok: true, lignes: rows.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Liste des accompagnants : [{prenom, nom}] envoyée par le site, sinon texte « Prénom Nom, Prénom Nom ». */
function companions_(d) {
  if (Array.isArray(d.accompagnants_liste) && d.accompagnants_liste.length) {
    return d.accompagnants_liste.map(function (c) {
      return { prenom: String(c.prenom || "").trim(), nom: String(c.nom || "").trim() };
    }).filter(function (c) { return c.prenom || c.nom; });
  }
  return String(d.accompagnants || "").split(",").map(function (s) { return s.trim(); })
    .filter(Boolean).map(function (full) {
      const parts = full.split(/\s+/);
      return { prenom: parts.shift(), nom: parts.join(" ") };
    });
}

/**
 * Vérification depuis le navigateur :
 *  - …/exec          → affiche la version si l'application Web répond
 *  - …/exec?test=1   → ajoute une réservation TEST (1 titulaire + 1 accompagnant)
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.test) {
    try {
      const sheet = getSheet_();
      const now = new Date();
      sheet.getRange(sheet.getLastRow() + 1, 1, 2, HEADERS.length).setValues([
        [now, "TEST", "Test", "Titulaire", "Oui", 2, "", "", "Lignes de test, à supprimer", "Titulaire"],
        [now, "TEST", "Test", "Accompagnant", "Oui", "", "", "", "", "Accompagnant"]
      ]);
      return ContentService.createTextOutput(
        "✅ Test réussi : 2 lignes TEST ont été ajoutées dans l'onglet « " + SHEET_NAME +
        " » du fichier « " + sheet.getParent().getName() + " ».");
    } catch (err) {
      return ContentService.createTextOutput("❌ Erreur : " + err);
    }
  }
  return ContentService.createTextOutput("JUNIOR 40 – enregistrement des réservations : OK (" + VERSION + ")");
}

/** À lancer une fois depuis l'éditeur pour créer ou mettre à jour l'onglet. */
function setup() {
  getSheet_();
}

function getSheet_() {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Aucun Google Sheet relié : remplissez SHEET_ID en haut du script.");
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm");
    sheet.getRange("H:H").setNumberFormat("@");
    sheet.getRange("L1").setValue("RÉCAPITULATIF");
  }

  // En-tête : colonne J = « Rôle » (remplace « Accompagnants » des versions précédentes)
  if (sheet.getRange(1, ROLE_COL).getValue() !== HEADERS[ROLE_COL - 1]) {
    sheet.getRange(1, ROLE_COL).setValue(HEADERS[ROLE_COL - 1]);
  }
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight("bold").setBackground(HEAD_STYLE.bg).setFontColor(HEAD_STYLE.fg);

  updateSummary_(sheet);
  return sheet;
}

/** Récapitulatif pour le traiteur : ne compte que les lignes « Titulaire » comme réponses. */
function updateSummary_(sheet) {
  // Retrouve le bloc récapitulatif (colonne K ou L selon la version qui l'a créé)
  const row1 = sheet.getRange(1, 11, 1, 3).getValues()[0];
  let col = 12;
  for (let i = 0; i < row1.length; i++) if (row1[i] === "RÉCAPITULATIF") col = 11 + i;
  if (col <= ROLE_COL) col = 12;

  const notComp = 'J2:J,"<>Accompagnant"';
  const labels = [["RÉCAPITULATIF", ""],
                  ["Réponses reçues", '=COUNTIFS(B2:B,"<>",' + notComp + ')'],
                  ["Réponses « Oui »", '=COUNTIFS(E2:E,"Oui",' + notComp + ')'],
                  ["Réponses « Non »", '=COUNTIFS(E2:E,"Non",' + notComp + ')'],
                  ["Total passagers confirmés", '=SUMIFS(F2:F,E2:E,"Oui",' + notComp + ')'],
                  ["Noms listés (Oui)", '=COUNTIF(E2:E,"Oui")']];
  const current = sheet.getRange(2, col + 1).getFormula();
  if (current === labels[1][1]) return;   // déjà à jour

  labels.forEach(function (l, i) {
    sheet.getRange(1 + i, col).setValue(l[0]);
    if (i > 0) sheet.getRange(1 + i, col + 1).setFormula(l[1]);
  });
  sheet.getRange(1, col).setFontWeight("bold");
  sheet.getRange(5, col, 1, 2).setFontWeight("bold").setBackground("#efe5cf");
  sheet.autoResizeColumn(col);
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

/**
 * JUNIOR 40 – THE JOURNEY
 * Enregistre chaque réservation du site dans ce Google Sheet.
 * Version 5 : une ligne par invité. Si un invité modifie sa réponse, sa ligne est mise à jour
 *             (repérée par son prénom + nom) au lieu d'en ajouter une nouvelle.
 * Version 6 : menu « JMS 40 » dans le Sheet pour préparer l'annonce du lieu sur WhatsApp.
 *
 * Installation : voir google-apps-script/INSTALLATION.md
 */

// Facultatif : l'identifiant du Google Sheet (la partie entre /d/ et /edit dans son adresse).
// À remplir uniquement si le script n'a PAS été créé depuis le Sheet (Extensions → Apps Script).
const SHEET_ID = "";

const SHEET_NAME = "Réservations";
const VERSION = "version 6, mise à jour des réponses + annonce du lieu";
const HEADERS = [
  "Date de réponse", "Référence", "Prénom", "Nom", "Présence",
  "Passagers", "Modifiée le", "Téléphone WhatsApp", "Message", "Accompagnants"
];
const COMP_COL = HEADERS.length;          // colonne J
const MODIF_COL = 7;                      // colonne G (ancienne colonne « Siège »)
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
    const names = yes ? companions_(d).map(function (c) { return (c.prenom + " " + c.nom).trim(); }) : [];

    const row = [
      now,
      clean_(d.reference),
      clean_(d.prenom),
      clean_(d.nom),
      clean_(d.presence),
      yes ? (Number(d.passagers) || 1) : 0,
      "",
      clean_(d.telephone),
      // Le site recopie les accompagnants dans le message (sécurité) : on les retire ici
      clean_(String(d.message || "").replace(/\s*\|?\s*Accompagnants : .*$/, "")),
      clean_(names.join(", "))
    ];

    // L'invité a-t-il déjà répondu ? (même prénom + nom, sans tenir compte des majuscules ni des accents)
    const key = person_(d.prenom, d.nom);
    const matches = findRows_(sheet, key);
    if (key && matches.length) {
      const target = matches[0];
      const first = sheet.getRange(target, 1).getValue();
      row[0] = first || now;               // on garde la date de la première réponse
      row[MODIF_COL - 1] = now;            // et on note la date de la modification
      sheet.getRange(target, 1, 1, HEADERS.length).setValues([row]);
      // Supprime les éventuels doublons plus anciens du même invité (du bas vers le haut)
      for (let i = matches.length - 1; i >= 1; i--) sheet.deleteRow(matches[i]);
      return json_({ ok: true, action: "mise à jour", ligne: target, doublons_supprimes: matches.length - 1 });
    }

    sheet.getRange(sheet.getLastRow() + 1, 1, 1, HEADERS.length).setValues([row]);
    return json_({ ok: true, action: "ajout" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Clé d'identification d'un invité : « prenom nom » en minuscules, sans accents ni tirets. */
function person_(prenom, nom) {
  const s = (String(prenom || "") + " " + String(nom || ""))
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().replace(/[-'’.]/g, " ").replace(/\s+/g, " ").trim();
  return s;
}

/** Numéros des lignes (≥ 2) dont le prénom + nom correspond à la clé, dans l'ordre du tableau. */
function findRows_(sheet, key) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 3, last - 1, 2).getValues();   // colonnes C (Prénom) et D (Nom)
  const out = [];
  values.forEach(function (v, i) { if (person_(v[0], v[1]) === key) out.push(i + 2); });
  return out;
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
 *  - …/exec?test=1   → ajoute une réservation TEST avec un accompagnant
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.test) {
    try {
      const sheet = getSheet_();
      const now = new Date();
      const existing = findRows_(sheet, person_("Test", "Navigateur"));
      const row = [now, "TEST", "Test", "Navigateur", "Oui", 2, existing.length ? now : "", "", "Ligne de test, à supprimer", "Accompagnant Test"];
      if (existing.length) {
        row[0] = sheet.getRange(existing[0], 1).getValue() || now;
        sheet.getRange(existing[0], 1, 1, HEADERS.length).setValues([row]);
      } else {
        sheet.getRange(sheet.getLastRow() + 1, 1, 1, HEADERS.length).setValues([row]);
      }
      return ContentService.createTextOutput(
        (existing.length
          ? "✅ Test réussi : la ligne TEST existante a été MISE À JOUR (colonne « Modifiée le »)"
          : "✅ Test réussi : une ligne TEST a été ajoutée. Rouvrez ce lien : elle sera mise à jour au lieu d'être dupliquée") +
        " — onglet « " + SHEET_NAME + " » du fichier « " + sheet.getParent().getName() + " ».");
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

  // Colonne G : « Siège » (plus utilisée) devient « Modifiée le »
  const g1 = sheet.getRange(1, MODIF_COL).getValue();
  if (g1 !== HEADERS[MODIF_COL - 1]) {
    if (g1 === "Siège" && sheet.getLastRow() > 1) sheet.getRange(2, MODIF_COL, sheet.getLastRow() - 1, 1).clearContent();
    sheet.getRange(1, MODIF_COL).setValue(HEADERS[MODIF_COL - 1]);
    sheet.getRange(1, MODIF_COL, sheet.getMaxRows(), 1).setNumberFormat("dd/mm/yyyy hh:mm");
  }

  // En-tête : colonne J = « Accompagnants »
  if (sheet.getRange(1, COMP_COL).getValue() !== HEADERS[COMP_COL - 1]) {
    sheet.getRange(1, COMP_COL).setValue(HEADERS[COMP_COL - 1]);
  }
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight("bold").setBackground(HEAD_STYLE.bg).setFontColor(HEAD_STYLE.fg);

  updateSummary_(sheet);
  return sheet;
}

/** Récapitulatif pour le traiteur (ignore d'éventuelles anciennes lignes « Accompagnant » de la version 3). */
function updateSummary_(sheet) {
  // Retrouve le bloc récapitulatif (colonne K ou L selon la version qui l'a créé)
  const row1 = sheet.getRange(1, 11, 1, 3).getValues()[0];
  let col = 12;
  for (let i = 0; i < row1.length; i++) if (row1[i] === "RÉCAPITULATIF") col = 11 + i;
  if (col <= COMP_COL) col = 12;

  const notComp = 'J2:J,"<>Accompagnant"';
  const labels = [["RÉCAPITULATIF", ""],
                  ["Réponses reçues", '=COUNTIFS(B2:B,"<>",' + notComp + ')'],
                  ["Réponses « Oui »", '=COUNTIFS(E2:E,"Oui",' + notComp + ')'],
                  ["Réponses « Non »", '=COUNTIFS(E2:E,"Non",' + notComp + ')'],
                  ["Total passagers confirmés", '=SUMIFS(F2:F,E2:E,"Oui",' + notComp + ')']];
  // Retire la ligne « Noms listés » ajoutée par la version 3
  if (sheet.getRange(6, col).getValue() === "Noms listés (Oui)") sheet.getRange(6, col, 1, 2).clearContent();
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


/* =====================================================================
 * MENU « JMS 40 » : annonce du lieu aux invités confirmés
 * ===================================================================== */
const ANNONCE_SHEET = "Annonce du lieu";

/** Ajoute le menu à l'ouverture du Google Sheet. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("JMS 40")
    .addItem("Préparer l'annonce du lieu", "prepareReveal")
    .addSeparator()
    .addItem("Vider l'onglet d'annonce", "clearReveal")
    .addToUi();
}

/** Demande l'adresse puis crée un lien WhatsApp prêt à envoyer pour chaque invité « Oui ». */
function prepareReveal() {
  const ui = SpreadsheetApp.getUi();
  const a = ui.prompt("Annonce du lieu (1/2)",
    "Adresse exacte de la soirée (ex. : Villa Les Palmiers, Rue des Jardins, Cocody, Abidjan) :",
    ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK || !a.getResponseText().trim()) return;
  const address = a.getResponseText().trim();

  const m = ui.prompt("Annonce du lieu (2/2)",
    "Lien Google Maps (facultatif, laissez vide si vous n'en avez pas) :",
    ui.ButtonSet.OK_CANCEL);
  if (m.getSelectedButton() !== ui.Button.OK) return;
  const maps = m.getResponseText().trim();

  const src = getSheet_();
  const last = src.getLastRow();
  if (last < 2) { ui.alert("Aucune réservation pour le moment."); return; }
  const data = src.getRange(2, 1, last - 1, HEADERS.length).getValues();

  const rows = [], skipped = [];
  data.forEach(function (r) {
    const prenom = String(r[2] || "").trim(), nom = String(r[3] || "").trim();
    if (String(r[4]) !== "Oui" || String(r[1]) === "TEST" || !prenom) return;
    const phone = String(r[7] || "").replace(/\D/g, "");
    if (phone.length < 8) { skipped.push(prenom + " " + nom); return; }
    const msg = message_(prenom, address, maps);
    const url = "https://wa.me/" + phone + "?text=" + encodeURIComponent(msg);
    rows.push([prenom, nom, "+" + phone, Number(r[5]) || 1, String(r[9] || ""),
               '=HYPERLINK("' + url.replace(/"/g, '""') + '","Envoyer ✈")', false]);
  });

  const ss = src.getParent();
  let sh = ss.getSheetByName(ANNONCE_SHEET);
  if (!sh) sh = ss.insertSheet(ANNONCE_SHEET);
  sh.clear();
  const head = ["Prénom", "Nom", "Téléphone WhatsApp", "Passagers", "Accompagnants", "Message WhatsApp", "Envoyé ?"];
  sh.getRange(1, 1, 1, head.length).setValues([head])
    .setFontWeight("bold").setBackground(HEAD_STYLE.bg).setFontColor(HEAD_STYLE.fg);
  sh.setFrozenRows(1);
  if (rows.length) {
    // (Apps Script attend toujours la syntaxe de formule anglaise, avec « , »)
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
    sh.getRange(2, 7, rows.length, 1).insertCheckboxes();
    sh.getRange(2, 3, rows.length, 1).setNumberFormat("@");
  }
  sh.getRange(rows.length + 3, 1).setValue("Adresse annoncée : " + address + (maps ? "  ·  " + maps : ""));
  sh.getRange(rows.length + 4, 1).setValue("Mode d'emploi : cliquez sur « Envoyer ✈ », WhatsApp s'ouvre avec le message prêt, appuyez sur Envoyer, puis cochez « Envoyé ? ».");
  sh.autoResizeColumns(1, head.length);
  ss.setActiveSheet(sh);

  ui.alert("Annonce prête",
    rows.length + " invité(s) confirmé(s) dans l'onglet « " + ANNONCE_SHEET + " »." +
    (skipped.length ? "\n\nSans numéro WhatsApp valide (à contacter autrement) : " + skipped.join(", ") : ""),
    ui.ButtonSet.OK);
}

/** Texte du message envoyé à chaque invité. Modifiable librement. */
function message_(prenom, address, maps) {
  return [
    "Bonjour " + prenom + " ✈️",
    "",
    "La porte d'embarquement du vol J40 est enfin révélée !",
    "",
    "📍 " + address,
    maps ? "🗺️ " + maps : "",
    "",
    "Rendez-vous le dimanche 20 décembre 2026, embarquement à 19h00.",
    "Votre présence est le plus beau des cadeaux.",
    "",
    "JMS Airways"
  ].filter(function (l, i, arr) { return l !== "" || arr[i - 1] !== ""; }).join("\n");
}

/** Vide l'onglet d'annonce (par exemple pour recommencer avec une autre adresse). */
function clearReveal() {
  const sh = getSheet_().getParent().getSheetByName(ANNONCE_SHEET);
  if (sh) sh.clear();
  SpreadsheetApp.getUi().alert("L'onglet « " + ANNONCE_SHEET + " » a été vidé.");
}

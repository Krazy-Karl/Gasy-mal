// =====================================================================
// Gasy-mal : script Apps Script lié au Google Sheets
// Une ligne = les points d'un joueur sur une partie, à une date.
// La saison (un mois, ex. 2026-11) est calculée automatiquement.
// Un joueur créé sans points est enregistré avec 0 point et sans partie.
// =====================================================================

const NOM_FEUILLE = "Points";
const ENTETES = ["ID", "Date", "Saison", "Partie", "Joueur", "Points"];
const LONGUEUR_MAX = 50;      // taille maximale d'un nom
const POINTS_MAX = 1000;      // valeur absolue maximale d'un ajout

// À lancer UNE FOIS à la main (menu "Exécuter") : prépare la feuille.
function initialiser() {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();
  let f = classeur.getSheetByName(NOM_FEUILLE);
  if (!f) f = classeur.insertSheet(NOM_FEUILLE);
  f.getRange(1, 1, 1, ENTETES.length).setValues([ENTETES]).setFontWeight("bold");
  f.getRange("C:C").setNumberFormat("@"); // "2026-11" reste du texte
  f.setFrozenRows(1);
}

// LECTURE : renvoie toutes les lignes
function doGet(e) {
  try {
    return reponse({ ok: true, lignes: lireLignes() });
  } catch (err) {
    return reponse({ ok: false, erreur: "Erreur du script : " + err.message });
  }
}

// ÉCRITURE : ajouter des points, créer un joueur ou supprimer une ligne
function doPost(e) {
  const verrou = LockService.getScriptLock();
  let verrouPris = false;
  try {
    verrou.waitLock(10000);
    verrouPris = true;
    const d = JSON.parse(e.postData.contents);

    if (d.action === "ajouter") {
      const partie = nettoyer(d.partie);   // facultatif
      const joueur = nettoyer(d.joueur);
      const points = Number(d.points);
      if (!joueur) return reponse({ ok: false, erreur: "Le nom du joueur est obligatoire." });
      if (!isFinite(points) || points === 0 || Math.abs(points) > POINTS_MAX) {
        return reponse({ ok: false, erreur: "Les points doivent être un nombre non nul (entre -" + POINTS_MAX + " et " + POINTS_MAX + ")." });
      }
      ajouterLigne(partie, joueur, points);
      return reponse({ ok: true, lignes: lireLignes() });
    }

    if (d.action === "creerJoueur") {
      const joueur = nettoyer(d.joueur);
      if (!joueur) return reponse({ ok: false, erreur: "Le nom du joueur est obligatoire." });
      const existe = lireLignes().some(l => l.joueur.toLowerCase() === joueur.toLowerCase());
      if (existe) return reponse({ ok: false, erreur: "Ce joueur existe déjà." });
      ajouterLigne("", joueur, 0);
      return reponse({ ok: true, lignes: lireLignes() });
    }

    if (d.action === "supprimer") {
      const f = feuille();
      const n = f.getLastRow();
      const ids = n < 2 ? [] : f.getRange(2, 1, n - 1, 1).getValues().map(l => String(l[0]));
      const i = ids.indexOf(String(d.id));
      if (i === -1) return reponse({ ok: false, erreur: "Cette ligne n'existe plus (déjà supprimée ?)." });
      f.deleteRow(i + 2);
      return reponse({ ok: true, lignes: lireLignes() });
    }

    return reponse({ ok: false, erreur: "Action inconnue." });
  } catch (err) {
    return reponse({ ok: false, erreur: "Erreur du script : " + err.message });
  } finally {
    if (verrouPris) verrou.releaseLock();
  }
}

// ---------- Fonctions utilitaires ----------

function ajouterLigne(partie, joueur, points) {
  const maintenant = new Date();
  const tz = Session.getScriptTimeZone();
  feuille().appendRow([
    Utilities.getUuid(),
    Utilities.formatDate(maintenant, tz, "yyyy-MM-dd HH:mm"),
    Utilities.formatDate(maintenant, tz, "yyyy-MM"),
    partie, joueur, points
  ]);
}

function feuille() {
  const f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  if (!f) throw new Error("La feuille \"" + NOM_FEUILLE + "\" est introuvable. Lancez la fonction initialiser().");
  return f;
}

function lireLignes() {
  const f = feuille();
  const n = f.getLastRow();
  if (n < 2) return [];
  const tz = Session.getScriptTimeZone();
  return f.getRange(2, 1, n - 1, ENTETES.length).getValues()
    .map(l => ({
      id: String(l[0]),
      date: l[1] instanceof Date ? Utilities.formatDate(l[1], tz, "yyyy-MM-dd HH:mm") : String(l[1]),
      saison: l[2] instanceof Date ? Utilities.formatDate(l[2], tz, "yyyy-MM") : String(l[2]),
      partie: String(l[3]),
      joueur: String(l[4]),
      points: Number(l[5])
    }))
    .filter(l => l.id);
}

// Enlève les espaces et les caractères qui feraient passer le texte pour une formule Excel/Sheets
function nettoyer(texte) {
  return String(texte || "").replace(/^[=+\-@\s]+/, "").trim().slice(0, LONGUEUR_MAX);
}

function reponse(objet) {
  return ContentService
    .createTextOutput(JSON.stringify(objet))
    .setMimeType(ContentService.MimeType.JSON);
}

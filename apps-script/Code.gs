// =====================================================================
// Gasy-mal : script Apps Script lié au Google Sheets
// Une ligne = les points d'un joueur, à une date.
// La saison (un mois, ex. 2026-11) est calculée automatiquement.
// Un joueur créé sans points est enregistré avec 0 point.
// =====================================================================

const NOM_FEUILLE = "Points";
const ENTETES = ["ID", "Date", "Saison", "Joueur", "Points"];
const LONGUEUR_MAX = 50;      // taille maximale d'un nom
const POINTS_MAX = 1000;      // valeur absolue maximale d'un ajout

// À lancer UNE FOIS à la main (menu "Exécuter") : prépare la feuille.
function initialiser() {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();
  let f = classeur.getSheetByName(NOM_FEUILLE);
  if (!f) f = classeur.insertSheet(NOM_FEUILLE);
  migrer(f);
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
    migrer(feuille()); // retire l'ancienne colonne "Partie" si elle existe encore
    const d = JSON.parse(e.postData.contents);

    if (d.action === "ajouter") {
      const joueur = nettoyer(d.joueur);
      const points = Number(d.points);
      if (!joueur) return reponse({ ok: false, erreur: "Le nom du joueur est obligatoire." });
      if (!isFinite(points) || points === 0 || Math.abs(points) > POINTS_MAX) {
        return reponse({ ok: false, erreur: "Les points doivent être un nombre non nul (entre -" + POINTS_MAX + " et " + POINTS_MAX + ")." });
      }
      // On renvoie seulement la ligne écrite (plus rapide que de relire toute la feuille)
      return reponse({ ok: true, ligne: ajouterLigne(joueur, points) });
    }

    if (d.action === "ajouterPlusieurs") { // fin de partie : tous les points d'un coup, en une seule écriture
      const gains = Array.isArray(d.gains) ? d.gains : [];
      if (!gains.length || gains.length > 50) return reponse({ ok: false, erreur: "Liste de points invalide." });
      const maintenant = new Date();
      const tz = Session.getScriptTimeZone();
      const date = Utilities.formatDate(maintenant, tz, "yyyy-MM-dd HH:mm");
      const saison = Utilities.formatDate(maintenant, tz, "yyyy-MM");
      const lignes = [], ajoutees = [];
      for (let i = 0; i < gains.length; i++) {
        const joueur = nettoyer(gains[i].joueur);
        const points = Number(gains[i].points);
        if (!joueur) return reponse({ ok: false, erreur: "Le nom du joueur est obligatoire." });
        if (!isFinite(points) || points === 0 || Math.abs(points) > POINTS_MAX) {
          return reponse({ ok: false, erreur: "Points invalides pour " + joueur + "." });
        }
        const id = Utilities.getUuid();
        lignes.push([id, date, saison, joueur, points]);
        ajoutees.push({ id: id, date: date, saison: saison, joueur: joueur, points: points });
      }
      const f = feuille();
      const debut = f.getLastRow() + 1;
      f.getRange(debut, 3, lignes.length, 1).setNumberFormat("@"); // "2026-11" reste du texte
      f.getRange(debut, 1, lignes.length, 5).setValues(lignes);
      return reponse({ ok: true, ajoutees: ajoutees });
    }

    if (d.action === "creerJoueur") {
      const joueur = nettoyer(d.joueur);
      if (!joueur) return reponse({ ok: false, erreur: "Le nom du joueur est obligatoire." });
      if (joueurExiste(joueur)) return reponse({ ok: false, erreur: "Ce joueur existe déjà." });
      return reponse({ ok: true, ligne: ajouterLigne(joueur, 0) });
    }

    if (d.action === "renommer") {
      const ancien = nettoyer(d.ancien), nouveau = nettoyer(d.nouveau);
      if (!ancien || !nouveau) return reponse({ ok: false, erreur: "Le nom du joueur est obligatoire." });
      const f = feuille();
      const n = f.getLastRow();
      if (n < 2) return reponse({ ok: false, erreur: "Joueur introuvable." });
      const plage = f.getRange(2, ancienneDisposition(f) ? 5 : 4, n - 1, 1);
      const noms = plage.getValues();
      const basAncien = ancien.toLowerCase(), basNouveau = nouveau.toLowerCase();
      let trouves = 0, conflit = false;
      noms.forEach(function (l) {
        const v = String(l[0]).toLowerCase();
        if (v === basAncien) trouves++;
        else if (v === basNouveau) conflit = true;
      });
      if (!trouves) return reponse({ ok: false, erreur: "Joueur introuvable (déjà renommé ou supprimé ?)." });
      if (conflit) return reponse({ ok: false, erreur: "Ce nom est déjà utilisé par un autre joueur." });
      plage.setValues(noms.map(l => String(l[0]).toLowerCase() === basAncien ? [nouveau] : l));
      return reponse({ ok: true, ancien: ancien, nouveau: nouveau });
    }

    if (d.action === "supprimer") {
      const f = feuille();
      const n = f.getLastRow();
      const ids = n < 2 ? [] : f.getRange(2, 1, n - 1, 1).getValues().map(l => String(l[0]));
      const i = ids.indexOf(String(d.id));
      if (i === -1) return reponse({ ok: false, erreur: "Cette ligne n'existe plus (déjà supprimée ?)." });
      f.deleteRow(i + 2);
      return reponse({ ok: true, id: String(d.id) });
    }

    return reponse({ ok: false, erreur: "Action inconnue." });
  } catch (err) {
    return reponse({ ok: false, erreur: "Erreur du script : " + err.message });
  } finally {
    if (verrouPris) verrou.releaseLock();
  }
}

// ---------- Fonctions utilitaires ----------

// Écrit une ligne et renvoie ce qui a été écrit (sans relire la feuille)
function ajouterLigne(joueur, points) {
  const maintenant = new Date();
  const tz = Session.getScriptTimeZone();
  const ligne = {
    id: Utilities.getUuid(),
    date: Utilities.formatDate(maintenant, tz, "yyyy-MM-dd HH:mm"),
    saison: Utilities.formatDate(maintenant, tz, "yyyy-MM"),
    joueur: joueur,
    points: points
  };
  feuille().appendRow([ligne.id, ligne.date, ligne.saison, ligne.joueur, ligne.points]);
  return ligne;
}

// Lit seulement la colonne des noms (rapide)
function joueurExiste(nom) {
  const f = feuille();
  const n = f.getLastRow();
  if (n < 2) return false;
  const col = ancienneDisposition(f) ? 5 : 4;
  const bas = nom.toLowerCase();
  return f.getRange(2, col, n - 1, 1).getValues().some(l => String(l[0]).toLowerCase() === bas);
}

let _feuille = null; // mémorisée pendant une même exécution
function feuille() {
  if (_feuille) return _feuille;
  const f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  if (!f) throw new Error("La feuille \"" + NOM_FEUILLE + "\" est introuvable. Lancez la fonction initialiser().");
  _feuille = f;
  return f;
}

// Ancienne feuille : colonnes ID, Date, Saison, Partie, Joueur, Points -> on retire "Partie"
function ancienneDisposition(f) {
  return f.getLastColumn() >= 4 && String(f.getRange(1, 4).getValue()) === "Partie";
}
function migrer(f) {
  if (ancienneDisposition(f)) f.deleteColumn(4);
}

function lireLignes() {
  const f = feuille();
  const n = f.getLastRow();
  if (n < 2) return [];
  const tz = Session.getScriptTimeZone();
  const ancien = ancienneDisposition(f); // lecture correcte même avant la migration
  const iJoueur = ancien ? 4 : 3, iPoints = ancien ? 5 : 4;
  return f.getRange(2, 1, n - 1, ancien ? 6 : 5).getValues()
    .map(l => ({
      id: String(l[0]),
      date: l[1] instanceof Date ? Utilities.formatDate(l[1], tz, "yyyy-MM-dd HH:mm") : String(l[1]),
      saison: l[2] instanceof Date ? Utilities.formatDate(l[2], tz, "yyyy-MM") : String(l[2]),
      joueur: String(l[iJoueur]),
      points: Number(l[iPoints])
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

// =====================================================================
// IMPORT UNIQUE : joueurs de la Saison 12 (novembre 2026)
// À lancer UNE FOIS à la main (menu "Exécuter"). Sans risque de doublon :
// un joueur déjà présent dans la feuille est ignoré.
// Mettez AVEC_POINTS à false pour créer les joueurs à 0 point.
// Vous pourrez supprimer cette fonction ensuite.
// =====================================================================
function importerSaison12() {
  const AVEC_POINTS = true;
  const JOUEURS = [
    ["Joel", 24], ["Diamondra", 23], ["Tsiory", 19], ["Hannah", 18],
    ["Saraha", 12], ["Koloina", 12], ["Tiffany", 3], ["Karl", 1]
  ];
  migrer(feuille());
  const existants = lireLignes().map(l => l.joueur.toLowerCase());
  let ajoutes = 0, ignores = [];
  JOUEURS.forEach(function (j) {
    if (existants.indexOf(j[0].toLowerCase()) !== -1) { ignores.push(j[0]); return; }
    ajouterLigne(j[0], AVEC_POINTS ? j[1] : 0);
    ajoutes++;
  });
  Logger.log(ajoutes + " joueur(s) ajouté(s)." + (ignores.length ? " Déjà présents, ignorés : " + ignores.join(", ") : ""));
}

// ÉTAPE 1 : le script répond simplement "ok" quand la page l'appelle.
// Il sera complété aux étapes suivantes (lecture, écriture).

function doGet(e) {
  return reponse({ ok: true, message: "Le script répond bien !" });
}

function reponse(objet) {
  return ContentService
    .createTextOutput(JSON.stringify(objet))
    .setMimeType(ContentService.MimeType.JSON);
}

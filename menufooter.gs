function showMenuFooter() {
  // Étape 1 : Création de la modale
  const html = HtmlService.createHtmlOutputFromFile('MenuFooter')
    .setWidth(1200)
    .setHeight(800); // Ajuste selon besoin

  // Étape 2 : Affichage dans l'UI
  SpreadsheetApp.getUi().showModalDialog(html, 'Créer Menu & Footer');
  Logger.log('[showMenuFooter][1] Fenêtre "Créer Menu & Footer" affichée (900x700)');
}

function getArborescenceForMenuFooter() {
  Logger.log("[MenuFooter][1] Début extraction Arborescence");

  // [1.1] Ouverture du classeur et de la feuille Arborescence
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = ss.getSheetByName("Arborescence");
  if (!feuille) {
    Logger.log("[MenuFooter][1.2] Feuille Arborescence introuvable");
    return [];
  }

  // [1.2] Lecture des données (hors en-têtes)
  const props = PropertiesService.getDocumentProperties();
  const urlPreprod = props.getProperty('urlPreprod') || '';

  const nbLignes = feuille.getLastRow() - 1;
  const nbColonnes = feuille.getLastColumn();
  if (nbLignes <= 0) {
    Logger.log("[MenuFooter][1.3] Aucune donnée dans Arborescence");
    return [];
  }
  const data = feuille.getRange(2, 1, nbLignes, nbColonnes).getValues();
  Logger.log("[MenuFooter][1.4] " + nbLignes + " lignes lues dans Arborescence");

  // [1.3] Lecture des en-têtes pour mapping colonnes
  const headers = feuille.getRange(1, 1, 1, nbColonnes).getValues()[0];

  // [1.4] Construction des objets pages
  const pages = data.map(row => {
    const obj = {};
    headers.forEach((col, i) => obj[col] = row[i]);
    // Ajout d'une clé urlRelative, qui retire urlPreprod du début de l'URL si présent
    if (obj.URL && urlPreprod && obj.URL.startsWith(urlPreprod)) {
      obj.urlRelative = obj.URL.substring(urlPreprod.length) || '/';
    } else {
      obj.urlRelative = obj.URL || '';
    }
    return obj;
  });


  Logger.log("[MenuFooter][2] Objets pages créés : " + pages.length);

  // [1.5] Construction de la hiérarchie via Niveau/Niveau parent
  // (optionnel : si tu veux proposer un menu pré-hiérarchisé dans l’UI)
  // Ici, on garde la structure à plat : chaque page avec toutes ses infos, pour filtrer dans l’UI

  Logger.log("[MenuFooter][3] Extraction terminée, retour à la fenêtre HTML UX");

  return pages;
}

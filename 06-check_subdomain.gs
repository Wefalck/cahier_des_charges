function check_subdomain() {
  const html = HtmlService.createHtmlOutputFromFile('CheckSubdomain')
    .setWidth(600)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Vérification des sous-domaines');
}

function obtenirDomaineNouveau() {
  // 1. Récupérer la propriété 'urlNouveau' depuis les propriétés du document
  const props = PropertiesService.getDocumentProperties();
  const urlNouveau = props.getProperty('urlNouveau');
  Logger.log("[obtenirDomaineNouveau] Valeur brute urlNouveau : " + urlNouveau);

  // 2. Utiliser la fonction extractDomain pour extraire le domaine racine
  if (urlNouveau && typeof extractDomain === "function") {
    const domaine = extractDomain(urlNouveau);
    Logger.log("[obtenirDomaineNouveau] Domaine extrait : " + domaine);
    return domaine || '';
  } else {
    Logger.log("[obtenirDomaineNouveau] urlNouveau non défini ou extractDomain non dispo");
    return '';
  }
}

/**
 * Étape 1 : Créer ou remplacer la feuille "Sous-domaine" avec formatage et injection des sous-domaines
 * Nom de la fonction : creerFeuilleSousDomaine
 * @param {string} data - Sous-domaines reçus du formulaire (un par ligne)
 */
function creerFeuilleSousDomaine(data) {
  Logger.log("[creerFeuilleSousDomaine] Début - données reçues :\n" + data);

  // 1. Vérification des données reçues
  if (!data || typeof data !== "string") {
    Logger.log("[creerFeuilleSousDomaine] Données invalides !");
    return "❌ Aucune donnée à traiter.";
  }

  // 2. Extraction et nettoyage des sous-domaines (un par ligne, sans doublons, sans vide)
  const sousDomaines = data
    .split('\n')
    .map(s => s.trim())
    .filter(s => s)
    .filter((s, idx, arr) => arr.indexOf(s) === idx); // unique

  if (sousDomaines.length === 0) {
    Logger.log("[creerFeuilleSousDomaine] Aucun sous-domaine valide trouvé.");
    return "❌ Aucun sous-domaine valide.";
  }
  Logger.log(`[creerFeuilleSousDomaine] ${sousDomaines.length} sous-domaines à insérer.`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 3. Création ou suppression + création de la feuille "Sous-domaine"
  let feuille = ss.getSheetByName("Sous-domaine");
  if (feuille) {
    ss.deleteSheet(feuille);
    Logger.log("[creerFeuilleSousDomaine] Ancienne feuille 'Sous-domaine' supprimée.");
  }
  feuille = ss.insertSheet("Sous-domaine");
  feuille.setTabColor("#2980b9");
  Logger.log("[creerFeuilleSousDomaine] Nouvelle feuille 'Sous-domaine' créée et colorée.");

  // 4. Définition des en-têtes (ligne 1)
  const entetes = ["Sous-domaines", "Site :", "URL indexées", "Commentaire"];
  feuille.getRange(1, 1, 1, entetes.length)
    .setValues([entetes])
    .setFontWeight("bold")
    .setFontColor("white")
    .setFontFamily("Arial")
    .setFontSize(10)
    .setVerticalAlignment("middle");
  Logger.log("[creerFeuilleSousDomaine] En-têtes ajoutés.");

  // 5. Construction des lignes à insérer
  // Col A : sous-domaine
  // Col B : lien Google "site:"
  // Col C : vide (format nombre)
  // Col D : vide
  const lignes = sousDomaines.map(sd => [
    sd,
    `https://www.google.com/search?q=site%3A${encodeURIComponent(sd)}`,
    "", // Colonne C
    ""  // Colonne D
  ]);

  feuille.getRange(2, 1, lignes.length, 4)
    .setValues(lignes)
    .setFontColor("black")
    .setFontSize(10)
    .setFontFamily("Arial")
    .setVerticalAlignment("middle");
  Logger.log("[creerFeuilleSousDomaine] Données sous-domaines insérées.");

  // 6. Ajout de liens hypertexte en colonne B
  const liens = lignes.map(row => [
    SpreadsheetApp.newRichTextValue()
      .setText("site:" + row[0])
      .setLinkUrl(row[1])
      .build()
  ]);
  feuille.getRange(2, 2, lignes.length, 1).setRichTextValues(liens);
  Logger.log("[creerFeuilleSousDomaine] Liens Google ajoutés en colonne B.");

  // 7. Formatage colonne C (nombre, séparateur milliers, pas de décimales)
  feuille.getRange(2, 3, lignes.length, 1)
    .setNumberFormat("#,##0")
    .setHorizontalAlignment("left");

  // 8. Masquage du quadrillage
  feuille.setHiddenGridlines(true);

  // 9. Suppression des colonnes E à Z
  const lastCol = feuille.getMaxColumns();
  if (lastCol > 4) {
    feuille.deleteColumns(5, lastCol - 4);
    Logger.log("[creerFeuilleSousDomaine] Colonnes E à Z supprimées.");
  }

  // 10. Nettoyage des lignes vides (colonne A)
  const lastRow = feuille.getLastRow();
  const valeursColA = feuille.getRange(2, 1, lastRow - 1).getValues().flat();
  const firstEmpty = valeursColA.findIndex(v => v.trim?.() === "");
  let rowASupprimer;
  if (firstEmpty !== -1) {
    rowASupprimer = firstEmpty + 2;
  } else {
    rowASupprimer = lastRow + 1;
  }
  const totalRows = feuille.getMaxRows();
  const aSupprimer = totalRows - rowASupprimer + 1;
  if (aSupprimer > 0) {
    feuille.deleteRows(rowASupprimer, aSupprimer);
    Logger.log(`[creerFeuilleSousDomaine] ${aSupprimer} lignes vides supprimées à partir de la ligne ${rowASupprimer}.`);
  }

  // 11. Banding
  const nbLignes = feuille.getLastRow();
  feuille.getBandings().forEach(b => b.remove());
  const plageBanding = feuille.getRange(1, 1, nbLignes, 4);
  plageBanding
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor("#073763")
    .setFirstRowColor("#FFFFFF")
    .setSecondRowColor("#F3F3F3");
  Logger.log("[creerFeuilleSousDomaine] Banding appliqué.");

  // 12. Figer la première ligne + filtre
  feuille.setFrozenRows(1);
  feuille.getRange(1, 1, 1, 4).createFilter();
  Logger.log("[creerFeuilleSousDomaine] Première ligne figée + filtre activé.");

  // 13. Largeur colonnes
  feuille.setColumnWidth(1, 200); // A
  feuille.setColumnWidth(2, 200); // B
  feuille.setColumnWidth(3, 150); // C
  feuille.setColumnWidth(4, 600); // D
  Logger.log("[creerFeuilleSousDomaine] Largeur des colonnes appliquée.");

  // 14. Ajout d’un lien vers l’onglet "Sous-domaine" dans la feuille "Suivi", colonne F, à chaque ligne où "Sous-domaine" est trouvé en colonne B
  try {
    const feuilleSuivi = ss.getSheetByName("Suivi");
    if (feuilleSuivi) {
      const lastRowSuivi = feuilleSuivi.getLastRow();
      const valeursColB = feuilleSuivi.getRange(1, 2, lastRowSuivi).getValues().flat();

      // Récupérer le GID de la feuille "Sous-domaine"
      const gidSousDomaine = feuille.getSheetId();

      let liensAjoutes = 0;

      for (let i = 0; i < valeursColB.length; i++) {
        if (typeof valeursColB[i] === "string" && valeursColB[i].trim().toLowerCase() === "sous-domaine") {
          // Met le lien dans la colonne F (colonne 6), ligne i+1
          const formuleLien = `=HYPERLINK("#gid=${gidSousDomaine}";"Sous-domaine")`;
          feuilleSuivi.getRange(i + 1, 6).setFormula(formuleLien);
          liensAjoutes++;
          Logger.log(`[creerFeuilleSousDomaine] Lien ajouté dans Suivi!F${i + 1}`);
        }
      }
      Logger.log(`[creerFeuilleSousDomaine] ${liensAjoutes} lien(s) vers "Sous-domaine" ajoutés dans Suivi!F`);
    } else {
      Logger.log("[creerFeuilleSousDomaine] Feuille Suivi non trouvée.");
    }
  } catch (e) {
    Logger.log("[creerFeuilleSousDomaine] Erreur lors de l’ajout du lien dans Suivi : " + e.message);
  }

  reordonnerFeuillesVisibles();

  Logger.log("[creerFeuilleSousDomaine] Feuille 'Sous-domaine' terminée !");
  return "success";
}


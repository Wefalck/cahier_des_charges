function afficherTemplateModal() {
  // Étape 1 - Création de la modale avec le fichier HTML Template.html
  const html = HtmlService.createHtmlOutputFromFile('Template')
    .setWidth(600)
    .setHeight(800);

  // Étape 2 - Affichage dans l’UI
  SpreadsheetApp.getUi().showModalDialog(html, 'Définir les templates');
  Logger.log('[INFO] Fenêtre "Définir les templates" affichée (600x800)');
}

function enregistrerTemplates(data) {
  if (!Array.isArray(data)) {
    Logger.log('[ERREUR] Données invalides : data n’est pas un tableau.');
    return;
  }

  // Étape 1 – Nettoyage et structure des templates
  const templates = data
    .map(t => ({
      nom: (t.nom || '').trim(),
      url: (t.url || '').trim()
    }))
    .filter(t => t.nom);

  Logger.log('[INFO] Templates valides : ' + JSON.stringify(templates));

  // Étape 2 – Enregistrement dans les propriétés du document
  PropertiesService.getDocumentProperties().setProperty('templates', JSON.stringify(templates));

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Étape 3 – Écriture des noms dans Config!B69:B79
  const feuilleConfig = ss.getSheetByName('Config');
  if (feuilleConfig) {
    const noms = templates.slice(0, 11).map(t => [t.nom]);
    const plageConfig = feuilleConfig.getRange(69, 2, noms.length, 1);
    plageConfig.setValues(noms);
    plageConfig.setFontFamily('Arial');
    plageConfig.setFontSize(10);
    plageConfig.setFontWeight('bold');
    plageConfig.setFontColor('white');
    Logger.log(`[INFO] ${noms.length} noms écrits dans Config!B69:B79.`);
  }

  // Étape 4 – Mise à jour de la feuille "Balisage template"
  const feuille = ss.getSheetByName('Balisage template');
  if (!feuille) {
    Logger.log('[ERREUR] Feuille "Balisage template" introuvable.');
    return;
  }

  // Affichage si la feuille est masquée
  if (feuille.isSheetHidden()) {
    feuille.showSheet();
    Logger.log('[INFO] Feuille "Balisage template" affichée.');
  }
  ss.setActiveSheet(feuille);

  const startRow = 15;
  const noms = templates.map(t => [t.nom]);
  const urls = templates.map(t => [t.url]);

  // Écriture des noms A15:A
  const plageNom = feuille.getRange(startRow, 1, noms.length, 1);
  plageNom.setValues(noms);
  plageNom.setFontFamily('Arial');
  plageNom.setFontSize(11);
  plageNom.setFontColor('black');
  plageNom.setVerticalAlignment('middle');

  // Écriture des URL B15:B
  const plageUrl = feuille.getRange(startRow, 2, urls.length, 1);
  plageUrl.setValues(urls);
  plageUrl.setFontFamily('Arial');
  plageUrl.setFontSize(11);
  plageUrl.setFontColor('black');
  plageUrl.setShowHyperlink(true);
  plageUrl.setVerticalAlignment('middle');
  plageUrl.setRichTextValues(
    urls.map(([url]) => {
      return [url.includes('drive.google.com')
        ? SpreadsheetApp.newRichTextValue().setText(url).setLinkUrl(url).build()
        : SpreadsheetApp.newRichTextValue().setText(url).build()];
    })
  );

  Logger.log(`[INFO] Données écrites dans Balisage template!A15:A et B15:B.`);

  // -------------------------------------------------------------------
  // Étape 5 – Ajout du lien dans Suivi vers "Balisage template"
  try {
    Logger.log('[INFO] Début ajout du lien vers "Balisage template" dans la feuille Suivi.');

    // 5.1. Récupération de la feuille Suivi
    const feuilleSuivi = ss.getSheetByName("Suivi");
    if (!feuilleSuivi) {
      Logger.log('[WARN] Feuille "Suivi" introuvable, lien non ajouté.');
      return;
    }

    // 5.2. Recherche de toutes les lignes où la colonne B contient "Template" (insensible à la casse/espaces)
    const lastRowSuivi = feuilleSuivi.getLastRow();
    const valeursColB = feuilleSuivi.getRange(1, 2, lastRowSuivi).getValues().flat();
    Logger.log(`[DEBUG] ${valeursColB.length} valeurs lues dans Suivi!B.`);

    // 5.3. Récupération du sheetId (gid) de la feuille Balisage template
    const gidBalisageTemplate = feuille.getSheetId();
    Logger.log(`[DEBUG] gid Balisage template = ${gidBalisageTemplate}`);

    let liensAjoutes = 0;

    // 5.4. Boucle sur toutes les lignes
    for (let i = 0; i < valeursColB.length; i++) {
      if (
        typeof valeursColB[i] === "string" &&
        valeursColB[i].trim().toLowerCase() === "template"
      ) {
        // 5.5. Formule du lien à insérer
        const formuleLien = `=HYPERLINK("#gid=${gidBalisageTemplate}";"Balisage template")`;
        feuilleSuivi.getRange(i + 1, 6).setFormula(formuleLien); // Colonne F (6)
        Logger.log(`[INFO] Lien ajouté dans Suivi!F${i + 1}`);
        liensAjoutes++;
      }
    }
    Logger.log(`[INFO] ${liensAjoutes} lien(s) vers "Balisage template" ajoutés dans Suivi!F.`);
  } catch (e) {
    Logger.log(`[ERREUR] lors de l’ajout du lien dans Suivi : ${e.message}`);
  }
}


function getTemplates() {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty('templates');

  Logger.log('[INFO] Contenu brut de la propriété "templates" : %s', raw);

  if (!raw) {
    Logger.log('[INFO] Aucune donnée trouvée pour "templates".');
    return [];
  }

  try {
    const templates = JSON.parse(raw);

    if (!Array.isArray(templates)) {
      Logger.log('[ERREUR] Le JSON parsé n’est pas un tableau : %s', JSON.stringify(templates));
      return [];
    }

    Logger.log('[INFO] Templates récupérés avec succès : %s', JSON.stringify(templates));
    return templates;

  } catch (e) {
    Logger.log('[ERREUR] JSON.parse échoué pour "templates" : %s', e.message);
    return [];
  }
}


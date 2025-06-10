function get_data() {
  const html = HtmlService.createHtmlOutputFromFile('GetData')
    .setWidth(900)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Récupérer des données');
}

function gererImportCsv(csvString, type) {
  const donnees = parseCSV(csvString);
  switch (type) {
    case 'crawl_prod':
      return importerDonneesCrawlProd(donnees);
    case 'semrush':
      return importerDonneesSemrush(donnees);
    default:
      throw new Error('❌ Type non reconnu : ' + type);
  }
}

function importerDonneesCrawlProd(donnees) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1️⃣ Création ou remplacement de la feuille "Inventaire"
  let feuille = ss.getSheetByName("Inventaire");
  if (feuille) ss.deleteSheet(feuille);
  feuille = ss.insertSheet("Inventaire");
  feuille.setTabColor("#f39c12");
  Logger.log("📄 Feuille 'Crawl - Prod' (re)créée avec couleur #f39c12");

  // 2️⃣ Récupération des en-têtes et index des colonnes utiles
  const enTetes = donnees[0];
  const indexAdresse = enTetes.indexOf("Adresse");
  const indexSegment = enTetes.indexOf("Segments");
  const indexCodeHTTP = enTetes.indexOf("Code HTTP");
  const indexTitle = enTetes.indexOf("Title 1");
  const indexH1 = enTetes.indexOf("H1-1");
  const indexMetaDesc = enTetes.indexOf("Meta Description 1"); // 🆕

  if ([indexAdresse, indexSegment, indexCodeHTTP, indexTitle, indexH1, indexMetaDesc].includes(-1)) {
    throw new Error("❌ Colonnes essentielles manquantes dans le CSV.");
  }

  // 3️⃣ Filtrage : lignes avec Code HTTP = 200 ET pas d'URL avec paramètre ou 'page'
  const donneesFiltrees = donnees.slice(1).filter(ligne =>
    ligne[indexCodeHTTP] === "200" &&
    !ligne[indexAdresse].includes("?") &&
    !/page(=|\/)/i.test(ligne[indexAdresse])
  );
  Logger.log(`🔎 ${donneesFiltrees.length} lignes conservées après filtre paramètres/page`);

  // 4️⃣ Construction des lignes à insérer
  // [Template, URL, Conserver ?, Nouvelle URL, Mot-clé visé, Title, <h1>, Meta description]
  const lignes = donneesFiltrees.map(ligne => [
    ligne[indexSegment] || "",         // A : Template
    ligne[indexAdresse] || "",         // B : URL
    false,                             // C : Conserver ? (case à cocher)
    "",                                // D : Nouvelle URL
    "",                                // E : Mot-clé visé
    ligne[indexTitle] || "",           // F : Title
    ligne[indexH1] || "",              // G : <h1>
    ligne[indexMetaDesc] || ""         // H : Meta description (🆕)
  ]);

  // 5️⃣ En-têtes + insertion
  const entetesFinales = [
    "Template", "URL", "Conserver ?", "Nouvelle URL", "Mot-clé visé", "Title", "<h1>", "Meta description"
  ];
  feuille.getRange(1, 1, 1, entetesFinales.length).setValues([entetesFinales]);
  feuille.getRange(2, 1, lignes.length, entetesFinales.length).setValues(lignes);
  Logger.log("📥 Données insérées avec en-têtes en ligne 1");

  // 6️⃣ Formatage : en-têtes
  feuille.getRange(1, 1, 1, entetesFinales.length)
    .setFontWeight("bold")
    .setFontColor("white")
    .setFontFamily("Arial")
    .setFontSize(10)
    .setVerticalAlignment("middle");

  // 7️⃣ Formatage : contenu
  feuille.getRange(2, 1, lignes.length, entetesFinales.length)
    .setFontColor("black")
    .setFontSize(10)
    .setFontFamily("Arial")
    .setVerticalAlignment("middle");

  // 8️⃣ Ajout des cases à cocher (colonne C = 3)
  feuille.getRange(2, 3, lignes.length).insertCheckboxes();
  Logger.log("☑️ Cases à cocher insérées en colonne C");

  // 9️⃣ Masquage du quadrillage
  feuille.setHiddenGridlines(true);
  Logger.log("🔲 Quadrillage désactivé");

  // 🔟 Suppression des colonnes inutiles (I à Z)
  const lastCol = feuille.getMaxColumns();
  if (lastCol > 8) {
    feuille.deleteColumns(9, lastCol - 8);
    Logger.log("🗑️ Colonnes I à Z supprimées");
  }

  // 11️⃣ Nettoyage des lignes vides
  let lastRow = feuille.getLastRow();
  const valeursColB = feuille.getRange(2, 2, lastRow - 1).getValues().flat();
  const firstEmpty = valeursColB.findIndex(v => v.trim?.() === "");

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
    Logger.log(`🧹 ${aSupprimer} lignes vides supprimées à partir de la ligne ${rowASupprimer}`);
  }

  // 12️⃣ Banding
  const nbColonnes = entetesFinales.length;
  const nbLignes = feuille.getLastRow();
  feuille.getBandings().forEach(b => b.remove());
  feuille.getRange(1, 1, nbLignes, nbColonnes)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor("#073763")
    .setFirstRowColor("#FFFFFF")
    .setSecondRowColor("#F3F3F3");
  Logger.log("🎨 Banding appliqué (gris clair) sur toute la feuille");

  // 13️⃣ Figer la première ligne
  feuille.setFrozenRows(1);

  // Création du filtre sur toute la plage (en-tête + données)
  lastRow = feuille.getLastRow();
  feuille.getRange(1, 1, lastRow, entetesFinales.length).createFilter();
  Logger.log("📌 Filtre créé sur toute la plage de données");
  if (lastRow > 1) {
    const filter = feuille.getFilter();
    if (filter) {
      filter.sort(2, true); // colonne 2 = B
      Logger.log("🔤 Tri A à Z appliqué sur la colonne URL (colonne B/2) via le filtre");
    }
  }

  // 14️⃣ Redimensionnement personnalisé des colonnes
  feuille.setColumnWidth(1, 125);   // Colonne A
  feuille.setColumnWidth(2, 750);   // Colonne B
  feuille.setColumnWidth(3, 125);   // Colonne C
  feuille.setColumnWidth(4, 750);   // Colonne D
  feuille.setColumnWidth(5, 250);   // Colonne E
  feuille.setColumnWidth(6, 750);   // Colonne F
  feuille.setColumnWidth(7, 750);   // Colonne G
  feuille.setColumnWidth(8, 750);   // Colonne H (🆕 Meta desc)
  Logger.log("📏 Largeurs de colonnes personnalisées appliquées");

  // 15️⃣ Ajout d’un lien vers "Inventaire" dans Suivi!F si B == "Arborescence"
  try {
    const feuilleSuivi = ss.getSheetByName("Suivi");
    if (feuilleSuivi) {
      const lastRowSuivi = feuilleSuivi.getLastRow();
      const valeursColB = feuilleSuivi.getRange(1, 2, lastRowSuivi).getValues().flat();
      const gidInventaire = feuille.getSheetId();
      let liensAjoutes = 0;

      for (let i = 0; i < valeursColB.length; i++) {
        if (typeof valeursColB[i] === "string" && valeursColB[i].trim().toLowerCase() === "inventaire") {
          const formuleLien = `=HYPERLINK("#gid=${gidInventaire}";"Inventaire")`;
          feuilleSuivi.getRange(i + 1, 6).setFormula(formuleLien);
          liensAjoutes++;
          Logger.log(`[importerDonneesCrawlProd] Lien ajouté dans Suivi!F${i + 1}`);
        }
      }
      Logger.log(`[importerDonneesCrawlProd] ${liensAjoutes} lien(s) vers "Inventaire" ajoutés dans Suivi!F`);
    } else {
      Logger.log("[importerDonneesCrawlProd] Feuille Suivi non trouvée.");
    }
  } catch (e) {
    Logger.log("[importerDonneesCrawlProd] Erreur lors de l’ajout du lien Inventaire dans Suivi : " + e.message);
  }
  completerInventaireStratEtPositionnement();
  reordonnerFeuillesVisibles();

  return "success";
}

function completerInventaireStratEtPositionnement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleInventaire = ss.getSheetByName("Inventaire");
  const feuilleStrat = ss.getSheetByName("Stratégie de positionnement");
  const feuillePositionnement = ss.getSheetByName("Positionnement");

  if (!feuilleInventaire || !feuilleStrat || !feuillePositionnement) {
    Logger.log("❌ Feuille manquante : Inventaire, Stratégie de positionnement, ou Positionnement");
    return "❌ Feuille manquante";
  }

  // 1️⃣ Lecture des données Inventaire
  const lastRowInv = feuilleInventaire.getLastRow();
  const donneesInventaire = feuilleInventaire.getRange(2, 1, lastRowInv - 1, 8).getValues(); // A2:H
  Logger.log("🔎 Données Inventaire récupérées (" + donneesInventaire.length + " lignes)");

  // 2️⃣ Lecture Stratégie de positionnement (on récupère tout pour faire des recherches rapides)
  const lastRowStrat = feuilleStrat.getLastRow();
  const donneesStrat = feuilleStrat.getRange(2, 1, lastRowStrat - 1, 8).getValues(); // A2:H
  Logger.log("🔎 Données Stratégie de positionnement récupérées (" + donneesStrat.length + " lignes)");

  // Construction d'un index rapide : URL (col D) => { ligne, valeurs }
  const mapStrat = {};
  donneesStrat.forEach((row, idx) => {
    const url = (row[3] || "").trim(); // Colonne D = index 3
    if (url) {
      mapStrat[url] = { idx, row };
    }
  });

  // 3️⃣ Lecture de la feuille Positionnement
  const lastRowPos = feuillePositionnement.getLastRow();
  const donneesPos = feuillePositionnement.getRange(2, 1, lastRowPos - 1, 5).getValues(); // A2:E
  Logger.log("🔎 Données Positionnement récupérées (" + donneesPos.length + " lignes)");

  // Construction d'un index rapide : URL (col D) => { ligne, valeurs }
  const mapPos = {};
  donneesPos.forEach((row, idx) => {
    const url = (row[3] || "").trim(); // Colonne D = index 3
    const position = parseInt(row[2], 10); // Colonne C = index 2
    if (url && !isNaN(position) && position <= 20) { // Positionnement <=20
      mapPos[url] = { idx, row, position };
    }
  });

  // Pour marquer les lignes d'inventaire déjà traitées par la strat
  const lignesTraitees = new Set();

  // 4️⃣ Boucle principale sur Inventaire
  donneesInventaire.forEach((ligne, i) => {
    const urlInv = (ligne[1] || "").trim(); // Col B = index 1

    // --- Étape 1 : Recherche dans Stratégie de positionnement ---
    if (urlInv && mapStrat[urlInv]) {
      const dataStrat = mapStrat[urlInv].row;

      // Case à cocher colonne C (index 2)
      feuilleInventaire.getRange(i + 2, 3).setValue(true);

      // Nouvelle URL (D, index 3) : remplir seulement si vide
      if (!ligne[3] || ligne[3].toString().trim() === "") {
        feuilleInventaire.getRange(i + 2, 4).setValue(dataStrat[4] || ""); // Col E de strat (index 4)
      }
      // Mot-clé visé (E, index 4) : remplir seulement si vide
      if (!ligne[4] || ligne[4].toString().trim() === "") {
        feuilleInventaire.getRange(i + 2, 5).setValue(dataStrat[1] || ""); // Col B de strat (index 1)
      }

      lignesTraitees.add(urlInv);
      Logger.log(`[Inventaire][${i + 2}] URL trouvée dans Stratégie : cochée, complétée si besoin`);
      return; // Ne passe pas à l’étape 2 si déjà traité
    }

    // --- Étape 2 : Recherche dans Positionnement ---
    if (urlInv && mapPos[urlInv]) {
      // Position (col C) doit être ≤ 20 (déjà filtré dans le mapPos)
      // Case à cocher colonne C (index 2)
      feuilleInventaire.getRange(i + 2, 3).setValue(true);

      // Mot-clé visé (E, index 4) : remplir seulement si vide
      if (!ligne[4] || ligne[4].toString().trim() === "") {
        feuilleInventaire.getRange(i + 2, 5).setValue(mapPos[urlInv].row[0] || ""); // Col A de positionnement (index 0)
      }
      Logger.log(`[Inventaire][${i + 2}] URL trouvée dans Positionnement (pos≤20) : cochée, mot-clé visé complété si besoin`);
    }
  });

  Logger.log("✅ Complétion de l’inventaire terminée.");
  return "success";
}

function importerStrategiePositionnement(url) {
  try {
    // 1️⃣ Extraction de l'ID du Google Sheet source depuis l'URL
    const regex = /\/d\/([a-zA-Z0-9-_]+)/;
    const match = url.match(regex);
    if (!match || !match[1]) {
      Logger.log("❌ Impossible d'extraire l'ID du fichier depuis l'URL.");
      throw new Error("❌ URL Google Sheet invalide. Vérifie l'URL collée.");
    }
    const fileId = match[1];
    Logger.log("🔑 ID du fichier source extrait : " + fileId);

    // 2️⃣ Ouverture du classeur source
    let classeurSource;
    try {
      classeurSource = SpreadsheetApp.openById(fileId);
    } catch (e) {
      Logger.log("❌ Impossible d'ouvrir le classeur source : " + e.message);
      throw new Error("❌ Impossible d'ouvrir le classeur source. Vérifie l'accès au fichier.");
    }

    // 3️⃣ Récupération de l'onglet "Choix mots-clés"
    let feuilleSource = classeurSource.getSheetByName("Choix mots-clés");
    if (!feuilleSource) {
      Logger.log("❌ Onglet 'Choix mots-clés' introuvable dans le fichier source.");
      throw new Error("❌ L'onglet 'Choix mots-clés' est introuvable dans le fichier source.");
    }

    // 4️⃣ Lecture de la plage B4:O (on prend toutes les lignes non vides)
    const lastRow = feuilleSource.getLastRow();
    if (lastRow < 4) {
      Logger.log("❌ Pas de données à partir de la ligne 4.");
      throw new Error("❌ La feuille source ne contient pas de données à partir de la ligne 4.");
    }
    const data = feuilleSource.getRange(4, 2, lastRow - 3, 14).getValues(); // B4:O
    Logger.log(`📥 ${data.length} lignes récupérées depuis Choix mots-clés!B4:O${lastRow}`);

    // 5️⃣ Filtrage des lignes où "Mot-clé visé" (colonne E, index 3) est vide
    const filteredData = data.filter(row => row[3] && row[3].toString().trim() !== "");
    Logger.log(`🔎 ${filteredData.length} lignes après filtrage des "Mot-clé visé" vides`);

    if (filteredData.length === 0) {
      Logger.log("❌ Aucune ligne avec 'Mot-clé visé' renseigné.");
      throw new Error("❌ Aucune ligne avec 'Mot-clé visé' renseigné dans la feuille source.");
    }

    // 6️⃣ Mapping des colonnes pour la feuille cible
    // [Template (B), Mot-clé visé (E), Volume (F), URL actuelle (G), URL nouvelle (H), Title (J), H1 (K), Meta desc (L)]
    const lignes = filteredData.map(row => [
      row[0] || "",   // Colonne A : Template (col B, index 0)
      row[3] || "",   // Colonne B : Mot-clé visé (col E, index 3)
      row[4] || "",   // Colonne C : Volume (col F, index 4)
      row[5] || "",   // Colonne D : URL actuelle (col G, index 5)
      row[6] || "",   // Colonne E : URL nouvelle (col H, index 6)
      row[8] || "",   // Colonne F : Title (col J, index 8)
      row[9] || "",   // Colonne G : <h1> (col K, index 9)
      row[10] || ""   // Colonne H : Meta description (col L, index 10)
    ]);
    Logger.log("🗃️ Mapping effectué. Ex : " + JSON.stringify(lignes[0]));

    // 7️⃣ Création ou remplacement de la feuille "Stratégie de positionnement"
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let feuille = ss.getSheetByName("Stratégie de positionnement");
    if (feuille) {
      ss.deleteSheet(feuille);
      Logger.log("📄 Feuille 'Stratégie de positionnement' supprimée.");
    }
    feuille = ss.insertSheet("Stratégie de positionnement");
    feuille.setTabColor("#434343");
    Logger.log("📄 Feuille 'Stratégie de positionnement' (re)créée avec couleur #434343");

    // 8️⃣ Insertion des en-têtes
    const entetesFinales = [
      "Template", "Mot-clé visé", "Volume", "URL actuelle", "URL nouvelle", "Title", "<h1>", "Meta description"
    ];
    feuille.getRange(1, 1, 1, entetesFinales.length).setValues([entetesFinales]);

    // 9️⃣ Insertion des lignes de données
    feuille.getRange(2, 1, lignes.length, entetesFinales.length).setValues(lignes);
    Logger.log("📝 Données insérées avec en-têtes en ligne 1.");

    // 🔟 Formatage des en-têtes
    feuille.getRange(1, 1, 1, entetesFinales.length)
      .setFontWeight("bold")
      .setFontColor("white")
      .setFontFamily("Arial")
      .setFontSize(10)
      .setVerticalAlignment("middle");

    // 11️⃣ Formatage du contenu
    feuille.getRange(2, 1, lignes.length, entetesFinales.length)
      .setFontColor("black")
      .setFontSize(10)
      .setFontFamily("Arial")
      .setVerticalAlignment("middle");

    // 12️⃣ Format de nombre pour Volume (colonne C = 3)
    feuille.getRange(2, 3, lignes.length, 1)
      .setNumberFormat("#,##0")
      .setHorizontalAlignment("left");

    // 13️⃣ Masquage du quadrillage
    feuille.setHiddenGridlines(true);

    // 14️⃣ Suppression des colonnes I à Z
    const lastCol = feuille.getMaxColumns();
    if (lastCol > 8) {
      feuille.deleteColumns(9, lastCol - 8);
      Logger.log("🗑️ Colonnes I à Z supprimées.");
    }

    // 15️⃣ Nettoyage des lignes vides (colonne A)
    const lastRowFeuille = feuille.getLastRow();
    const valeursColA = feuille.getRange(2, 1, lastRowFeuille - 1).getValues().flat();
    const firstEmpty = valeursColA.findIndex(v => (typeof v === "string" ? v.trim() : v) === "");
    let rowASupprimer;
    if (firstEmpty !== -1) {
      rowASupprimer = firstEmpty + 2;
    } else {
      rowASupprimer = lastRowFeuille + 1;
    }
    const totalRows = feuille.getMaxRows();
    const aSupprimer = totalRows - rowASupprimer + 1;
    if (aSupprimer > 0) {
      feuille.deleteRows(rowASupprimer, aSupprimer);
      Logger.log(`🧹 ${aSupprimer} lignes vides supprimées à partir de la ligne ${rowASupprimer}`);
    }

    // 16️⃣ Banding (alternance de couleur)
    const nbColonnes = entetesFinales.length;
    const nbLignes = feuille.getLastRow();
    feuille.getBandings().forEach(b => b.remove());
    feuille.getRange(1, 1, nbLignes, nbColonnes)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
      .setHeaderRowColor("#073763")
      .setFirstRowColor("#FFFFFF")
      .setSecondRowColor("#F3F3F3");
    Logger.log("🎨 Banding appliqué sur toute la feuille.");

    // 17️⃣ Figer la première ligne + filtre
    feuille.setFrozenRows(1);
    feuille.getRange(1, 1, 1, nbColonnes).createFilter();
    Logger.log("📌 Première ligne figée + filtre activé.");

    // 18️⃣ Redimensionnement des colonnes
    feuille.setColumnWidth(1, 200);   // Template
    feuille.setColumnWidth(2, 250);   // Mot-clé visé
    feuille.setColumnWidth(3, 100);   // Volume
    feuille.setColumnWidth(4, 500);   // URL actuelle
    feuille.setColumnWidth(5, 500);   // URL nouvelle
    feuille.setColumnWidth(6, 350);   // Title
    feuille.setColumnWidth(7, 250);   // <h1>
    feuille.setColumnWidth(8, 350);   // Meta description
    Logger.log("📏 Largeurs de colonnes personnalisées appliquées.");

    reordonnerFeuillesVisibles();

    return "success";

  } catch (error) {
    Logger.log("❌ Erreur dans importerStrategiePositionnement : " + error.message);
    throw error;
  }

}

function importerDonneesSemrush(donnees) {
  // Étape 1 : Initialisation et récupération du Spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Étape 2 : Suppression de la feuille "Positionnement" si elle existe
  let feuille = ss.getSheetByName("Positionnement");
  if (feuille) {
    ss.deleteSheet(feuille);
    Logger.log("📄 Feuille 'Positionnement' supprimée.");
  }

  // Étape 3 : Création de la nouvelle feuille "Positionnement"
  feuille = ss.insertSheet("Positionnement");
  feuille.setTabColor("#434343");
  Logger.log("📄 Feuille 'Positionnement' (re)créée avec couleur #5b0f00");

  // Étape 4 : Extraction des indices des colonnes nécessaires dans l'en-tête
  const enTetes = donnees[0];
  const idxKeyword = enTetes.indexOf("Keyword");
  const idxVolume = enTetes.indexOf("Search Volume");
  const idxPosition = enTetes.indexOf("Position");
  const idxURL = enTetes.indexOf("URL");
  const idxTraffic = enTetes.indexOf("Traffic");

  if ([idxKeyword, idxVolume, idxPosition, idxURL, idxTraffic].includes(-1)) {
    Logger.log("❌ Colonnes essentielles manquantes dans le CSV SEMrush.");
    throw new Error("❌ Colonnes essentielles manquantes dans le CSV SEMrush.");
  }

  // Étape 5 : Construction des lignes à insérer (A=Mot-clé, B=Volume, C=Position, D=URL, E=Trafic)
  const lignes = donnees.slice(1).map(ligne => [
    ligne[idxKeyword] || "",      // Colonne A : Mot-clé
    ligne[idxVolume] || "",       // Colonne B : Volume
    ligne[idxPosition] || "",     // Colonne C : Position
    ligne[idxURL] || "",          // Colonne D : URL
    ligne[idxTraffic] || ""       // Colonne E : Trafic
  ]);
  Logger.log(`📥 ${lignes.length} lignes importées depuis le CSV.`);

  // Étape 6 : Insertion des en-têtes et du contenu
  const entetesFinales = ["Mot-clé", "Volume", "Position", "URL", "Trafic"];
  feuille.getRange(1, 1, 1, entetesFinales.length).setValues([entetesFinales]);
  feuille.getRange(2, 1, lignes.length, entetesFinales.length).setValues(lignes);
  Logger.log("📝 Données insérées avec en-têtes en ligne 1.");

  // Étape 7 : Formatage des en-têtes
  const rangeEntetes = feuille.getRange(1, 1, 1, entetesFinales.length);
  rangeEntetes.setFontWeight("bold")
    .setFontColor("white")
    .setFontFamily("Arial")
    .setFontSize(10)
    .setVerticalAlignment("middle");
  // Alignement horizontal milieu pour Volume (B) et Position (C)
  feuille.getRange(1, 2).setHorizontalAlignment("center");
  feuille.getRange(1, 3).setHorizontalAlignment("center");

  // Étape 8 : Formatage du contenu
  const nbLignes = lignes.length;
  const rangeContenu = feuille.getRange(2, 1, nbLignes, entetesFinales.length);
  rangeContenu.setFontColor("black")
    .setFontSize(10)
    .setFontFamily("Arial")
    .setVerticalAlignment("middle");

  // Étape 9 : Alignements personnalisés pour les données
  feuille.getRange(2, 1, nbLignes, 1).setHorizontalAlignment("left");   // Mot-clé (A)
  feuille.getRange(2, 2, nbLignes, 1).setHorizontalAlignment("center"); // Volume (B)
  feuille.getRange(2, 3, nbLignes, 1).setHorizontalAlignment("center"); // Position (C)
  feuille.getRange(2, 4, nbLignes, 1).setHorizontalAlignment("left");   // URL (D)
  feuille.getRange(2, 5, nbLignes, 1).setHorizontalAlignment("left");   // Trafic (E)

  // Étape 10 : Format de nombre sur Volume (B) et Trafic (E) : séparateur de milliers, 0 décimale
  feuille.getRange(2, 2, nbLignes, 1).setNumberFormat("#,##0");
  feuille.getRange(2, 5, nbLignes, 1).setNumberFormat("#,##0");

  // Étape 11 : Masquage du quadrillage
  feuille.setHiddenGridlines(true);

  // Étape 12 : Suppression des colonnes F à Z
  const lastCol = feuille.getMaxColumns();
  if (lastCol > 5) {
    feuille.deleteColumns(6, lastCol - 5);
    Logger.log("🗑️ Colonnes F à Z supprimées.");
  }

  // Étape 13 : Nettoyage des lignes vides
  const lastRow = feuille.getLastRow();
  const valeursColA = feuille.getRange(2, 1, lastRow - 1).getValues().flat();
  const firstEmpty = valeursColA.findIndex(v => (typeof v === "string" ? v.trim() : v) === "");
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
    Logger.log(`🧹 ${aSupprimer} lignes vides supprimées à partir de la ligne ${rowASupprimer}`);
  }

  // Étape 14 : Banding (alternance de couleur)
  const nbColonnes = entetesFinales.length;
  const nbLignesTotales = feuille.getLastRow();
  feuille.getBandings().forEach(b => b.remove());
  const plageBanding = feuille.getRange(1, 1, nbLignesTotales, nbColonnes);
  plageBanding
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor("#073763")
    .setFirstRowColor("#FFFFFF")
    .setSecondRowColor("#F3F3F3");
  Logger.log("🎨 Banding appliqué sur toute la feuille.");

  // Étape 15 : Figer la première ligne + filtre
  feuille.setFrozenRows(1);
  feuille.getRange(1, 1, 1, nbColonnes).createFilter();
  Logger.log("📌 Première ligne figée + filtre activé.");

  // Étape 16 : Redimensionnement des colonnes
  feuille.setColumnWidth(1, 250);   // Colonne A : Mot-clé
  feuille.setColumnWidth(2, 150);   // Colonne B : Volume
  feuille.setColumnWidth(3, 150);   // Colonne C : Position
  feuille.setColumnWidth(4, 800);   // Colonne D : URL
  feuille.setColumnWidth(5, 150);   // Colonne E : Trafic
  Logger.log("📏 Largeurs de colonnes personnalisées appliquées.");

  // Étape 17 : Mise en forme conditionnelle (dégradé bleu sur B2:B)
  const nbLignesData = feuille.getLastRow() - 1; // On retire l'en-tête
  if (nbLignesData > 0) {
    const volumeRange = feuille.getRange(2, 2, nbLignesData, 1);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpoint("#cfe2f3")
      .setGradientMaxpoint("#0b5394")
      .setRanges([volumeRange])
      .build();
    const rules = feuille.getConditionalFormatRules();
    rules.push(rule);
    feuille.setConditionalFormatRules(rules);
    Logger.log("🌈 Dégradé bleu appliqué sur la colonne Volume (B2:B)");
  }

  reordonnerFeuillesVisibles();

  return "success";
}

function insertAuditRowsForSegments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  const auditSheet = ss.getSheetByName("Audit");
  if (!configSheet || !auditSheet) return;

  const startRow = 15;
  const lastRow = configSheet.getLastRow();
  const segmentUrlData = configSheet.getRange(`B${startRow}:C${lastRow}`).getValues();
  const validEntries = segmentUrlData.filter(([segment, url]) => segment && url);

  if (validEntries.length === 0) return;

  const auditValues = auditSheet.getRange(`B1:B${auditSheet.getLastRow()}`).getValues();

  const findRowIndex = label =>
    auditValues.findIndex(row => (row[0] || '').toString().trim() === label);

  const rowBalise = findRowIndex("Balise hn");
  const rowDonnees = findRowIndex("Données structurées");
  const rowRendu = findRowIndex("Rendu");
  const rowMobile = findRowIndex("Compatibilité mobile");

  if ([rowBalise, rowDonnees, rowRendu, rowMobile].some(index => index === -1)) {
    throw new Error('❌ Une ou plusieurs ancres manquent dans Audit!B:B (Rendu, Balise hn, Données structurées, Compatibilité mobile)');
  }

  const tasks = [
    {
      label: "Balisage hn",
      description: "Inspection du balisage hn",
      colA: "Positionnement",
      baseRow: rowBalise
    },
    {
      label: "Données structurées",
      description: "Inspection des données structurées",
      colA: "Positionnement",
      baseRow: rowDonnees
    },
    {
      label: "Rendu",
      description: "Inspection du rendu",
      colA: "Structure",
      baseRow: rowRendu
    },
    {
      label: "Compatibilité mobile",
      description: "Inspection compatibilité mobile",
      colA: "Web performance",
      baseRow: rowMobile
    }
  ];

  // 🧠 Insérer du bas vers le haut pour éviter de fausser les indices
  tasks.sort((a, b) => b.baseRow - a.baseRow);

  tasks.forEach(({ label, description, colA, baseRow }) => {
    const insertAt = baseRow + 2;
    auditSheet.insertRows(insertAt, validEntries.length);

    const rows = validEntries.map(([segment, url]) => [
      colA,
      label,
      `${description} - ${segment}`,
      "Recette à faire",
      url
    ]);

    auditSheet.getRange(insertAt, 1, rows.length, 5).setValues(rows);
  });
}

function getDocumentProperties() {
  // Étape 1 - Accès au service de propriétés
  const props = PropertiesService.getDocumentProperties();
  const allProps = props.getProperties();

  // Étape 2 - Log d'information pour vérification
  Logger.log('[INFO] Propriétés du document récupérées : %s', JSON.stringify(allProps));

  // Vérification des clés attendues (log d'avertissement si absentes)
  if (!allProps.urlActuel) Logger.log('[WARN] urlActuel est vide ou non défini');
  if (!allProps.urlPreprod) Logger.log('[WARN] urlPreprod est vide ou non défini');
  if (!allProps.templates) Logger.log('[WARN] templates est vide ou non défini');

  // Étape 3 - Retour des propriétés au client
  return allProps;
}

function creerArborescence() {
  Logger.log("=== [1] DÉBUT creerArborescence ===");

  // [1] Ouverture du classeur et récupération des propriétés document
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const urlActuel = props.getProperty('urlActuel') || '';
  const urlPreprod = props.getProperty('urlPreprod') || '';
  Logger.log(`[1.1] Propriétés récupérées : urlActuel=${urlActuel}, urlPreprod=${urlPreprod}`);

  // [2] Lecture des données sources
  const feuilleInventaire = ss.getSheetByName('Inventaire');
  const feuilleStrat = ss.getSheetByName('Stratégie de positionnement');
  if (!feuilleInventaire || !feuilleStrat) {
    Logger.log("[1.2] ❌ Feuille Inventaire ou Stratégie de positionnement manquante");
    throw new Error("Feuille Inventaire ou Stratégie de positionnement manquante.");
  }
  const dataInv = feuilleInventaire.getRange(2, 1, feuilleInventaire.getLastRow()-1, 8).getValues(); // A2:H
  const dataStrat = feuilleStrat.getRange(2, 1, feuilleStrat.getLastRow()-1, 8).getValues();         // A2:H
  Logger.log(`[1.3] ${dataInv.length} lignes Inventaire, ${dataStrat.length} lignes Strat récupérées`);

  // [3] Extraction et normalisation des URLs (fusion et suppression des doublons)
  const mapUrls = construireMapUrls(dataInv, dataStrat, urlActuel, urlPreprod);
  Logger.log(`[2] ${Object.keys(mapUrls).length} URLs uniques normalisées extraites`);

  // [4] Extraction des templates uniques pour la liste déroulante
  const templatesUniques = [...new Set(dataInv.map(row => (row[0] || '').toString()).filter(t => t))];
  Logger.log(`[3] ${templatesUniques.length} templates uniques extraits pour la liste déroulante`);

  // [5] Génération des URLs pages système Mentions légales, Politique de confidentialité, Cookies
  // ---------------------------------------------------------------------------------------------
  // 1️⃣ On récupère les propriétés nécessaires
  const finUrlNouveau = props.getProperty('finurlnouveau') || '';
  Logger.log("[creerArborescence][pages systeme][1] finurlnouveau = " + finUrlNouveau);

  // 2️⃣ On construit les URLs (attention à ne pas doubler le slash)
  function joinWithFin(urlBase, fin) {
    // Si la fin = "/" ou "", on ne double pas le slash
    if (fin === "/" || fin === "") return urlBase.replace(/\/+$/, "") + "/";
    return urlBase.replace(/\/+$/, "") + fin;
  }

  // URLs
  const urlMentions = urlPreprod.replace(/\/+$/, "") + "/mentions-legales/";
  const urlPolitique = joinWithFin(urlPreprod + "/mentions-legales/politique-de-confidentialite", finUrlNouveau);
  const urlCookies = joinWithFin(urlPreprod + "/mentions-legales/cookies", finUrlNouveau);

  Logger.log("[creerArborescence][pages systeme][2] URLs générées :");
  Logger.log("  - Mentions légales : " + urlMentions);
  Logger.log("  - Politique de confidentialité : " + urlPolitique);
  Logger.log("  - Cookies : " + urlCookies);

  // 3️⃣ Génération des infos de chaque page via analyserUrlArbo
  const pagesSysteme = [
    {
      url: urlMentions
    },
    {
      url: urlPolitique
    },
    {
      url: urlCookies
    }
  ].map(page => {
    const analyse = analyserUrlArbo(page.url, urlPreprod);
    Logger.log("[creerArborescence][pages systeme][3] Analyse pour URL : " + page.url + " => " + JSON.stringify(analyse));
    // [Template] => "" (menu déroulant appliqué après), [Title/H1/Meta desc] => vide
    return [
      '',                        // Template (vide → menu déroulant)
      analyse.nom,               // Nom
      page.url,                  // URL
      false,                     // Intégré ?
      analyse.niveau,            // Niveau
      analyse.niveauParent,      // Niveau parent
      analyse.filAriane,         // Fil d’Ariane
      '',                        // Title
      '',                        // <h1>
      '',                        // Meta description
      ''                         // Recette
    ];
  });

  Logger.log("[creerArborescence][pages systeme][4] Lignes pages système générées : " + JSON.stringify(pagesSysteme));

  // 4️⃣ On génère les lignes habituelles pour le reste des URLs (comme avant)
  const lignesArboBase = Object.keys(mapUrls).map(url => {
    const meta = mapUrls[url];
    const analyse = analyserUrlArbo(url, urlPreprod);
    return [
      meta.template || '',                // Template (menu déroulant après si vide)
      analyse.nom,                        // Nom
      url,                                // URL
      false,                              // Intégré ?
      analyse.niveau,                     // Niveau
      analyse.niveauParent,               // Niveau parent
      analyse.filAriane,                  // Fil d'Ariane
      meta.title || '',                   // Title
      meta.h1 || '',                      // <h1>
      meta.metaDesc || '',                // Meta description
      ''                                  // Recette
    ];
  });

  // 5️⃣ On concatène les pages système et le reste (ordre peu importe, ce sera trié après)
  const lignesArbo = [...lignesArboBase, ...pagesSysteme];
  Logger.log("[creerArborescence][pages systeme][5] Lignes finales pour Arborescence : " + lignesArbo.length);


  // [6] Création/remplacement de la feuille "Arborescence"
  let feuilleArbo = ss.getSheetByName('Arborescence');
  if (feuilleArbo) {
    ss.deleteSheet(feuilleArbo);
    Logger.log("[5.1] Ancienne feuille Arborescence supprimée");
  }
  feuilleArbo = ss.insertSheet('Arborescence');
  feuilleArbo.setTabColor("#f39c12");
  Logger.log("[5.2] Nouvelle feuille Arborescence créée");

  // [7] Insertion des en-têtes
  const entetes = ["Template", "Nom", "URL", "Intégré ?", "Niveau", "Niveau parent", "Fil d'Ariane", "Title", "<h1>", "Meta description", "Recette"];
  feuilleArbo.getRange(1, 1, 1, entetes.length).setValues([entetes]);
  Logger.log("[6] En-têtes insérés");

  // [8] Insertion des données de l'arborescence
  if (lignesArbo.length > 0) {
    feuilleArbo.getRange(2, 1, lignesArbo.length, entetes.length).setValues(lignesArbo);
    Logger.log("[7] Données arborescence insérées");
  }

  // [9] Formatage premium (banding, largeur, etc.)
  appliquerFormatageArborescence(feuilleArbo, lignesArbo.length, entetes.length, templatesUniques);
  Logger.log("[8] Formatage premium appliqué");

  // [9.1] Coloration spécifique APRÈS banding ! 
  const colTitle = 8, colH1 = 9, colMeta = 10;
  const urls = feuilleArbo.getRange(2, 3, lignesArbo.length, 1).getValues().flat();
  for (let i = 0; i < lignesArbo.length; i++) {
    const url = urls[i];
    if (mapUrls[url]) {
      if (mapUrls[url].provenance && mapUrls[url].provenance.title === "inventaire")
        feuilleArbo.getRange(i + 2, colTitle).setBackground("#fce5cd");
      if (mapUrls[url].provenance && mapUrls[url].provenance.h1 === "inventaire")
        feuilleArbo.getRange(i + 2, colH1).setBackground("#fce5cd");
      if (mapUrls[url].provenance && mapUrls[url].provenance.metaDesc === "inventaire")
        feuilleArbo.getRange(i + 2, colMeta).setBackground("#fce5cd");
    }
  }
  Logger.log("[Color] Coloration #fce5cd sur Title, <h1> et Meta description issus d'Inventaire.");

  // [10] Ajout du lien dans Suivi!F pour chaque ligne où B == "Arborescence"
  ajouterLienSuiviVersArborescence(ss, feuilleArbo);
  Logger.log("[9] Lien ajouté dans Suivi!F");

  Logger.log("=== [10] FIN creerArborescence ===");

  reordonnerFeuillesVisibles();
}

function construireMapUrls(dataInv, dataStrat, urlActuel, urlPreprod) {
  const mapInv = {};
  dataInv.forEach(row => {
    const urlRaw = (row[3] || '').toString().trim(); // D
    if (!urlRaw) return;
    const urlNorm = remplacerUrl(urlRaw, urlActuel, urlPreprod);
    mapInv[urlNorm] = {
      source: 'inventaire',
      template: (row[0] || '').toString(),
      title: (row[5] || '').toString(),
      h1: (row[6] || '').toString(),
      metaDesc: (row[7] || '').toString(),
      // On marque explicitement la provenance
      provenance: {
        title: 'inventaire',
        h1: 'inventaire',
        metaDesc: 'inventaire'
      }
    };
  });

  const mapUrls = { ...mapInv };
  dataStrat.forEach(row => {
    const urlRaw = (row[4] || '').toString().trim(); // E
    if (!urlRaw) return;
    const urlNorm = remplacerUrl(urlRaw, urlActuel, urlPreprod);
    if (!mapUrls[urlNorm]) {
      // URL présente uniquement dans Strat, pas dans Inventaire
      mapUrls[urlNorm] = {
        source: 'strat',
        template: '', // Sera rempli par une liste déroulante plus tard
        title: (row[5] || '').toString(),
        h1: (row[6] || '').toString(),
        metaDesc: (row[7] || '').toString(),
        provenance: {
          title: 'strat',
          h1: 'strat',
          metaDesc: 'strat'
        }
      };
    } else {
      // Si déjà présent, priorité à la strat pour les champs title/h1/metaDesc si dispo
      // On garde trace de la provenance réelle de chaque champ
      if (row[5]) {
        mapUrls[urlNorm].title = (row[5] || '').toString();
        mapUrls[urlNorm].provenance.title = 'strat';
      }
      if (row[6]) {
        mapUrls[urlNorm].h1 = (row[6] || '').toString();
        mapUrls[urlNorm].provenance.h1 = 'strat';
      }
      if (row[7]) {
        mapUrls[urlNorm].metaDesc = (row[7] || '').toString();
        mapUrls[urlNorm].provenance.metaDesc = 'strat';
      }
    }
  });

  return mapUrls;
}

function remplacerUrl(url, urlActuel, urlPreprod) {
  if (!urlActuel || !urlPreprod) return url;
  return url.split(urlActuel).join(urlPreprod);
}

function appliquerFormatageArborescence(feuille, nbLignes, nbColonnes, templatesUniques) {
  Logger.log("[F1] Début formatage premium Arborescence");

  // 1️⃣ En-têtes en gras, fond bleu, police Arial, taille 10
  feuille.getRange(1, 1, 1, nbColonnes)
    .setFontWeight("bold")
    .setFontColor("white")
    .setBackground("#073763")
    .setFontFamily("Arial")
    .setFontSize(10)
    .setVerticalAlignment("middle");

  // 2️⃣ Contenu police Arial, taille 10, couleur noire
  if (nbLignes > 0) {
    feuille.getRange(2, 1, nbLignes, nbColonnes)
      .setFontColor("black")
      .setFontSize(10)
      .setFontFamily("Arial")
      .setVerticalAlignment("middle");
  }

  // 3️⃣ Cases à cocher sur "Intégré ?" (colonne 4)
  if (nbLignes > 0) {
    feuille.getRange(2, 4, nbLignes).insertCheckboxes();
    Logger.log("[F2] Cases à cocher insérées sur colonne 4");
  }

  // 4️⃣ Liste déroulante dans la colonne Template (colonne 1) là où vide
  if (nbLignes > 0 && templatesUniques.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(templatesUniques)
      .setAllowInvalid(true)
      .build();

    const valeursTemplate = feuille.getRange(2, 1, nbLignes, 1).getValues().flat();
    valeursTemplate.forEach((v, i) => {
      if (!v || v === '') {
        feuille.getRange(i + 2, 1).setDataValidation(rule);
      }
    });
    Logger.log("[F3] Liste déroulante appliquée sur colonnes Template vides");
  }

  // 5️⃣ Masquage du quadrillage
  feuille.setHiddenGridlines(true);
  Logger.log("[F4] Quadrillage désactivé");

  // 6️⃣ Banding (alternance couleurs)
  feuille.getBandings().forEach(b => b.remove());
  const plageBanding = feuille.getRange(1, 1, nbLignes + 1, nbColonnes);
  plageBanding
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor("#073763")
    .setFirstRowColor("#FFFFFF")
    .setSecondRowColor("#F3F3F3");
  Logger.log("[F5] Banding appliqué");

  // 7️⃣ Figer la première ligne
  feuille.setFrozenRows(1);

  // 8️⃣ Largeur colonnes (alignée avec Inventaire)
  const largeurs = [125, 250, 750, 100, 75, 120, 300, 600, 350, 500, 200];
  for (let i = 0; i < nbColonnes; i++) {
    feuille.setColumnWidth(i + 1, largeurs[i] || 120);
  }
  Logger.log("[F7] Largeurs de colonnes appliquées");

  // 9️⃣ Suppression des colonnes inutiles (J+)
  const maxCols = feuille.getMaxColumns();
  if (maxCols > nbColonnes) {
    feuille.deleteColumns(nbColonnes + 1, maxCols - nbColonnes);
    Logger.log("[F8] Colonnes supplémentaires supprimées");
  }

  // 🔟 Suppression des lignes vides après la dernière ligne de donnée
  const maxRows = feuille.getMaxRows();
  if (maxRows > nbLignes + 1) {
    feuille.deleteRows(nbLignes + 2, maxRows - nbLignes - 1);
    Logger.log("[F9] Lignes supplémentaires supprimées");
  }

  // (Nouveau calcul pour le filtre/tri)
  const lastRow = feuille.getLastRow();

  // 1️⃣1️⃣ Création du filtre sur toute la plage de données (en-tête inclus)
  if (lastRow > 1) {
    feuille.getRange(1, 1, lastRow, nbColonnes).createFilter();
    Logger.log("[F10] Filtre créé sur toute la plage de Arborescence.");

    // 1️⃣2️⃣ Tri via le filtre sur la colonne 3 (URL), sans toucher à l'en-tête
    const filter = feuille.getFilter();
    if (filter) {
      filter.sort(3, true); // Colonne 3 (URL), tri croissant (A>Z)
      Logger.log("[F11] Tableau trié de A à Z sur la colonne URL (colonne 3) via le filtre.");
    }
  }

  Logger.log("[F12] Fin formatage Arborescence");
}

function ajouterLienSuiviVersArborescence(ss, feuilleArbo) {
  try {
    const feuilleSuivi = ss.getSheetByName("Suivi");
    if (!feuilleSuivi) {
      Logger.log("[ajouterLienSuiviVersArborescence] Feuille Suivi introuvable.");
      return;
    }
    const lastRow = feuilleSuivi.getLastRow();
    const valeursColB = feuilleSuivi.getRange(1, 2, lastRow).getValues().flat();
    const gidArbo = feuilleArbo.getSheetId();
    let liensAjoutes = 0;
    for (let i = 0; i < valeursColB.length; i++) {
      if (typeof valeursColB[i] === "string" && valeursColB[i].trim().toLowerCase() === "arborescence") {
        // Colonne F = colonne 6
        const formuleLien = `=HYPERLINK("#gid=${gidArbo}";"Arborescence")`;
        feuilleSuivi.getRange(i + 1, 6).setFormula(formuleLien);
        liensAjoutes++;
        Logger.log(`[ajouterLienSuiviVersArborescence] Lien ajouté en Suivi!F${i + 1}`);
      }
    }
    Logger.log(`[ajouterLienSuiviVersArborescence] ${liensAjoutes} lien(s) ajoutés`);
  } catch (e) {
    Logger.log(`[ajouterLienSuiviVersArborescence] Erreur : ${e.message}`);
  }
}

function analyserUrlArbo(url, urlPreprod) {
  // 1️⃣ On enlève le domaine (et urlPreprod si fourni)
  let chemin = url;
  // Retirer protocole + domaine
  chemin = chemin.replace(/^https?:\/\/[^\/]+/i, "");
  // Retirer urlPreprod si présent en début de l'URL (pour projets avec changement de domaine)
  if (urlPreprod) {
    const preprodClean = urlPreprod.replace(/^https?:\/\//i, "");
    chemin = chemin.replace(preprodClean, "");
  }
  if (!chemin.startsWith("/")) chemin = "/" + chemin;

  // 2️⃣ Découpe en segments (enlevant d'abord les éventuels paramètres ou fragments)
  chemin = chemin.replace(/[#?].*$/, ""); // enlève ?param, #fragment
  let cheminNet = chemin.replace(/\/{2,}/g, "/"); // plusieurs "/" → un seul
  if (cheminNet.endsWith("/")) cheminNet = cheminNet.slice(0, -1);
  if (cheminNet === "") cheminNet = "/";

  // 3️⃣ On isole les segments non vides
  const segments = cheminNet.split("/").filter(seg => seg.trim() !== "");
  const niveau = segments.length;
  // Nom = dernier segment propre
  let nom = "";
  if (segments.length > 0) {
    nom = segments[segments.length - 1]
      .replace(/\.html?$/i, "")      // enlève .html/.htm
      .replace(/[-_]+/g, " ")        // remplace tirets/bas par espace
      .replace(/\/$/, "");           // enlève slash final
    nom = nom.charAt(0).toUpperCase() + nom.slice(1);
  }

  // Niveau parent = segment précédent (ou "Aucun")
  let niveauParent = "";
  if (segments.length > 1) {
    let parent = segments[segments.length - 2]
      .replace(/\.html?$/i, "")
      .replace(/[-_]+/g, " ")
      .replace(/\/$/, "");
    niveauParent = parent.charAt(0).toUpperCase() + parent.slice(1);
  }

  // Fil d’Ariane = concat de tous les segments formatés
  const filAriane = segments.map(seg => {
    let v = seg.replace(/\.html?$/i, "")
      .replace(/[-_]+/g, " ")
      .replace(/\/$/, "");
    return v.charAt(0).toUpperCase() + v.slice(1);
  }).join(" > ");

  return {
    cheminSansDomaine: cheminNet,
    nom: nom,
    niveau: niveau,
    niveauParent: niveauParent,
    filAriane: filAriane
  };
}

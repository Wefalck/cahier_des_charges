function generateBalisageHnSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  if (!configSheet) throw new Error("Feuille 'Config' introuvable");

  const data = configSheet.getRange("B15:C" + configSheet.getLastRow()).getValues();
  const validTemplates = data.filter(([segment, url]) => segment && url);

  if (validTemplates.length === 0) {
    throw new Error("Aucun couple segment/URL trouvé dans Config!B15:C");
  }

  const sheetName = "Balisage hn";
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  const row1 = [];
  const row2 = [];
  const row3 = [];

  validTemplates.forEach(([template, url]) => {
    const name = template.toLowerCase().trim();
    row1.push(`Template ${name} - Balisage actuel`);
    row1.push(`Template ${name} - Balisage préconisé`);
    
    row2.push("Exemple pour :");
    row2.push(""); // Colonne préconisé vide

    row3.push(url);
    row3.push(""); // Colonne préconisé vide
  });

  const totalCols = row1.length;
  const totalRows = 10;

    // 🟢 Ajout des légendes explicatives en haut
  const legendValues = [
    ["Balise & contenu OK"],
    ["Contenu OK mais balise KO"],
    ["Balise & contenu KO"]
  ];

  const legendColors = ["#d9ead3", "#fce5cd", "#ffcfc9"];

  legendValues.forEach((row, i) => {
    const range = sheet.getRange(i + 1, 1);
    range.setValue(row[0]);
    range.setFontFamily("Arial");
    range.setFontSize(12);
    range.setHorizontalAlignment("center");
    range.setBackground(legendColors[i]);
  });

  // Insérer les trois lignes à partir de la ligne 4
  sheet.getRange(4, 1, 1, row1.length).setValues([row1]);
  sheet.getRange(5, 1, 1, row2.length).setValues([row2]);
  sheet.getRange(6, 1, 1, row3.length).setValues([row3]);

  // Mise en forme ligne 4 (entêtes)
  const headerRange = sheet.getRange(4, 1, 1, totalCols);
  headerRange.setFontFamily("Arial");
  headerRange.setFontSize(12);
  headerRange.setFontWeight("bold");
  headerRange.setFontColor("white");
  headerRange.setBackground("#073763");
  headerRange.setHorizontalAlignment("center");

  // Ligne 5
  const secondLine = sheet.getRange(5, 1, 1, totalCols);
  secondLine.setFontFamily("Arial");
  secondLine.setFontSize(10);
  secondLine.setFontColor("black");

  // Supprimer les colonnes à droite
  const maxCols = sheet.getMaxColumns();
  if (maxCols > totalCols) {
    sheet.deleteColumns(totalCols + 1, maxCols - totalCols);
  }

  // Supprimer les lignes après la 10e
  const maxRows = sheet.getMaxRows();
  if (maxRows > totalRows) {
    sheet.deleteRows(totalRows + 1, maxRows - totalRows);
  }

  // Figer jusqu'à la ligne 4 et cacher quadrillage
  sheet.setFrozenRows(4); // ❄️ On fige jusqu'à la ligne 4 (l'entête du tableau)
  sheet.setHiddenGridlines(true);

  // Couleur de l'onglet
  sheet.setTabColor("#3d85c6");

  extractHeadings(); // ⬅️ Lancement automatique après la création
  sheet.hideSheet(); // ⬅️ Et on la cache juste après
}

function generateDonneesStructureesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  if (!configSheet) throw new Error("Feuille 'Config' introuvable");

  const data = configSheet.getRange("B15:C" + configSheet.getLastRow()).getValues();
  const validTemplates = data.filter(([segment, url]) => segment && url);

  if (validTemplates.length === 0) {
    throw new Error("Aucun couple segment/URL trouvé dans Config!B15:C");
  }

  const sheetName = "Données structurées";
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  const row1 = [];
  const row2 = [];
  const row3 = [];

  validTemplates.forEach(([template, url]) => {
    const name = template.toLowerCase().trim();
    row1.push(`Template ${name} - Données structurées actuelles`);
    row1.push(`Template ${name} - Données structurées préconisées`);
    
    row2.push("Exemple pour :");
    row2.push(""); // Colonne préconisé vide

    row3.push(url);
    row3.push(""); // Colonne préconisé vide
  });

  const totalCols = row1.length;
  const totalRows = 50;

  // Insérer les trois lignes
  sheet.getRange(1, 1, 1, row1.length).setValues([row1]);
  sheet.getRange(2, 1, 1, row2.length).setValues([row2]);
  sheet.getRange(3, 1, 1, row3.length).setValues([row3]);


  // Mise en forme ligne 1 (entêtes)
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange.setFontFamily("Arial");
  headerRange.setFontSize(12);
  headerRange.setFontWeight("bold");
  headerRange.setFontColor("white");
  headerRange.setBackground("#073763");
  headerRange.setHorizontalAlignment("center");

  // Ligne 2
  const secondLine = sheet.getRange(2, 1, 1, totalCols);
  secondLine.setFontFamily("Arial");
  secondLine.setFontSize(10);
  secondLine.setFontColor("black");

  // Largeur & bordures ligne 1 (blanches)
  for (let col = 1; col <= totalCols; col++) {
    sheet.setColumnWidth(col, 500);

    if (col % 2 === 0) { // uniquement les colonnes 2, 4, 6, etc.
      // Ligne 1 : bordure blanche
      const rangeHeader = sheet.getRange(1, col);
      rangeHeader.setBorder(null, null, null, true, null, null, "#ffffff", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      rangeHeader.setHorizontalAlignment("center");

      // Lignes 2 à 50 : bordure noire
      const rangeBody = sheet.getRange(2, col, totalRows - 1, 1);
      rangeBody.setBorder(null, null, null, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

       // ➡️ Activation du retour automatique
        const entireColumn = sheet.getRange(1, col, totalRows, 1);
        entireColumn.setWrap(true);

        Logger.log(`Col ${col}: ➤ bordure droite + retour automatique appliqués`);
      } else {
        Logger.log(`Col ${col}: (aucune bordure)`);
      }
  }

  // Supprimer les colonnes à droite
  const maxCols = sheet.getMaxColumns();
  if (maxCols > totalCols) {
    sheet.deleteColumns(totalCols + 1, maxCols - totalCols);
  }

  // Supprimer les lignes après la 50e
  const maxRows = sheet.getMaxRows();
  if (maxRows > totalRows) {
    sheet.deleteRows(totalRows + 1, maxRows - totalRows);
  }

  // Figer la ligne 1 et cacher quadrillage
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(true);

  // Couleur de l'onglet
  sheet.setTabColor("#3d85c6");

    // Ajout des listes déroulantes dans les colonnes paires (préconisées) de la ligne 5 à 9
  const dropdownValues = [
    "Organization",
    "WebSite",
    "BreadcrumbList",
    "Product",
    "Article / BlogPosting",
    "FAQ / Answer",
    "Local Business",
    "WebPage",
    "ItemList"
  ];

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(dropdownValues)
    .setAllowInvalid(false)
    .build();

  for (let col = 2; col <= totalCols; col += 2) { // colonnes 2, 4, 6, etc.
    const range = sheet.getRange(5, col, 6, 1); // lignes 5 à 10
    range.setDataValidation(rule);
  }

  Logger.log("🎯 Listes déroulantes appliquées sur les colonnes préconisées (2, 4, ..., jusqu'à " + totalCols + ")")

  extractStructuredDataFromSheet();
  sheet.hideSheet(); // ⬅️ Et on la cache juste après
}


function extractHeadings() {
  const usernameDefault = 'sfteam';
  const passwordDefault = 'SF@Team17';

  let username = '';
  let password = '';
  
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Balisage hn");
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  const startRow = 8; // 🟢 Ligne de départ mise à 8

  // Colonnes A, C, E, etc.
  const columnsToClear = Array.from({ length: lastColumn }, (_, i) => i + 1).filter(c => c % 2 === 1);

  if (lastRow >= startRow) {
    columnsToClear.forEach(col => {
      const range = sheet.getRange(startRow, col, lastRow - startRow + 1, 1);
      range.clearContent();
    });
  }

  // Recherche de la dernière colonne avec une URL (ligne 6)
  const range = sheet.getRange(6, 1, 1, lastColumn);
  const values = range.getValues()[0];
  let lastColumnWithUrl = 0;

  for (let col = values.length - 1; col >= 0; col--) {
    if (values[col] !== '') {
      lastColumnWithUrl = col + 1;
      break;
    }
  }

  // URL dans A6, C6, etc.
  const urlCells = [];
  for (let col = 1; col <= lastColumnWithUrl; col += 2) {
    urlCells.push({ cell: sheet.getRange(6, col), col });
  }

  urlCells.forEach(({ cell, col }) => {
    const url = cell.getValue();
    if (!url) return;

    let htmlContent = '';
    let urlAuthSuccess = false;

    try {
      htmlContent = fetchUrlContent(url, '', '');
      urlAuthSuccess = true;
    } catch (e) {
      Logger.log(`⚠️ Sans auth - ${url} : ${e.message}`);
      if (e.message.includes("401")) {
        try {
          htmlContent = fetchUrlContent(url, usernameDefault, passwordDefault);
          urlAuthSuccess = true;
        } catch (e2) {
          Logger.log(`⚠️ Avec identifiants défaut - ${url} : ${e2.message}`);
          if (e2.message.includes("401")) {
            const ask = ui.alert("Authentification requise pour " + url, ui.ButtonSet.YES_NO);
            if (ask === ui.Button.YES) {
              username = ui.prompt("Nom d’utilisateur :").getResponseText();
              password = ui.prompt("Mot de passe :", ui.ButtonSet.OK_CANCEL).getResponseText();
              try {
                htmlContent = fetchUrlContent(url, username, password);
                urlAuthSuccess = true;
              } catch (e3) {
                ui.alert("❌ Authentification échouée pour : " + url);
              }
            }
          }
        }
      }
    }

    if (urlAuthSuccess && htmlContent) {
      const headings = getHeadings(htmlContent);
      let row = startRow;

      headings.forEach(heading => {
        const levelMatch = heading.match(/^<h([1-6])>/i);
        let formattedHeading = heading;
        if (levelMatch) {
          const level = parseInt(levelMatch[1], 10);
          if (level > 1) {
            formattedHeading = ' '.repeat((level - 1) * 3) + heading;
          }
        }
        sheet.getRange(row, col).setValue(formattedHeading);
        row++;
      });
    }
  });

  compareHeadings(); // Appel de comparaison
}

function compareHeadings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Balisage hn");
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    const startRow = 8;

    // Recherche de la dernière colonne avec URL (ligne 6 maintenant)
    const values = sheet.getRange(6, 1, 1, lastColumn).getValues()[0];
    let lastColumnWithUrl = 0;
    for (let col = values.length - 1; col >= 0; col--) {
      if (values[col] !== '') {
        lastColumnWithUrl = col + 1;
        break;
      }
    }

    for (let col = 1; col <= lastColumnWithUrl; col += 2) {
      const rangeA = sheet.getRange(startRow, col, lastRow - startRow + 1, 1);
      const rangeB = sheet.getRange(startRow, col + 1, lastRow - startRow + 1, 1);

      const valuesA = rangeA.getValues();
      const valuesB = rangeB.getValues();

      rangeA.setBackground(null);
      rangeB.setBackground(null);

      const originalA = valuesA.map(row => row[0]?.toString().trim() || '');
      const originalB = valuesB.map(row => row[0]?.toString().trim() || '');

      let remainingA = [...originalA];
      let remainingB = [...originalB];

      for (let i = 0; i < valuesA.length; i++) {
        const valueA = originalA[i];
        const valueB = originalB[i];

    if (valueA) {
      const indexInB = remainingB.indexOf(valueA);
      if (indexInB !== -1) {
        const cleanedA = valueA.replace(/^<h[1-6]>/i, '').trim();
        const cleanedB = remainingB[indexInB].replace(/^<h[1-6]>/i, '').trim();
        const tagA = valueA.match(/^<h[1-6]>/i)?.[0] || '';
        const tagB = remainingB[indexInB].match(/^<h[1-6]>/i)?.[0] || '';

        if (cleanedA === cleanedB && tagA.toLowerCase() !== tagB.toLowerCase()) {
          rangeA.getCell(i + 1, 1).setBackground('#fce5cd');
          rangeB.getCell(indexInB + 1, 1).setBackground('#fce5cd');
        } else {
          rangeA.getCell(i + 1, 1).setBackground('#d9ead3');
        }
        remainingB[indexInB] = null;
      } else {
        const contentOnlyA = valueA.replace(/^<h[1-6]>/i, '').trim();
        const matchContentInB = remainingB.findIndex(val => val && val.replace(/^<h[1-6]>/i, '').trim() === contentOnlyA);

        if (matchContentInB !== -1) {
          rangeA.getCell(i + 1, 1).setBackground('#fce5cd');
          rangeB.getCell(matchContentInB + 1, 1).setBackground('#fce5cd');
          remainingB[matchContentInB] = null;
        } else {
          rangeA.getCell(i + 1, 1).setBackground('#FFCFC9');
        }
      }
    }

    if (valueB) {
      const indexInA = remainingA.indexOf(valueB);
      if (indexInA !== -1) {
        const cleanedA = remainingA[indexInA].replace(/^<h[1-6]>/i, '').trim();
        const cleanedB = valueB.replace(/^<h[1-6]>/i, '').trim();
        const tagA = remainingA[indexInA].match(/^<h[1-6]>/i)?.[0] || '';
        const tagB = valueB.match(/^<h[1-6]>/i)?.[0] || '';

        if (cleanedA === cleanedB && tagA.toLowerCase() !== tagB.toLowerCase()) {
          rangeA.getCell(indexInA + 1, 1).setBackground('#fce5cd');
          rangeB.getCell(i + 1, 1).setBackground('#fce5cd');
        } else {
          rangeB.getCell(i + 1, 1).setBackground('#d9ead3');
        }
        remainingA[indexInA] = null;
      } else {
        const contentOnlyB = valueB.replace(/^<h[1-6]>/i, '').trim();
        const matchContentInA = remainingA.findIndex(val => val && val.replace(/^<h[1-6]>/i, '').trim() === contentOnlyB);

        if (matchContentInA !== -1) {
          rangeB.getCell(i + 1, 1).setBackground('#fce5cd');
          rangeA.getCell(matchContentInA + 1, 1).setBackground('#fce5cd');
          remainingA[matchContentInA] = null;
        } else {
          rangeB.getCell(i + 1, 1).setBackground('#FFCFC9');
        }
      }
    }
      }
    }

  
 // 🔽 Bordures verticales (ligne 3 = blanche, lignes 4+ = noire)
  for (let col = 1; col <= lastColumn; col++) {
    if (col % 2 === 0) {
      // Ligne 3 → bordure blanche
      const rangeHeader = sheet.getRange(3, col);
      rangeHeader.setBorder(null, null, null, true, null, null, "#ffffff", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      rangeHeader.setHorizontalAlignment("center");

      // Lignes 4 à lastRow → bordure noire
      const rangeBody = sheet.getRange(4, col, lastRow - 3, 1); // lastRow - (start at 4)
      rangeBody.setBorder(null, null, null, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }

  // Ajustement des largeurs
  for (let col = 1; col <= lastColumn; col++) {
    const values = sheet.getRange(1, col, lastRow, 1).getValues();
    let maxLen = 0;
    let longest = '';

    values.forEach(([val]) => {
      const raw = val ? val.toString().trim() : '';
      const len = raw.length;
      if (len > maxLen) {
        maxLen = len;
        longest = raw;
      }
    });

    const width = Math.max(500, Math.round(maxLen * 8.5));
    sheet.setColumnWidth(col, width);
    Logger.log(`📏 Col ${col}: "${longest}" → ${maxLen} chars → ${width}px`);
  }
}

function columnFromLetter(letter) {
  return letter.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
}

function fetchUrlContent(url, username, password) {
  var options = {};
  if (username && password) {
    var credentials = Utilities.base64Encode(username + ':' + password);
    options.headers = {
      'Authorization': 'Basic ' + credentials
    };
  }

  var response = UrlFetchApp.fetch(url, options);
  var content = response.getContentText();

  Logger.log('Contenu HTML récupéré pour l’URL: ' + url);
  return content;
}

function getHeadings(html) {
  var output = [];

  // Extraire le contenu entre les balises <body> et </body>
  var bodyMatch = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  if (!bodyMatch) {
    Logger.log("Aucune balise <body> trouvée.");
    return output;
  }
  
  var bodyContent = bodyMatch[1]; // Contenu du body

  // Expression régulière pour les balises <h1> à <h6>
  var re = /<(h[1-6])[^>]*>([\s\S]*?)<\/\1>/gi;
  var match;

  while (match = re.exec(bodyContent)) {
    var headingTag = match[1];
    var headingContent = match[2];

    // Normaliser le contenu des en-têtes (suppression des sauts de ligne et espaces multiples)
    var normalizedContent = decodeHtmlEntities(headingContent.replace(/\s+/g, ' ').trim());

    // Log si c'est un <h3>
    if (headingTag === 'h3') {
      Logger.log("Balise <h3> trouvée : " + normalizedContent);
    }

    // Extraction des sous-headings s'il y en a
    var nestedMatch;
    var nestedRe = /<(h[1-6])[^>]*>([\s\S]*?)<\/\1>/gi;
    while (nestedMatch = nestedRe.exec(normalizedContent)) {
      var nestedHeadingTag = nestedMatch[1];
      var nestedHeadingContent = nestedMatch[2];
      var nestedTextContent = decodeHtmlEntities(nestedHeadingContent.replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim());
      if (nestedTextContent) {
        output.push('<' + nestedHeadingTag + '>' + nestedTextContent + ' (Dans <' + headingTag + '>)');
      }
    }

    var mainTextContent = decodeHtmlEntities(normalizedContent.replace(nestedRe, '').replace(/<[^>]*>/g, '').trim());
    if (mainTextContent) {
      output.push('<' + headingTag + '>' + mainTextContent);
    } else {
      var imgMatch = headingContent.match(/<img[^>]+>/i);
      if (imgMatch) {
        var altTextMatch = imgMatch[0].match(/alt=["']([^"']*)["']/i);
        if (altTextMatch && altTextMatch[1].trim()) {
          output.push('<' + headingTag + '>(' + altTextMatch[1].trim() + ')');
        } else {
          output.push('<' + headingTag + '>Image sans Alt');
        }
      } else {
        output.push('<' + headingTag + '>Pas de texte');
      }
    }
  }
  return output;
}

function extractStructuredDataFromSheet() {
  var ui = SpreadsheetApp.getUi();
  var username = '';
  var password = '';
  var defaultUsername = 'sfteam';
  var defaultPassword = 'SF@Team17';

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Données structurées");

  var lastRow = sheet.getLastRow(); // 🔥 On prend la dernière ligne avec contenu
  var lastColumn = sheet.getLastColumn(); // 🔥 Dernière colonne utilisée

  // Colonnes spécifiques : A, C, E, G, I, K, M, O jusqu'à lastColumn
  for (var col = 1; col <= lastColumn; col += 2) { // 🔥 col=1(A), col=3(C), col=5(E)...
    var range = sheet.getRange(5, col, lastRow - 4, 1); // 🔥 De la ligne 5 jusqu'à la dernière
    range.clearContent();
    range.setBackground(null);
  }
  
  // Liste des cellules contenant les URL
  var urlCells = ['A3', 'C3', 'E3', 'G3', 'I3', 'K3', 'M3', 'O3'];
  urlCells.forEach(function(cell) {
    var url = sheet.getRange(cell).getValue();
    if (url) {
      try {
        var result = extractStructuredDataWithAuth(url, username, password, defaultUsername, defaultPassword);
        var types = result.types;
        var jsonLdBlocks = result.jsonLdBlocks;

        if (types.length > 0) {
          var cellColumn = columnFromLetter(cell.charAt(0));
          if (sheet.getMaxColumns() >= cellColumn) {
            var row = 5;
            types.forEach(function(type) {
              var range = sheet.getRange(row, cellColumn);
              range.setValue(type);
              if (type.includes(' - Microdata')) {
                range.setBackground('#FCE5CD');
              }
              row++;
            });
          }
        }

        // 🔥 Afficher les JSON-LD extraits avant la comparaison
        if (jsonLdBlocks.length > 0) {
          var cellColumn = columnFromLetter(cell.charAt(0));
          displayExtractedJsonLd(sheet, jsonLdBlocks, cellColumn);
        }

      } catch (e) {
        Logger.log("Erreur pour l’URL " + url + " → " + e.message);
      }
    }
  });

  // 🧠 Fixe les colonnes à 500px sans ajustement automatique
  var lastCol = sheet.getLastColumn();
  for (let col = 1; col <= lastCol; col++) {
    sheet.setColumnWidth(col, 500);
  }

  // 🔥 Comparaison finale
  compareStructuredData();
  finalFormatting();
}

function extractStructuredData(url, username, password) {
  const options = {};
  if (username && password) {
    const credentials = Utilities.base64Encode(username + ':' + password);
    options.headers = { 'Authorization': 'Basic ' + credentials };
  }

  const response = UrlFetchApp.fetch(url, options);
  const html = response.getContentText();
  const typesSet = new Set();
  const jsonLdBlocks = []; // 🔥 On initialise la liste pour les JSON-LD extraits

  // Extraction des JSON-LD
  const jsonLdRegex = /<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/g;
  const jsonLdMatches = html.match(jsonLdRegex);

  if (jsonLdMatches) {
    jsonLdMatches.forEach((match) => {
      let cleanJson = match
        .replace(/<script[^>]*type=["']application\/ld\+json["'][^>]*>|<\/script>/g, '')
        .trim();

      cleanJson = decodeHtmlEntitiesForStructuredData(cleanJson);
      cleanJson = cleanJson.replace(/#\\\//g, "#/");
      cleanJson = cleanJson.replace(/\\\//g, "/");
      cleanJson = cleanJson.replace(/(#)(\/)/g, '$1\\$2');

      // Vérification de longueur du JSON
      if (cleanJson.length < 50) {
        return;
      }

      try {
        const parsedJson = JSON.parse(cleanJson);

        const processType = (item) => {
          if (item['@type']) {
            if (Array.isArray(item['@type'])) {
              item['@type'].forEach(subType => typesSet.add(subType));
            } else {
              typesSet.add(item['@type']);
            }
          }
        };

        if (Array.isArray(parsedJson)) {
          parsedJson.forEach(obj => processType(obj));
        } else if (parsedJson['@graph']) {
          parsedJson['@graph'].forEach(item => processType(item));
        } else {
          processType(parsedJson);
        }

        // 🔥 On sauvegarde le JSON-LD complet propre
        jsonLdBlocks.push(parsedJson);

      } catch (e) {
        Logger.log(`❌ Erreur JSON.parse : ${e.message}`);
      }
    });
  }

  // Extraction des Microdonnées
  const microdataRegex = /itemtype="http(s)?:\/\/schema\.org\/([^"]+)"/g;
  const microdataMatches = html.match(microdataRegex);
  if (microdataMatches) {
    microdataMatches.forEach(match => {
      const type = match.split('schema.org/')[1].replace(/"/g, '');
      typesSet.add(type + ' - Microdata');
    });
  }

  // 🔥 Retourne à la fois les types ET les JSON-LD extraits
  return {
    types: Array.from(typesSet),
    jsonLdBlocks: jsonLdBlocks
  };
}

function displayExtractedJsonLd(sheet, jsonLdBlocks, startColumn) {
  let currentRow = 15;

  const numColumns = 2;
  sheet.getRange(currentRow, startColumn, 500, numColumns).clearContent().clearFormat();

  jsonLdBlocks.forEach((jsonBlock, index) => {
    if (!jsonBlock['@type'] || typeof jsonBlock['@type'] !== 'string' || jsonBlock['@type'].includes('Microdata')) {
      return; // On ignore les Microdata
    }

    const titleCell = sheet.getRange(currentRow, startColumn, 1, numColumns);
    titleCell.merge();
    titleCell.setValue(`${jsonBlock['@type']}`);
    titleCell.setBackground('#073763');
    titleCell.setFontColor('white');
    titleCell.setFontWeight('bold');
    titleCell.setFontSize(10);
    titleCell.setHorizontalAlignment('center');
    titleCell.setVerticalAlignment('middle');

    currentRow += 2;

    const jsonCell = sheet.getRange(currentRow, startColumn);
    jsonCell.setValue(JSON.stringify(jsonBlock, null, 2)); // 🔥 plus de replace
    jsonCell.setWrap(true);
    jsonCell.setVerticalAlignment('top');

    currentRow += 2;
  });
}

function compareStructuredData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Données structurées");
  const lastRow = sheet.getLastRow();
  const startRow = 5;
  const lastColumn = sheet.getLastColumn();

  for (let col = 1; col <= lastColumn; col += 2) {
    const rangeExtracted = sheet.getRange(startRow, col, lastRow - startRow + 1, 1);     // Colonnes A, C, E, etc.
    const rangeExpected = sheet.getRange(startRow, col + 1, lastRow - startRow + 1, 1);   // Colonnes B, D, F, etc.

    const extractedValues = rangeExtracted.getValues();
    const expectedValues = rangeExpected.getValues();

    rangeExpected.setBackground(null); // 🔵 Seulement reset des couleurs dans les colonnes attendues (B, D, F...)

    const originalExtracted = extractedValues.map(row => row[0]?.toString().trim() || '');
    const originalExpected = expectedValues.map(row => row[0]?.toString().trim() || '');

    let remainingExtracted = [...originalExtracted];
    let remainingExpected = [...originalExpected];

    for (let i = 0; i < extractedValues.length; i++) {
      const valueExtracted = originalExtracted[i];
      const valueExpected = originalExpected[i];

      if (valueExpected) {
        const indexInExtracted = remainingExtracted.indexOf(valueExpected);
        if (indexInExtracted !== -1) {
          rangeExpected.getCell(i + 1, 1).setBackground('#d9ead3'); // ✅ Vert clair si trouvé
          remainingExtracted[indexInExtracted] = null;
        } else {
          rangeExpected.getCell(i + 1, 1).setBackground('#FFCFC9'); // ❌ Rouge clair si non trouvé
        }
      }
    }
  }
}

function extractStructuredDataWithAuth(url, username, password, defaultUsername, defaultPassword) {
  var ui = SpreadsheetApp.getUi(); // Interface utilisateur
  var result = { types: [], jsonLdBlocks: [] };
  var success = false; // Indique si une méthode a réussi

  // Essai initial sans authentification
  try {
    result = extractStructuredData(url, '', '');
    success = true; // Succès sans authentification
  } catch (e) {
    Logger.log("Erreur sans authentification : " + e.message);
    if (e.message.indexOf("401") > -1) {
      Logger.log("Tentative avec identifiants par défaut.");

      // Essai avec identifiants par défaut
      try {
        result = extractStructuredData(url, defaultUsername, defaultPassword);
        success = true; // Succès avec identifiants par défaut
        username = defaultUsername;
        password = defaultPassword;
      } catch (e) {
        Logger.log("Erreur avec identifiants par défaut : " + e.message);
        if (e.message.indexOf("401") > -1) {
          Logger.log("Demande des identifiants à l'utilisateur.");

          // Demander des identifiants à l'utilisateur
          var response = ui.alert(
            'L\'authentification est requise pour accéder à l\'URL. Voulez-vous fournir des identifiants ?',
            ui.ButtonSet.YES_NO
          );
          if (response == ui.Button.YES) {
            var usernameResponse = ui.prompt('Entrez votre nom d’utilisateur :');
            username = usernameResponse.getResponseText();
            var passwordResponse = ui.prompt('Entrez votre mot de passe :', ui.ButtonSet.OK_CANCEL);
            password = passwordResponse.getResponseText();

            // Réessai avec les identifiants fournis
            try {
              result = extractStructuredData(url, username, password);
              success = true;
            } catch (e) {
              Logger.log("Erreur avec identifiants utilisateur : " + e.message);
              ui.alert("Échec de l'authentification. Impossible de récupérer les données pour l'URL : " + url);
            }
          }
        }
      }
    }
  }

  if (!success) {
    Logger.log("Impossible de récupérer les données pour l'URL : " + url);
    return { types: [], jsonLdBlocks: [] };
  }

  return result;
}

function finalFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Données structurées");
  const lastColumn = sheet.getLastColumn();
  const maxRows = sheet.getMaxRows();

  // 🔍 Trouver la dernière ligne non vide sur toutes les colonnes utilisées
  let lastRowWithContent = sheet.getRange('A:Z').getValues()
    .reduce((maxRow, row, index) => {
      return row.some(cell => cell !== '') ? index + 1 : maxRow;
    }, 0);

  if (lastRowWithContent > 0 && lastRowWithContent < maxRows) {
    // ✂️ Supprimer toutes les lignes vides en-dessous
    sheet.deleteRows(lastRowWithContent + 1, maxRows - lastRowWithContent);
  }

  // 🔽 Puis ajouter bordures verticales sur colonnes paires uniquement (B, D, F, etc.)
  for (let col = 1; col <= lastColumn; col++) {
    if (col % 2 === 0) { // Colonnes 2, 4, 6, etc.
      for (let row = 15; row <= lastRowWithContent; row++) {
        const isWhiteLine = (row - 15) % 4 === 0; // ligne 15, 19, 23, etc.
        const range = sheet.getRange(row, col);

        // 🎨 Bordures verticales
        range.setBorder(
          null, null, null, true, null, null,
          isWhiteLine ? "#ffffff" : "#000000",
          SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        );

        // 🔥 Alignement vertical en haut pour lignes 17, 21, 25, etc.
        if ((row - 17) % 4 === 0) {
          range.setVerticalAlignment('top');
        } else {
          range.setVerticalAlignment('middle'); // ✅ Autrement, on remet vertical centré (important pour ne pas tout décaler)
        }
      }
    }
  }
}
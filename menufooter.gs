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

function createMenuFooter(data) {
  /*
   * Crée ou met à jour la feuille "Menu & Footer" avec les données structurées
   * du menu et du footer, en appliquant un formatage spécifique.
   */
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Menu & Footer");

    // Sauvegarde de la structure actuelle pour la persistance
    PropertiesService.getDocumentProperties().setProperty('menuFooterData', JSON.stringify(data));

    if (sheet) {
      sheet.clear();
    } else {
      sheet = ss.insertSheet("Menu & Footer");
    }
    
    // Application de la couleur à l'onglet de la feuille
    sheet.setTabColor("#f39c12");

    let finalRows = [];
    let bandingRangesInfo = []; 

    // Traitement de la section MENU
    if (data.menu && data.menu.length > 0) {
      const menuRows = processSectionData(data.menu);
      finalRows.push(['MENU', '', '', '', '']);
      const columnHeaderRowIndex = finalRows.length + 1;
      finalRows.push(['Entrée', 'Niveau', 'Libellé', 'URL', 'Note']);
      finalRows = finalRows.concat(menuRows);
      const endDataRow = finalRows.length;
      if (columnHeaderRowIndex <= endDataRow) {
        bandingRangesInfo.push({ start: columnHeaderRowIndex, end: endDataRow });
      }
    }

    // Ligne vide de séparation
    if (finalRows.length > 0 && data.footer && data.footer.length > 0) {
      finalRows.push(['', '', '', '', '']); 
    }

    // Traitement de la section FOOTER
    if (data.footer && data.footer.length > 0) {
      const footerRows = processSectionData(data.footer);
      finalRows.push(['FOOTER', '', '', '', '']);
      const columnHeaderRowIndex = finalRows.length + 1;
      finalRows.push(['Entrée', 'Niveau', 'Libellé', 'URL', 'Note']);
      finalRows = finalRows.concat(footerRows);
      const endDataRow = finalRows.length;
       if (columnHeaderRowIndex <= endDataRow) {
        bandingRangesInfo.push({ start: columnHeaderRowIndex, end: endDataRow });
      }
    }

    if (finalRows.length > 0) {
      sheet.getRange(1, 1, finalRows.length, 5).setValues(finalRows);
      formatMenuFooterSheet(sheet, bandingRangesInfo);
      cleanupSheet(sheet, 5);
    }

  } catch (e) {
    Logger.log('Erreur dans createMenuFooter: ' + e.message + ' Stack: ' + e.stack);
    throw new Error('Une erreur est survenue lors de la création de la feuille Menu & Footer: ' + e.message);
  }
}

function processSectionData(sectionData) {
  /*
   * Traite les données hiérarchiques d'une section (menu ou footer) pour les
   * transformer en un tableau de lignes à plat.
   */
  const flattenedRows = [];
  
  function flatten(items, level, l1EntryName) {
    items.forEach(item => {
      const currentL1Entry = level === 1 ? (item.label || item.Nom) : l1EntryName;
      const finalUrl = item.nonClickable ? '#' : (item.url || '#');
      const row = [
        currentL1Entry,
        level,
        (level > 1 ? '↳ ' : '') + (item.label || item.Nom),
        finalUrl,
        item.nonClickable ? 'Non cliquable' : ''
      ];
      flattenedRows.push(row);

      if (item.children && item.children.length > 0) {
        flatten(item.children, level + 1, currentL1Entry);
      }
    });
  }

  flatten(sectionData, 1, '');
  return flattenedRows;
}

function formatMenuFooterSheet(sheet, bandingRangesInfo) {
    /* * Applique tout le formatage à la feuille : polices, couleurs, largeurs,
     * fusion, banding, etc.
     */
    const HEADER_BACKGROUND_COLOR = '#073763';
    const BANDING_COLOR_1 = '#FFFFFF';
    const BANDING_COLOR_2 = '#F3F3F3';

    const lastRow = sheet.getLastRow();
    if (lastRow === 0) return;

    sheet.setHiddenGridlines(true);

    // Formatage global de la police et des alignements
    const contentRange = sheet.getRange("A1:E" + lastRow);
    contentRange.setFontFamily("Arial").setFontSize(11).setFontColor("black").setVerticalAlignment('middle');
    sheet.getRange("B1:B" + lastRow).setHorizontalAlignment('center');
    sheet.getRange("E1:E" + lastRow).setHorizontalAlignment('center');

    // Formatage des en-têtes principaux (MENU / FOOTER)
    for (let i = 1; i <= lastRow; i++) {
        const firstCell = sheet.getRange(i, 1).getValue().toString().toUpperCase();
        if (firstCell === 'MENU' || firstCell === 'FOOTER') {
            const headerRange = sheet.getRange(i, 1, 1, 5);
            headerRange.merge().setBackground(HEADER_BACKGROUND_COLOR).setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
        }
    }

    // Application du banding et formatage des en-têtes de colonnes
    bandingRangesInfo.forEach(info => {
      const columnHeaderRange = sheet.getRange(info.start, 1, 1, 5);
      columnHeaderRange.setBackground(HEADER_BACKGROUND_COLOR).setFontColor('white').setFontWeight('bold');

      // Banding pour les lignes de données
      for (let i = info.start + 1; i <= info.end; i++) {
        const rowRange = sheet.getRange(i, 1, 1, 5);
        if ((i - (info.start + 1)) % 2 === 0) { // Ligne paire de données
          rowRange.setBackground(BANDING_COLOR_1);
        } else { // Ligne impaire de données
          rowRange.setBackground(BANDING_COLOR_2);
        }
      }
    });

    // Définition de la largeur des colonnes
    sheet.setColumnWidth(1, 200); // A
    sheet.setColumnWidth(2, 100); // B
    sheet.setColumnWidth(3, 200); // C
    sheet.autoResizeColumn(4);    // D
    sheet.setColumnWidth(5, 150); // E
}

function cleanupSheet(sheet, lastDataColumnIndex) {
  /*
   * Supprime les lignes et colonnes vides à la fin d'une feuille.
   */
  const maxRows = sheet.getMaxRows();
  const lastRow = sheet.getLastRow();
  if (maxRows > lastRow) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }

  const maxCols = sheet.getMaxColumns();
  if (maxCols > lastDataColumnIndex) {
    sheet.deleteColumns(lastDataColumnIndex + 1, maxCols - lastDataColumnIndex);
  }
}

function getInitialMenuFooterData() {
  /* On suppose que la fonction getArborescenceForMenuFooter() existe déjà
   * car elle est appelée par la version originale du HTML.
   */
  const arboPages = getArborescenceForMenuFooter(); 
  
  const savedStateJson = PropertiesService.getDocumentProperties().getProperty('menuFooterData');
  const savedState = savedStateJson ? JSON.parse(savedStateJson) : { menu: [], footer: [] };
  
  return {
    arboPages: arboPages,
    savedState: savedState
  };
}
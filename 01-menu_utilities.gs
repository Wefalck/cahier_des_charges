function onOpen() {
  var userEmail = Session.getActiveUser().getEmail(); // Récupère l'e-mail de l'utilisateur
  var authorizedUsers = getAuthorizedUsers(); // Liste des e-mails autorisés

  // Si l'utilisateur n'est pas autorisé, on quitte sans créer le menu
  if (!authorizedUsers.includes(userEmail)) {
    return;
  }

  // Création du menu personnalisé uniquement pour les utilisateurs autorisés
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ Admin SF')
    .addItem('📝 Prise de brief', 'afficherPriseDeBrief')
    .addItem('📝 Définir les templates de pages', 'afficherTemplateModal')
    .addSubMenu(ui.createMenu('📊 Récuperer des données')
      .addItem('📊 Data crawl + positionnement', 'get_data')
      .addItem('🔄 MAJ inventaire', 'completerInventaireStratEtPositionnement'))
    .addSubMenu(ui.createMenu('Arborescence, Menu & Footer')
      .addItem('🌳 Créer arborescence', 'creerArborescence'))
    .addSubMenu(ui.createMenu('🔎 Audit')
      .addItem('🔍 Vérifier les sous-domaines', 'check_subdomain'))
    .addSubMenu(ui.createMenu('📌 Stratégie de positionnement')
      .addItem('✔️ Checker les préconisations éditoriales', 'checkStratPos'))
    .addSubMenu(ui.createMenu('🧱 Balisage <hn>')
      .addItem('📝 Extraire les structures', 'extractHeadings')
      .addItem('🔍 Comparer les structures', 'compareHeadings'))
    .addSubMenu(ui.createMenu('🗂️ Données structurées')
      .addItem('📊 Extraire les données structurées', 'extractStructuredDataFromSheet')
      .addItem('🔍 Comparer les données structurées', 'compareStructuredData'))
    .addSubMenu(ui.createMenu('🛡️ Protection')
      .addItem('🔒 Protéger les feuilles masquées', 'protectHiddenSheetsAndAudit')
      .addItem('🔓 Déprotéger toutes les feuilles', 'deprotection_feuilles'))
    .addToUi();
}

function onEdit(e) {
  const sheet = e.range.getSheet();
  const editedCol = e.range.getColumn();
  const editedRow = e.range.getRow();

  // On agit uniquement dans la feuille "Audit", colonne D
  if (sheet.getName() !== "Audit" || editedCol !== 4) return;

  const newValue = e.value;
  if (!["Passable", "Échoué"].includes(newValue)) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = sheet.getRange(editedRow, 7).getValue().toString().trim(); // Colonne G

  if (!sheetName) return;

  const targetSheet = ss.getSheetByName(sheetName);
  const auditSheet = ss.getSheetByName("Audit");

  if (!targetSheet) {
    Logger.log(`❌ Feuille "${sheetName}" introuvable.`);
    return;
  }

  if (!targetSheet.isSheetHidden()) {
    Logger.log(`ℹ️ Feuille "${sheetName}" déjà visible.`);
    return;
  }

  // On affiche la feuille et la déplace juste après "Audit"
  targetSheet.showSheet();

  const allSheets = ss.getSheets();
  const auditIndex = allSheets.findIndex(s => s.getName() === "Audit");

  if (auditIndex !== -1) {
    ss.setActiveSheet(targetSheet);
    ss.moveActiveSheet(auditIndex + 1); // Déplacement juste après "Audit"
    Logger.log(`✅ Feuille "${sheetName}" démasquée et déplacée.`);
  }

  // Réorganiser les feuilles visibles (le placement sera automatique)
  reordonnerFeuillesVisibles();

  // Activer la feuille ciblée une fois qu’elle a été repositionnée
  ss.setActiveSheet(targetSheet);
  Logger.log(`📌 Feuille "${sheetName}" activée après réorganisation.`);
}

function decodeHtmlEntities(str) {
  if (!str) return '';
  
  // On décode toutes les entités HTML connues
  const entities = {
    '&amp;': '&',
    '&lt;': '<',
    '&gt;': '>',
    '&quot;': '"',
    '&apos;': '\'',
    '&#39;': '\'',
    '&#x27;': '\''
  };
  
  str = str.replace(/&[a-zA-Z0-9#]+;/g, (match) => entities[match] || match);

  // Ensuite on traite les entités numériques
  str = str.replace(/&#(\d+);/g, (match, dec) => String.fromCharCode(dec));
  str = str.replace(/&#x([0-9A-Fa-f]+);/g, (match, hex) => String.fromCharCode(parseInt(hex, 16)));

  return str;
}

function decodeHtmlEntitiesForStructuredData(str) {
  if (!str) return '';

  // Décodage spécial : on supprime les &quot;
  const entities = {
    '&amp;': '&',
    '&lt;': '<',
    '&gt;': '>',
    '&apos;': '\'',
    '&#39;': '\'',
    '&#x27;': '\''
    // Pas de &quot; ici volontairement
  };

  // Remplacement des entités connues (sauf &quot;)
  str = str.replace(/&[a-zA-Z0-9#]+;/g, (match) => {
    if (match === '&quot;') {
      return ''; // 🔥 Supprimer les &quot; directement
    }
    return entities[match] || match;
  });

  // Traitement des entités numériques
  str = str.replace(/&#(\d+);/g, (match, dec) => String.fromCharCode(dec));
  str = str.replace(/&#x([0-9A-Fa-f]+);/g, (match, hex) => String.fromCharCode(parseInt(hex, 16)));

  return str;
}

function reordonnerFeuillesVisibles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const toutesFeuilles = ss.getSheets();

  const ordreCible = [
    "Suivi",
    "Balisage template",
    "Inventaire",
    "Arborescence",
    "Crawl - Préprod",
    "Sous-domaine",
    "URL",
    "Automatisation-Redirection",
    "Automatisation-Balises",
    "Balises SEO",
    "Filtre",
    "Menu de navigation",
    "Balisage hn",
    "Données structurées",
    "Robots.txt",
    "Hreflang",
    "Pagination",
    "Fil d'Ariane",
    "Maillage",
    "5XX",
    "Lien interne",
    "Sitemap",
    "Directive",
    "Canonique",
    "Title",
    "h1",
    "Meta description",
    "OpenGraph",
    "Image",
    "Vidéo",
    "Contenu",
    "Mobile",
    "Web performance",
    "Positionnement",
    "Stratégie de positionnement",
    "Config"
  ];

  // Filtrer les feuilles visibles uniquement
  const visibles = toutesFeuilles.filter(s => !s.isSheetHidden());
  const nonCiblées = visibles.filter(s => !ordreCible.includes(s.getName()));

  // Ordre final : feuilles de l’ordre défini + autres visibles
  const ordreFinal = [
    ...ordreCible.filter(nom => visibles.some(s => s.getName() === nom)),
    ...nonCiblées.map(s => s.getName())
  ];

  Logger.log("🔃 Réorganisation des feuilles visibles :");
  Logger.log(ordreFinal.join(" → "));

  ordreFinal.forEach((nomFeuille, i) => {
    const feuille = ss.getSheetByName(nomFeuille);
    if (feuille) {
      ss.setActiveSheet(feuille);
      ss.moveActiveSheet(i + 1); // 📌 1-based
    }
  });
}

function extractDomain(url) {
  var match = url.match(/^https?:\/\/([^\/]+)/i); // Cette regex extrait le domaine avec le protocole
  if (match) {
    var hostname = match[1]; // Récupère le domaine complet (avec sous-domaines)
    var domain = hostname;

    if (hostname != null) {
      var parts = hostname.split('.').reverse(); // Découpe le domaine en parties
      if (parts != null && parts.length > 1) {
        domain = parts[1] + '.' + parts[0];  // Concatène les deux dernières parties pour le domaine
        // Gère les cas spéciaux comme .co.uk
        if (hostname.toLowerCase().indexOf('.co.uk') != -1 && parts.length > 2) {
          domain = parts[2] + '.' + domain;
        }
      }
    }
    return domain;
  }
  return null; // Retourne null si l'URL n'est pas valide
}

function extractFullDomain(url) {
  var match = url.match(/^https?:\/\/([^\/]+)/i); // Cette regex extrait le domaine avec le protocole
  if (match) {
    return match[1]; // Retourne le domaine complet avec les sous-domaines
  }
  return null; // Retourne null si l'URL n'est pas valide
}

function afficherProprietesDocument() {
  const proprietes = PropertiesService.getDocumentProperties().getProperties();

  console.log("[INFO] Propriétés du document :", proprietes);
}

function resetDocumentProperties() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteAllProperties();
  console.log("[INFO] Toutes les propriétés du document ont été supprimées.");
}

function normaliserDomaineBrut(texte) {
  if (!texte) return "";

  let propre = texte.trim().toLowerCase();

  // Supprime le protocole si présent
  propre = propre.replace(/^https?:\/\//, '');

  // Supprime tout ce qui suit un slash (ex: /fr/, /path)
  propre = propre.split('/')[0];

  const parts = propre.split('.');

  if (parts.length < 2) return "";

  const last = parts[parts.length - 1];
  const secondLast = parts[parts.length - 2];
  let domaine = secondLast + '.' + last;

  // Gestion des domaines en .co.uk, .com.br, etc.
  if ((last === "uk" && secondLast === "co") || (last === "br" && secondLast === "com")) {
    if (parts.length >= 3) {
      domaine = parts[parts.length - 3] + '.' + secondLast + '.' + last;
    }
  }

  return domaine;
}

function protectHiddenSheetsAndAudit() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var allowedEditors = getAuthorizedUsers(); // Appel de getAuthorizedUsers

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (sheet.isSheetHidden() || sheet.getName() === "Audit") {
      var protection = sheet.protect().setDescription(sheet.getName() + ' Protégé');
      // Remove all editors from the protection
      protection.removeEditors(protection.getEditors());
      // Add specific editors
      protection.addEditors(allowedEditors);

      if (sheet.getName() === "Audit") {
        // Unprotect columns K, L, N for "Audit" sheet
        var unprotected = [
          sheet.getRange('K:K'),
          sheet.getRange('L:L'),
          sheet.getRange('N:N')
        ];
        protection.setUnprotectedRanges(unprotected);
      }
      
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }
    }
  }
}

function deprotection_feuilles() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for ( var i = 0 ; i<sheets.length ; i++) {
  var sheet = sheets[i];
  var protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  if (protection && protection.canEdit()) {
    protection.remove();
    }
  }
}

function updateSheetVisibility() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userEmail = Session.getActiveUser().getEmail();

    var authorizedUsers = getAuthorizedUsers(); // Appel de getAuthorizedUsers

    // Onglets restreints
    var restrictedSheets = ["Crawl", "Crawl - Image", "Crawl - Sitemap", "Config classeur", "Identité"];

    // Vérifier si l'utilisateur est autorisé
    if (!authorizedUsers.includes(userEmail)) {
      ss.getSheets().forEach(function(sheet) {
        var sheetName = sheet.getName();

        // Masquer l'onglet s'il est restreint
        if (restrictedSheets.includes(sheetName)) {
          sheet.hideSheet();
        }
      });
    }
}

function getAuthorizedUsers() {
  return [
    "philippe.vesin@search-factory.fr",
    "lea.deshayes@search-factory.fr",
    "benjamin.gennequin@search-factory.fr",
    "achille.catel@search-factory.fr",
    "ronan.cassin@search-factory.fr",
    "claire.chamaillard@search-factory.fr",
    "quentin.pareyn@search-factory.fr",
    "robin.ansaldi@search-factory.fr"
  ];
}

function parseCSV(content, delimiter = ",") {
  const rows = [];
  let current = '';
  let inQuotes = false;
  let row = [];

  for (let i = 0; i < content.length; i++) {
    const char = content[i];
    const nextChar = content[i + 1];

    if (char === '"' && inQuotes && nextChar === '"') {
      current += '"';
      i++; // guillemet échappé
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === delimiter && !inQuotes) {
      row.push(current);
      current = '';
    } else if ((char === '\n' || char === '\r') && !inQuotes) {
      if (current || row.length > 0) {
        row.push(current);
        rows.push(row.map(c => c.trim()));
        row = [];
        current = '';
      }
    } else {
      current += char;
    }
  }

  if (current || row.length > 0) {
    row.push(current);
    rows.push(row.map(c => c.trim()));
  }

  return rows;
}
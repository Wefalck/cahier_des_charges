function afficherPriseDeBrief() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('PriseDeBrief')
    .setWidth(600)
    .setHeight(800);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Prise de brief - Refonte / Création');
}

function enregistrerPriseDeBrief(donnees) {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleConfig = classeur.getSheetByName('Config');

  if (!feuilleConfig) {
    SpreadsheetApp.getUi().alert("Erreur : La feuille 'Config' est introuvable.");
    return;
  }

  if (donnees.client) feuilleConfig.getRange('B3').setValue(donnees.client);
  if (donnees.typeProjet) {
    feuilleConfig.getRange('B6').setValue(donnees.typeProjet);
    PropertiesService.getDocumentProperties().setProperty('typeProjet', donnees.typeProjet);
  }

  if (donnees.registrar) feuilleConfig.getRange('C9').setValue(donnees.registrar);
  if (donnees.hebergeur) feuilleConfig.getRange('C10').setValue(donnees.hebergeur);
  if (donnees.serveur) feuilleConfig.getRange('C11').setValue(donnees.serveur);

  if (donnees.cmsActuel) feuilleConfig.getRange('C14').setValue(donnees.cmsActuel);
  if (donnees.nouveauCms) {
    const value = donnees.nouveauCms === "Pas de changement"
      ? feuilleConfig.getRange('C14').getValue()
      : donnees.nouveauCms;
    feuilleConfig.getRange('C15').setValue(value);
  }

  if (donnees.domaineActuel) {
    const domaineActuel = normaliserDomaineBrut(donnees.domaineActuel);
    feuilleConfig.getRange('C19').setValue(domaineActuel);
  }

  if (donnees.sousDomaineActuel) {
    feuilleConfig.getRange('C20').setValue(donnees.sousDomaineActuel);
  }

  if (donnees.domaineNouveau) {
    const domaineNouveau = donnees.domaineNouveau === "Pas de changement"
      ? feuilleConfig.getRange('C19').getValue()
      : normaliserDomaineBrut(donnees.domaineNouveau);
    feuilleConfig.getRange('C22').setValue(domaineNouveau);
  }

  if (donnees.sousDomaineNouveau) {
    const value = donnees.sousDomaineNouveau === "Pas de changement"
      ? feuilleConfig.getRange('C20').getValue()
      : donnees.sousDomaineNouveau;
    feuilleConfig.getRange('C23').setValue(value);
  }

  if (donnees.finUrlActuel) feuilleConfig.getRange('C27').setValue(donnees.finUrlActuel);
  if (donnees.sousRepertoireActuel) feuilleConfig.getRange('C28').setValue(donnees.sousRepertoireActuel);
  if (donnees.finUrlNouveau) {
    const value = donnees.finUrlNouveau === "Pas de changement"
      ? feuilleConfig.getRange('C27').getValue()
      : donnees.finUrlNouveau;
    feuilleConfig.getRange('C30').setValue(value);
  }
  if (donnees.sousRepertoireNouveau) {
    feuilleConfig.getRange('C31').setValue(donnees.sousRepertoireNouveau);
  }


  if (donnees.urlPreprod) feuilleConfig.getRange('C34').setValue(donnees.urlPreprod);
  if (donnees.blocagePreprod) feuilleConfig.getRange('C35').setValue(donnees.blocagePreprod);
  if (donnees.doubleSecurite) feuilleConfig.getRange('C36').setValue(donnees.doubleSecurite);

  if (donnees.utilisationCdn) feuilleConfig.getRange('C39').setValue(donnees.utilisationCdn);
  if (donnees.cdn) feuilleConfig.getRange('C40').setValue(donnees.cdn);

  if (donnees.multilingue) feuilleConfig.getRange('C43').setValue(donnees.multilingue);
  if (donnees.seoFiltres) feuilleConfig.getRange('C44').setValue(donnees.seoFiltres);
  if (donnees.video) feuilleConfig.getRange('C45').setValue(donnees.video);
  if (donnees.partageSocial) feuilleConfig.getRange('C46').setValue(donnees.partageSocial);

  if (donnees.chefSeo) feuilleConfig.getRange('C49').setValue(donnees.chefSeo);
  if (donnees.chefWeb) feuilleConfig.getRange('C50').setValue(donnees.chefWeb);
  if (donnees.developpeur) feuilleConfig.getRange('C51').setValue(donnees.developpeur);
  if (donnees.adminServeur) feuilleConfig.getRange('C52').setValue(donnees.adminServeur);
  if (donnees.graphiste) feuilleConfig.getRange('C53').setValue(donnees.graphiste);
  if (donnees.integrateurMaquette) feuilleConfig.getRange('C54').setValue(donnees.integrateurMaquette);
  if (donnees.integrateurContenu) feuilleConfig.getRange('C55').setValue(donnees.integrateurContenu);
  if (donnees.integrateurSeo) feuilleConfig.getRange('C56').setValue(donnees.integrateurSeo);
  if (donnees.gestionDomaine) feuilleConfig.getRange('C57').setValue(donnees.gestionDomaine);
  if (donnees.gestionHebergement) feuilleConfig.getRange('C58').setValue(donnees.gestionHebergement);
  if (donnees.gestionRedirections) feuilleConfig.getRange('C59').setValue(donnees.gestionRedirections);

  // --- Construction des URLs avec gestion "Pas de changement"
  const props = PropertiesService.getDocumentProperties();

  // Domaine actuel
  const domaineActuel = feuilleConfig.getRange('C19').getValue();
  const sousDomaineActuel = feuilleConfig.getRange('C20').getValue();
  const urlActuel = construireUrlActuelNouveau(sousDomaineActuel, domaineActuel);
  if (urlActuel) {
    feuilleConfig.getRange('C62').setValue(urlActuel);
    props.setProperty('urlActuel', urlActuel);
  }

  // Domaine nouveau
  const domaineNouveau = donnees.domaineNouveau === "Pas de changement"
    ? domaineActuel
    : normaliserDomaineBrut(donnees.domaineNouveau);
  const sousDomaineNouveau = donnees.sousDomaineNouveau === "Pas de changement"
    ? sousDomaineActuel
    : donnees.sousDomaineNouveau;
  const urlNouveau = construireUrlActuelNouveau(sousDomaineNouveau, domaineNouveau);
  if (urlNouveau) {
    feuilleConfig.getRange('C63').setValue(urlNouveau);
    props.setProperty('urlNouveau', urlNouveau);
  }

  // Préprod
  const urlPreprod = nettoyerUrlPreprod(donnees.urlPreprod);
  if (urlPreprod) {
    feuilleConfig.getRange('C64').setValue(urlPreprod);
    props.setProperty('urlPreprod', urlPreprod);
  }

  // --- Ajout d'autres propriétés utiles
  if (donnees.client) {
    props.setProperty('client', donnees.client);
  }

  appliquerNettoyageFonctionnalites(donnees);

  if (donnees.typeProjet === "Création") {
    const feuillesASupprimer = ["Crawl - Prod", "Positionnement"];
    feuillesASupprimer.forEach(nom => {
      const feuille = classeur.getSheetByName(nom);
      if (feuille) classeur.deleteSheet(feuille);
    });
  }
}

function construireUrlActuelNouveau(sousDomaine, domaine) {
  if (!domaine) return "";

  // Nettoyage du domaine
  domaine = domaine.trim().toLowerCase().replace(/^https?:\/\//, '');
  domaine = domaine.split('/')[0]; // Retire les chemins

  // Reconstruit l'URL complète
  let url = "https://";
  if (sousDomaine && sousDomaine.toLowerCase() !== "sans") {
    url += sousDomaine + ".";
  }
  url += domaine;

  return url;
}

function nettoyerUrlPreprod(url) {
  if (!url) return "";

  let net = url.trim().toLowerCase().replace(/^https?:\/\//, '');
  net = net.split('/')[0]; // Supprime tout chemin

  return "https://" + net;
}

function appliquerNettoyageFonctionnalites(donnees) {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();

  // ----------------------------
  // MULTILINGUE
  if (donnees.multilingue === "Non") {
    try {
      const feuilleSuivi = classeur.getSheetByName("Suivi");
      if (feuilleSuivi) {
        const lignes = feuilleSuivi.getRange("B1:B").getValues();
        for (let i = lignes.length - 1; i >= 0; i--) {
          if (typeof lignes[i][0] === "string" && lignes[i][0].toLowerCase().includes("hreflang")) {
            feuilleSuivi.deleteRow(i + 1);
            console.log(`[INFO] Ligne Hreflang supprimée en ligne ${i + 1} de la feuille "Suivi"`);
            break;
          }
        }
      }
      const feuilleHreflang = classeur.getSheetByName("Hreflang");
      if (feuilleHreflang) {
        classeur.deleteSheet(feuilleHreflang);
        console.log(`[INFO] Feuille "Hreflang" supprimée`);
      }
    } catch (e) {
      console.error(`[ERREUR] Multilingue : ${e}`);
    }
  }

  // ----------------------------
  // FILTRES SEO
  if (donnees.seoFiltres === "Non") {
    try {
      const feuilleFiltre = classeur.getSheetByName("Filtre");
      if (feuilleFiltre) {
        feuilleFiltre.deleteRows(8, 11); // Lignes 8 à 18 incluses
        console.log(`[INFO] Lignes 8 à 18 supprimées dans la feuille "Filtre"`);
      }
    } catch (e) {
      console.error(`[ERREUR] SEO Filtres : ${e}`);
    }
  }

  // ----------------------------
  // VIDÉO
  if (donnees.video === "Non") {
    try {
      const feuilleSuivi = classeur.getSheetByName("Suivi");
      if (feuilleSuivi) {
        const lignes = feuilleSuivi.getRange("B1:B").getValues();
        for (let i = lignes.length - 1; i >= 0; i--) {
          if (typeof lignes[i][0] === "string" && lignes[i][0].toLowerCase().includes("vidéo")) {
            feuilleSuivi.deleteRow(i + 1);
            console.log(`[INFO] Ligne Vidéo supprimée en ligne ${i + 1} de la feuille "Suivi"`);
            break;
          }
        }
      }
      const feuilleVideo = classeur.getSheetByName("Vidéo");
      if (feuilleVideo) {
        classeur.deleteSheet(feuilleVideo);
        console.log(`[INFO] Feuille "Vidéo" supprimée`);
      }
    } catch (e) {
      console.error(`[ERREUR] Vidéo : ${e}`);
    }
  }

  // ----------------------------
  // RÉSEAUX SOCIAUX / OpenGraph
  if (donnees.partageSocial === "Non") {
    try {
      const feuilleSuivi = classeur.getSheetByName("Suivi");
      if (feuilleSuivi) {
        const lignes = feuilleSuivi.getRange("B1:B").getValues();
        for (let i = lignes.length - 1; i >= 0; i--) {
          if (typeof lignes[i][0] === "string" && lignes[i][0].toLowerCase().includes("open graph")) {
            feuilleSuivi.deleteRow(i + 1);
            console.log(`[INFO] Ligne Open Graph supprimée en ligne ${i + 1} de la feuille "Suivi"`);
            break;
          }
        }
      }
      const feuilleOG = classeur.getSheetByName("OpenGraph");
      if (feuilleOG) {
        classeur.deleteSheet(feuilleOG);
        console.log(`[INFO] Feuille "OpenGraph" supprimée`);
      }
    } catch (e) {
      console.error(`[ERREUR] Open Graph : ${e}`);
    }
  }
}

function getDonneesPriseDeBrief() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');

  return {
    client: feuille.getRange('B3').getValue(),
    typeProjet: feuille.getRange('B6').getValue(),
    registrar: feuille.getRange('C9').getValue(),
    hebergeur: feuille.getRange('C10').getValue(),
    serveur: feuille.getRange('C11').getValue(),
    cmsActuel: feuille.getRange('C14').getValue(),
    nouveauCms: feuille.getRange('C15').getValue(),
    domaineActuel: feuille.getRange('C19').getValue(),
    sousDomaineActuel: feuille.getRange('C20').getValue(),
    domaineNouveau: feuille.getRange('C22').getValue(),
    sousDomaineNouveau: feuille.getRange('C23').getValue(),
    finUrlActuel: feuille.getRange('C27').getValue(),
    sousRepertoireActuel: feuille.getRange('C28').getValue(),
    finUrlNouveau: feuille.getRange('C30').getValue(),
    sousRepertoireNouveau: feuille.getRange('C31').getValue(),
    urlPreprod: feuille.getRange('C34').getValue(),
    blocagePreprod: feuille.getRange('C35').getValue(),
    doubleSecurite: feuille.getRange('C36').getValue(),
    utilisationCdn: feuille.getRange('C39').getValue(),
    cdn: feuille.getRange('C40').getValue(),
    multilingue: feuille.getRange('C43').getValue(),
    seoFiltres: feuille.getRange('C44').getValue(),
    video: feuille.getRange('C45').getValue(),
    partageSocial: feuille.getRange('C46').getValue(),
    chefSeo: feuille.getRange('C49').getValue(),
    chefWeb: feuille.getRange('C50').getValue(),
    developpeur: feuille.getRange('C51').getValue(),
    adminServeur: feuille.getRange('C52').getValue(),
    graphiste: feuille.getRange('C53').getValue(),
    integrateurMaquette: feuille.getRange('C54').getValue(),
    integrateurContenu: feuille.getRange('C55').getValue(),
    integrateurSeo: feuille.getRange('C56').getValue(),
    gestionDomaine: feuille.getRange('C57').getValue(),
    gestionHebergement: feuille.getRange('C58').getValue(),
    gestionRedirections: feuille.getRange('C59').getValue()
  };
}
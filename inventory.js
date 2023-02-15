/** @OnlyCurrentDoc */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('[Découvrir G Suite]')
  .addItem('Inventorier vos fichiers', 'listeInventaire')    
  .addToUi(); 
}

// Dressez la liste de tous les fichiers et dossiers, et écrire dans la feuille actuelle
function listeInventaire(){
  var ui = SpreadsheetApp.getUi(); // 
  var resultat = ui.prompt(
    'Inventaire à réaliser',
    'Indiquer l\'ID  du dossier:',
    ui.ButtonSet.OK_CANCEL);
  var button = resultat.getSelectedButton();
  var idDuDossier = resultat.getResponseText();
  if (button == ui.Button.OK) {
    obtenirArborescenceDossiers(idDuDossier, true);
  }   
}

// =======================================
// Obtenir l'arborescence des dossiers
// =======================================

function obtenirArborescenceDossiers(idDuDossier, listeTotale) {
  try {
    // Récupérer le dossier par identifiant
    var dossierParent = DriveApp.getFolderById(idDuDossier);
    
    // Initialiser la feuille de calcul
    var data;
    var feuille = SpreadsheetApp.getActiveSheet();
    feuille.clear();
    feuille.appendRow(["Chemin", "Nom", "Type", "Date", "URL", "Dernière mise à jour", "Propriétaire", "Taille"]);
    feuille.getRange(feuille.getLastRow(), 1, 1, feuille.getLastColumn()).setFontWeight("bold");
    
    // Obtenir les fichiers et les dossiers
    obtenirDossiersEnfants(dossierParent.getName(), dossierParent, data, feuille, listeTotale);
    obtenirFichiersRacine(dossierParent.getName(), dossierParent, data, feuille, listeTotale);
    
  } catch (e) {
    Logger.log(e.toString());
  }
};

// Obtenir la liste des fichiers et dossiers et leurs métadonnées en mode récursif
function obtenirDossiersEnfants(nomDossierParent, dossierParent, data, feuille, listeTotale) {
  var dossierEnfants = dossierParent.getFolders();
  
  // Liste des dossiers à l'intérieur du dossier
  while (dossierEnfants.hasNext()) {
    var dossierEnfant = dossierEnfants.next();
    data = [
      nomDossierParent + "/" + dossierEnfant.getName(),
      dossierEnfant.getName(),
      ' ',
      dossierEnfant.getDateCreated(),
      dossierEnfant.getUrl(),
      dossierEnfant.getLastUpdated(),
      dossierEnfant.getOwner().getName(),
      dossierEnfant.getSize()
    ];
    // Ecriture dans la feuille
    feuille.appendRow(data);
    feuille.getRange(feuille.getLastRow(), 1, 1, 1).setFontWeight("bold");
    
    // Liste des fichiers contenus dans le dossier
    var fichiers = dossierEnfant.getFiles();
    while (listeTotale & fichiers.hasNext()) {
      var fichierEnfant = fichiers.next();
      data = [
        "  " + nomDossierParent + "/" + dossierEnfant.getName() + "/" + fichierEnfant.getName(),
        fichierEnfant.getName(),
        fichierEnfant.getMimeType(),    
        fichierEnfant.getDateCreated(),
        fichierEnfant.getUrl(),
        fichierEnfant.getLastUpdated(),
        fichierEnfant.getOwner().getName(),
        fichierEnfant.getSize()
      ];
      // Ecriture dans la feuille
      feuille.appendRow(data);
    }
    
    // Appel récursif du sous-dossier
    obtenirDossiersEnfants(nomDossierParent + "/" + dossierEnfant.getName(), dossierEnfant, data, feuille, listeTotale);  
  }
  
};

// Obtenir la liste des fichiers racine
function obtenirFichiersRacine(nomDossierParent, dossierParent, data, feuille, listeTotale) {
  
  // Liste des fichiers contenus dans le dossier
  var fichiers = dossierParent.getFiles();
  while (listeTotale & fichiers.hasNext()) {
    var fichierEnfant = fichiers.next();
    data = [
      nomDossierParent,
      fichierEnfant.getName(),
      fichierEnfant.getMimeType(),    
      fichierEnfant.getDateCreated(),
      fichierEnfant.getUrl(),
      fichierEnfant.getLastUpdated(),
      fichierEnfant.getOwner().getName(),
      fichierEnfant.getSize()
    ];
    // Ecriture dans la feuille
    feuille.appendRow(data);
  }
  
}

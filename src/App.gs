function onOpen() {
  initialize();
}

function initialize(libraryName = "", localScriptProperties = PropertiesService.getScriptProperties()){ 
  
  let prefix = libraryName;
  if(libraryName){
    prefix = prefix + ".";
  }
  
  // Create menu
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('SUIVI PATIENT')
  .addItem('Envoyer la fiche au médecin traitant',prefix + 'sendMRSPatientSheetByMailToGP')
  .addItem('Créer une fiche vierge',prefix + 'createEmptyMRSPatientSheet')
  .addItem('Archiver la fiche',prefix + 'archivePatientSheet')  
  .addItem('Générer les fiches correspondantes',prefix + 'createMRSPatientSheet')
  .addSeparator()
  .addItem("Créer une tournée", prefix +  "createWardRoundSheet")  
  .addItem("Démarrer la tournée",prefix + "startWardRoundSheet")
  .addItem('Archiver la tournée',prefix + 'archiveWardRoundSheet')
  .addSeparator()
  .addItem('Trier les onglets alph.',prefix + 'sortSheetsByName')
  .addToUi();
  
  refreshPatientSheetNames();
  refreshWardRoundInSystem();
  
}

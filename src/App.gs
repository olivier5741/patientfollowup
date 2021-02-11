function onOpen() {
  initialize();
}

function initialize(libraryName = "", localScriptProperties = PropertiesService.getScriptProperties()){ 
  
  let prefix = libraryName;
  if(libraryName){
    prefix = prefix + ".";
  }
  
  const documentProperties = PropertiesService.getDocumentProperties();
  
  const propTable = [
    ["templateSpreadsheetUrl", "https://docs.google.com/spreadsheets/d/19hR21cKs-ho_L7WvwqE1obdkCgE802athUov3CYe8bc/edit"],
    ["patientSheetTemplateSheetName", "zzz_template_patient_1.5.0"],
    ["wardRoundTemplateSheetName", "zzz_template_tour_1.5.0"]
  ];

  for(const i in propTable){
    const p = documentProperties.getProperty(propTable[i][0]);
    if(!p)
      documentProperties.setProperty(propTable[i][0],propTable[i][1])
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

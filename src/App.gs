// TODO USERS

// * delete rule for TA diast
// * add rule for pouls <50
// * push new version for param column width in ward round
// * group param in patient sheet -> to discuss

// TODO google sheet not performing action when asking rights -> could not find a way to solve this
// TODO improve user experience : defense coding
// TODO add link to mode d emploi in menu
// TODO set developer mode false in every user (otherwise version not working)
// TODO set templates as local cache

// TODO URGENT in wardround, lookup for patient sheet link expects the list to be sorted

// https://leanpub.com/understandinges6/read/
// template literals with `
// let name = "Nicholas", message = `Hello, ${name}.`;

// use const or let
// use regex with "pattern".exec()

// default parameters
// function add(first, second = getValue()) {
//    return first + second;
//}

// function pick(object, ...keys) {

// 2 sec per sheet when generating patients
// 0.5 sec per sheet when transfering parameters
function setProperty()
{
  PropertiesService.getScriptProperties().setProperty("templateSpreadsheetUrl","https://docs.google.com/spreadsheets/d/1jaGvPITO4F-7gfh7x_QkmRI_wyLkahj14UsJkHUNlt0/edit");
 }


function deletePatientSheets(){
    const regExp = new RegExp("^[_z]+.*$")
    const names = patientSheetNames();
    const currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  
    for (let key in names)
    {
      const sheet = currentSpreadSheet.getSheetByName(names[key]);
      currentSpreadSheet.deleteSheet(sheet);
    }
}

function onOpen() {
  initialize();
}

var localScriptProperties;

function initialize(libraryName = "", localScriptProperties = PropertiesService.getScriptProperties()){
  
  if(localScriptProperties.getProperty("patientSheetTemplateSpreadsheetUrl") == null){
    localScriptProperties.setProperty("patientSheetTemplateSpreadsheetUrl", "https://docs.google.com/spreadsheets/d/11IcFYz3_zF-6VdUTA-UZJMmH2Q1-M-LJGDUtz3zoKvI/edit");
  }
  
  if(localScriptProperties.getProperty("templateSpreadsheetUrl") == null){
    localScriptProperties.setProperty("templateSpreadsheetUrl", "https://docs.google.com/spreadsheets/d/11IcFYz3_zF-6VdUTA-UZJMmH2Q1-M-LJGDUtz3zoKvI/edit");
  }
  
  if(localScriptProperties.getProperty("patientSheetTemplateSheetName") == null){
    localScriptProperties.setProperty("patientSheetTemplateSheetName", "zzz_template_patient_1.5.0");
  }
  
  if(localScriptProperties.getProperty("wardRoundTemplateSheetName") == null){
    localScriptProperties.setProperty("wardRoundTemplateSheetName", "zzz_template_tour_1.5.0");
  }
  
  PropertiesService.getDocumentProperties().deleteAllProperties();
  PropertiesService.getDocumentProperties().setProperties(localScriptProperties.getProperties());
   
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

// proxy for sheet buttons

function function1(){
  sendMRSPatientSheetByMailToGP();
}

function function2(){
}

function function3(){
}

function function4(){
}

function function5(){
}

function function6(){
}

function function7(){
}

function function8(){
}

function function9(){
}

function function10(){
}
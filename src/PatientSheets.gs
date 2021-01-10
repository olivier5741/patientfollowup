function patientSheetArchived(){
  refreshPatientSheetNames();
}

function patientSheetCreated(){
  refreshPatientSheetNames();
}

function patientSheetsGenerated(){
  refreshPatientSheetNames();
}

function patientSheetIdAndNameArray() {
  const regExp = new RegExp("^[_z]+.*$")
  return SpreadsheetApp.getActiveSpreadsheet()
  .getSheets()
  .map(s => [s.getSheetId(),s.getName()] )           
  .filter(n => !regExp.exec(n[1]) );
}

// @deprecated
function patientSheetNames() {
  const regExp = new RegExp("^[_z]+.*$")
  return SpreadsheetApp.getActiveSpreadsheet()
           .getSheets()
           .map(s => s.getName())           
           .filter(n => !regExp.exec(n) );
}

function refreshPatientSheetNames(){
  const systemSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("zzz_system");
  const a = patientSheetIdAndNameArray();
  systemSheet.getRange(3,1,1000,2).clearContent();
  
  if(a.length > 0)
    systemSheet.getRange(3,1,a.length,2).setValues(a);
}

function importPatientSheetTemplate(destination){
  
 // TODO change way of coding to call PropertiesService earlier 
 const documentProperties = PropertiesService.getDocumentProperties();
 const source = SpreadsheetApp.openByUrl(documentProperties.getProperty("patientSheetTemplateSpreadsheetUrl"))
 const template = source.getSheetByName(documentProperties.getProperty("patientSheetTemplateSheetName"));
 const sheet = template.copyTo(destination);
 copySheetRangeProtectionWarnings(template,sheet);
 return sheet;
}

function createEmptyMRSPatientSheet(){  
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Cr�er une nouvelle fiche patient', 'Entrer le nom de la fiche', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.CANCEL){
    return;
  }
  
  const sheetName = response.getResponseText();
  const currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();  
  const sheet = importPatientSheetTemplate(currentSpreadSheet);
  sheet.setName(sheetName);
  sheet.activate();
  
  patientSheetCreated();
}

// copySheetRangeProtectionWarnings(template,sheet);
// rename to generate
function createMRSPatientSheet(){ 
  
  const currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = currentSpreadSheet.getActiveSheet();
  const currentRange = currentSheet.getActiveRange().getValues();
  
  const sheetsAmount = currentRange.length;
  
  // Demande de confirmation
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
     "G�n�rer fiche(s) patient(s)",
     `Voulez-vous g�n�rer ${sheetsAmount} fiche(s) patient(s).`,
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.NO) {
    return;
  }
  
  const localTemplate = importPatientSheetTemplate(currentSpreadSheet);
  localTemplate.setName(Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd") + "-temporary-template");
  
  const rangeToUpdate = "B4:B12";
  const templateValues = localTemplate.getRange(rangeToUpdate).getValues();
  
  for (let key in currentRange)
  {
    const row = currentRange[key]; 

    const sheetName = row[0];
    
    const newSheet = currentSpreadSheet.insertSheet(sheetName, {template: localTemplate});    
    copySheetRangeProtectionWarnings(localTemplate,newSheet);
    
    const templateValuesCopy = Array.from(templateValues);
    
    templateValuesCopy[0][0] = row[1];
    templateValuesCopy[1][0] = row[2];
    templateValuesCopy[3][0] = row[3];
    templateValuesCopy[4][0] = row[4];
    templateValuesCopy[5][0] = row[5];
    templateValuesCopy[7][0] = row[6];
    templateValuesCopy[8][0] = row[7];
    
    newSheet.getRange(rangeToUpdate).setValues(templateValuesCopy)
        
    currentSpreadSheet.toast(`La fiche patient ${sheetName} a �t� g�n�r�e`,"Fiche g�n�r�e");
  }
  
  currentSpreadSheet.deleteSheet(localTemplate)
  
  patientSheetsGenerated();
  
}

function sendMRSPatientSheetByMailToGP() {
  
  const currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = currentSpreadSheet.getActiveSheet();
  
  // Obtenir la valeur de la cellule MT dans la feuille patient
  const patientPractitionerListItem = currentSheet.getRange('B9').getValue();
  const patientName = currentSheet.getRange('B4').getValue() + " " + currentSheet.getRange('B5').getValue();
  const nurseName = currentSheet.getRange('B60').getValue();
  
  // Trouvez l'index de la ligne du medecin traitant correspondant dans la feuille coord MG
  const practitionerIndex = currentSpreadSheet.getRangeByName("PractitionerListItems").getValues().findIndex(r => {return r[0] == patientPractitionerListItem});
  
  const practitionerEmail = currentSpreadSheet.getRangeByName("PractitionerEmails").getCell(practitionerIndex + 1,1).getValue();
  const practitionerFullName = currentSpreadSheet.getRangeByName("PractitionerFullNames").getCell(practitionerIndex + 1,1).getValue() 
  
  // Demande de confirmation
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
     "Envoi de la fiche au m�decin traitant",
     `Voulez-vous envoyer la fiche de ${patientName} � l'adresse ${practitionerEmail}.`,
      ui.ButtonSet.YES_NO); // TODO change to YES_CANCEL

  // Process the user's response.
  if (result == ui.Button.NO) {
    return;
  }
  
  // Envoie de l'email
  const message = 
`Docteur ${practitionerFullName},
  
Vous trouverez en attach� le compte rendu journalier de votre patient ${patientName} h�berg� ds notre MR(S).
Nous vous invitons � en prendre connaissance et nous transmettre vos remarques �ventuelles.

Pour la MRS,
${nurseName}`;
  
  const subject = `Fiche patient covid en MRS : ${patientName}`;

  const pdf = convertSheetToPdf(currentSpreadSheet,currentSheet,"Fiche MRS de " + patientName);
  MailApp.sendEmail(practitionerEmail, subject, message, {attachments:pdf});
  
  currentSpreadSheet.toast('Mail envoy�', 'Mail envoy�');
}

function archivePatientSheet(){
  const currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = currentSpreadSheet.getActiveSheet();
  const sheetName = sheet.getName();
  
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
     "Archiver la fiche patient",
     "Voulez-vous archiver la fiche patient : " + sheetName,
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  if (result == ui.Button.CANCEL) {
    return;
  }
  
  const folder = DriveApp.getFileById(currentSpreadSheet.getId()).getParents().next();
  const archiveFolder = folder.getFoldersByName("Archives").next().getFoldersByName("Patients").next();
  
  const pdf = convertSheetToPdf(currentSpreadSheet,sheet,Utilities.formatDate(new Date(), "GMT+1", "yyyy_MM") + "_" + sheetName);
  archiveFolder.createFile(pdf);
  
  currentSpreadSheet.deleteSheet(sheet);
  
  patientSheetArchived();
}

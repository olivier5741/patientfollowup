// TODO could also finish and protect the whole sheet, paraph and check if everything is done
// TODO make sure sheet is a wardround before starting it

function wardRoundCreated(){
  refreshWardRoundInSystem();
}

function wardRoundSheetArchived(){
  refreshWardRoundInSystem();
}

function wardRoundSheetIdAndNameArray() {
  const regExp = new RegExp("^_tourn�e.*$")
  return SpreadsheetApp.getActiveSpreadsheet()
  .getSheets()
  .map(s => [s.getSheetId(),s.getName()] )           
  .filter(n => regExp.exec(n[1]) );
}

function refreshWardRoundInSystem(){
  const systemSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("zzz_system");
  const a = wardRoundSheetIdAndNameArray();
  systemSheet.getRange(3,5,1000,2).clearContent();
  
  if(a.length > 0)
    systemSheet.getRange(3,5,a.length,2).setValues(a);
}

function createWardRoundSheet(){  
  
  const ui = SpreadsheetApp.getUi();
  
  const date = new Date();
  const sheetNameSuggestion = "_tourn�e " + Utilities.formatDate(date, "GMT+1", "yyyy-MM-dd") 
                   + " ou " + "_tourn�e " +  Utilities.formatDate(date, "GMT+1", "yyyy-MM-dd HH") + "h";
  
  let wardRoundNameValid = false;
  let sheetName;
  
  while(wardRoundNameValid == false){
    
    const response = ui.prompt(`Entrer le nom de la tourn�e (suggestion: ${sheetNameSuggestion})`, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.CANCEL){
      return;
    }
    const regExp = new RegExp("^_tourn�e.*$")
    
    if(regExp.exec(response.getResponseText()) == null){
      ui.alert("Le nom doit commencer par _tourn�e", ui.ButtonSet.OK);      
    }else{
      wardRoundNameValid = true;
      sheetName = response.getResponseText();
    }
    
  }
  
  const currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();  
  const sheet = importWardRoundSheetTemplate(currentSpreadSheet);
  sheet.setName(sheetName);  
  sheet.getRange("B1").setValue(Utilities.formatDate(date, "GMT+1", "dd/MM/yyyy HH:00"));
  
  const patients = currentSpreadSheet.getRangeByName("SystemPatientSheets").getValues().filter(r => r[2]).map(r => [r[1]]);
  
  const nbRowsToKeep = patients.length + 4 + 10;
  sheet.deleteRows(nbRowsToKeep+1,sheet.getLastRow()-nbRowsToKeep);
  
  const nbColumnsToKeep = 4;
  sheet.deleteColumns(nbColumnsToKeep+1, sheet.getLastColumn()-nbColumnsToKeep);
  
  if(patients.length > 0)
    sheet.getRange("A5:A"+(patients.length + 4)).setValues(patients);
  
  sheet.activate(); 
  currentSpreadSheet.moveActiveSheet(2)
  
  wardRoundCreated();
}

function importWardRoundSheetTemplate(destination){
  
 // TODO change way of coding to call PropertiesService earlier 
 const documentProperties = PropertiesService.getDocumentProperties();
 const source = SpreadsheetApp.openByUrl(documentProperties.getProperty("templateSpreadsheetUrl"))
 const template = source.getSheetByName(documentProperties.getProperty("wardRoundTemplateSheetName"));
 const sheet = template.copyTo(destination);
 copySheetRangeProtectionWarnings(template,sheet);
 return sheet;
}

// set status started to prevent creating multi columns
function startWardRoundSheet(){
  const currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();  
  const sheet = currentSpreadSheet.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  const sheetName = sheet.getName();
  const regExp = new RegExp("^_tourn�e.*$")
  
  if(regExp.exec(sheetName) == null){
    ui.alert("Le fichier n'est pas une tourn�e. Le nom doit commencer par _tourn�e", ui.ButtonSet.OK);
    return;
  }
  
  const statusRange = sheet.getRange("B3");
  
  if(statusRange.getValue() == "D�marr�e"){
    const result = ui.alert("La tourn�e a d�j� �t� d�marr�e", ui.ButtonSet.OK);
    return;
  }
  
  statusRange.setValue("D�marr�e")
  
  const template = importWardRoundSheetTemplate(currentSpreadSheet);
  
  // E to AD
  const notation = "E1:AO"+sheet.getLastRow();
  sheet.insertColumnsAfter(4, 34)

  const templateRange = template.getRange(notation);
  const sheetRange = sheet.getRange(notation);
  
  templateRange.copyTo(sheetRange);
  templateRange.copyTo(sheetRange, SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS,false);
  
  currentSpreadSheet.deleteSheet(template);
  
  protectFormulaRangeWithWarning(sheet.getRange("C5:C"));
  protectFormulaRangeWithWarning(sheet.getRange("E5:AO"));
  
  const patientNames = sheet.getRange("A5:A").getValues().map(r => r[0]).filter(n => n != "");
  const date = sheet.getRange("B1").getValue();
  
  for (let key in patientNames)
  {
    const patientSheet = currentSpreadSheet.getSheetByName(patientNames[key]);
    
    const dates = patientSheet.getRange('B19:19').getValues()[0];
    const column = dates.findIndex(e => e == "");
    
    patientSheet.getRange(19, column + 2).setValue(date);
  }
}

// reutiliser du code avec patient
function archiveWardRoundSheet(){
  const currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = currentSpreadSheet.getActiveSheet();
  const sheetName = sheet.getName();
  
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
     "Archiver la tourn�e",
     "Voulez-vous archiver la tourn�e : " + sheetName,
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  if (result == ui.Button.CANCEL) {
    return;
  }
  
  const folder = DriveApp.getFileById(currentSpreadSheet.getId()).getParents().next();
  const archiveFolder = folder.getFoldersByName("Archives").next().getFoldersByName("Tourn�es").next();
  
  const pdf = convertSheetToPdf(currentSpreadSheet,sheet,Utilities.formatDate(new Date(), "GMT+1", "yyyy_MM") + sheetName,"H");
  archiveFolder.createFile(pdf);
  
  currentSpreadSheet.deleteSheet(sheet);
  
  wardRoundSheetArchived();
}

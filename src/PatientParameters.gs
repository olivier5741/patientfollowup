// @deprecated
function insertTodaysPatientParameters(){
  const currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = currentSpreadSheet.getActiveSheet();
  
  // TODO warning if no row selected in this sheet
  
  const row = parametersSheet.getActiveRange().getRow();
  const sheetNames = patientSheetNames();
  const date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
  
  const range = parametersSheet.getRange(row, 1, sheetNames.length, 2);
  range.setValues(sheetNames.map( n => [date,n] ));  
}

// @deprecated
function transferDailyPatientParametersToPatientSheet(){
  const currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = currentSpreadSheet.getActiveSheet();
  const currentRange = currentSheet.getActiveRange();
  const currentValues = currentRange.getValues();
  
  for (let key in currentValues)
  {
    const row = currentValues[key]; 
    const sheetName = row[1];
    const sheet = currentSpreadSheet.getSheetByName(sheetName);
    
    const dates = sheet.getRange('B19:S19').getValues()[0];
    const column = dates.findIndex(e => e == "");
    const columnLetter = columnToLetter(column + 2);
    
    const a1 = new Array(1);
    a1[0] = row.slice(4, 15);
    
    const a2 = new Array(1);
    a2[0] = row.slice(16, 23);
    
    const a3 = new Array(1);
    a3[0] = row.slice(23,28);
    
    sheet.getRange(columnLetter + '19').setValue(row[0]);
    sheet.getRange(columnLetter + '20:' + columnLetter + '30').setValues(transpose(a1));
    sheet.getRange(columnLetter + '32:' + columnLetter + '38').setValues(transpose(a2));
    sheet.getRange(columnLetter + '54:' + columnLetter + '58').setValues(transpose(a3));
    sheet.getRange(columnLetter + '59').setValue(row[29]);
    
    if(row[28] != "")
      sheet.getRange("D1").setValue(row[28]);
  }
  
  const formulas = currentRange.getFormulas();
  currentRange.clearContent();
  currentRange.setFormulas(formulas);
  
}


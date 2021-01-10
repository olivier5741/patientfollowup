// initialize library on opening the spreadsheet
function onOpen(){
  SuiviPatientCovidEnMRS.initialize("SuiviPatientCovidEnMRS",PropertiesService.getScriptProperties());
}

// legacy
function sendMRSPatientSheetByMailToGP(){
  SuiviPatientCovidEnMRS.function1();
}

// proxy for functions used by buttons or potential buttons
function function1(){ SuiviPatientCovidEnMRS.function1(); }
function function2(){ SuiviPatientCovidEnMRS.function2(); }
function function3(){ SuiviPatientCovidEnMRS.function3(); }
function function4(){ SuiviPatientCovidEnMRS.function4(); }
function function5(){ SuiviPatientCovidEnMRS.function5(); }
function function6(){ SuiviPatientCovidEnMRS.function6(); }
function function7(){ SuiviPatientCovidEnMRS.function7(); }
function function8(){ SuiviPatientCovidEnMRS.function8(); }
function function9(){ SuiviPatientCovidEnMRS.function9(); }
function function10(){ SuiviPatientCovidEnMRS.function10(); }

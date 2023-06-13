function MarkAsPaid() {
  var datasource = SpreadsheetApp.getActiveSheet()
  var datamodified = datasource.getDataRange().getValues()

  Logger.log(datamodified[0][3]);
  
 
}

function MarkAsPaid2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dSample = ss.getSheetByName('dataSample');
  var dForm = ss.getSheetByName('dataForm');
  var dExtract = ss.getSheetByName('dataExtract');
  var dCommander = ss.getSheetByName('dataCommander');
  var dC = dCommander.getLastRow()+1;
  var dcCommand = dForm.getRange(2,7).getValue();
  var dcColDate = dForm.getRange(2,8).getValue();
  var dcDate = dForm.getRange(2,9).getValue();
  var dcColRef = dForm.getRange(2,10).getValue();
  var dcRef = dForm.getRange(2,11).getValue();

  Logger.log(dExtract.getDataRange().getValues().length);
  
  for (i = 0; i <= (dExtract.getDataRange().getValues().length - 1); i++){
//  Logger.log(dExtract.getDataRange().getValues()[i][0]);
 // Logger.log(dExtract.getDataRange().getValues()[i][1]);
 // Logger.log(dExtract.getDataRange().getValues()[i][2]);
 // Logger.log(dExtract.getDataRange().getValues()[i][3]);  
  }

  for (i = 0; i <= (dExtract.getDataRange().getValues().length - 1); i++){
//  Logger.log(dExtract.getDataRange().getValues()[i][0]);
 // Logger.log(dExtract.getDataRange().getValues()[i][1]);
 // Logger.log(dExtract.getDataRange().getValues()[i][2]);
 // Logger.log(dExtract.getDataRange().getValues()[i][3]);  
//  dCommander.getRange(dC+i,1).setValue(dExtract.getDataRange().getValues()[i][0]);
// dCommander.getRange(dC+i,2).setValue(dExtract.getDataRange().getValues()[i][1]);
 dCommander.getRange(dC+i,1).setValue(dcCommand);
 dCommander.getRange(dC+i,2).setValue(dExtract.getDataRange().getValues()[i][2]);
 dCommander.getRange(dC+i,3).setValue(dExtract.getDataRange().getValues()[i][3]);
   dCommander.getRange(dC+i,4).setValue(dcColDate);
   dCommander.getRange(dC+i,5).setValue(dcDate);
  dCommander.getRange(dC+i,6).setValue(dcColRef);
  dCommander.getRange(dC+i,7).setValue(dcRef);
}
}

function RunDataCommander(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var dCommander = ss.getSheetByName('dataCommander');




Logger.log(dCommander.getDataRange().getValues().length);

  for (i = 1; i <= (dCommander.getDataRange().getValues().length - 1); i++){
 // Logger.log(dCommander.getDataRange().getValues()[i][0]); // Command
 // Logger.log(dCommander.getDataRange().getValues()[i][1]); // Sheetnme
 // Logger.log(dCommander.getDataRange().getValues()[i][2]); // Row#
 // Logger.log(dCommander.getDataRange().getValues()[i][3]); // Col#
 //  Logger.log(dCommander.getDataRange().getValues()[i][4]); // Date
 // Logger.log(dCommander.getDataRange().getValues()[i][5]); // Col#
 // Logger.log(dCommander.getDataRange().getValues()[i][6]);  // Ref. No.

ss.getSheetByName(dCommander.getDataRange().getValues()[i][1]).getRange(dCommander.getDataRange().getValues()[i][2],dCommander.getDataRange().getValues()[i][3]).setValue(dCommander.getDataRange().getValues()[i][4]);

ss.getSheetByName(dCommander.getDataRange().getValues()[i][1]).getRange(dCommander.getDataRange().getValues()[i][2],dCommander.getDataRange().getValues()[i][5]).setValue(dCommander.getDataRange().getValues()[i][6]);

  }

}

function Jexperiment(){
   var res = UrlFetchApp.fetch("https://omsysapi.omaserver.com/index.php/survey/content/1?id=107574&token=839C80BE-9862-4475-AC29-3F98E3DA36B5&siteid=221975");
   var content = res.getContentText();

   //var json = JASON.parse(content);
   //var base = json["rates"]["EUR"];
   Logger.log(content);
}

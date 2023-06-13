function GCRefNumFBName() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var targetSheet = ss.getSheetByName('RefNoFBName');
var targetSheet2 = ss.getSheetByName('BankFBName');
var sourceSheet = ss.getSheetByName('VerifyReferenceNo');
var RefNum = sourceSheet.getRange('D1').getValue();
var GCashNum = sourceSheet.getRange('D2').getValue();
var FBName = sourceSheet.getRange('D3').getValue();
var Bank = sourceSheet.getRange('F3').getValue();
var LastRow = targetSheet.getLastRow()+1;
var LastRow2 = targetSheet2.getLastRow()+1;
if (RefNum == "") return;
targetSheet.getRange(LastRow,1).setValue(new Date());
targetSheet.getRange(LastRow,2).setValue(RefNum);
targetSheet.getRange(LastRow,3).setValue(GCashNum);
targetSheet.getRange(LastRow,4).setValue(FBName);
if (Bank == "") return;
targetSheet2.getRange(LastRow2,1).setValue(new Date());
targetSheet2.getRange(LastRow2,2).setValue(Bank);
targetSheet2.getRange(LastRow2,3).setValue(FBName);
}

function onEdit(e) {
ConvertInvoice2RefNum(e);
EntryRefNum(e);
}

function ConvertInvoice2RefNum(e){
  const RangeModified = e.range
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('VerifyReferenceNo');
  var RefNum = sourceSheet.getRange('F1').getValue();
    if (RangeModified.getSheet().getSheetName() !== "VerifyReferenceNo") return
  if (RangeModified.getColumn() !== 6) return 
  if (RangeModified.getRow() !== 8 ) return
  RangeModified.offset(0,-3).setValue(RefNum)
}

function DefaultTransaction(e){
  const RangeModified = e.range
  
  console.log("A1 Notation", RangeModified.getA1Notation())
  console.log("sheetname", RangeModified.getSheet().getSheetName())
  console.log("Column Number", RangeModified.getColumn())
  console.log("Row Number", RangeModified.getRow())

  if (RangeModified.getSheet().getSheetName() !== "CASH-IN_Inventory-OUT") return
  if (RangeModified.getColumn() !== 3) return 
  if (RangeModified.getRow() < 2 ) return

  RangeModified.offset(0,8).setValue("pending")
  //RangeModified.offset(0,-1).setValue(new Time().toLocaleTimeString())
}

function AutofillQuantity(e){
  const RangeModified = e.range
  
  console.log("A1 Notation", RangeModified.getA1Notation())
  console.log("sheetname", RangeModified.getSheet().getSheetName())
  console.log("Column Number", RangeModified.getColumn())
  console.log("Row Number", RangeModified.getRow())

  if (RangeModified.getSheet().getSheetName() !== "CASH-IN_Inventory-OUT") return
  if (RangeModified.getColumn() !== 5) return 
  if (RangeModified.getRow() < 2 ) return
  var aQty = RangeModified.offset(0,23).getValue()

  RangeModified.offset(0,1).setValue(aQty)
  
}

function AutofillTotalAmt(e){
  const RangeModified = e.range
  
  console.log("A1 Notation", RangeModified.getA1Notation())
  console.log("sheetname", RangeModified.getSheet().getSheetName())
  console.log("Column Number", RangeModified.getColumn())
  console.log("Row Number", RangeModified.getRow())

  if (RangeModified.getSheet().getSheetName() !== "CASH-IN_Inventory-OUT") return
  if (RangeModified.getColumn() !== 6) return 
  if (RangeModified.getRow() < 2 ) return
  var aTotalAmt = RangeModified.offset(0,24).getValue()

  RangeModified.offset(0,3).setValue(aTotalAmt)
  
}

function LogForCashINInventoryOUT(e){
  const RangeModified = e.range
  const RangeData = e.value
  var vv = SpreadsheetApp.getActiveSheet().getActiveCell().getValue()
  console.log("A1 Notation", RangeModified.getA1Notation())
  console.log("sheetname", RangeModified.getSheet().getSheetName())
  console.log("Column Number", RangeModified.getColumn())
  console.log("Row Number", RangeModified.getRow())
  console.log("Data", vv)
  console.log("Data", RangeData)


  if (RangeModified.getSheet().getSheetName() !== "CASH-IN_Inventory-OUT") return
  if (RangeModified.getColumn() !== 3) return 
  if (RangeModified.getRow() < 2 ) return

  const TargetSheet = e.source.getSheetByName("Logs")

  TargetSheet.appendRow([new Date(),RangeModified.getSheet().getSheetName(),RangeModified.getColumn(),RangeModified.getRow(),RangeModified.getA1Notation(),vv,RangeData])
  
}

function UpdateInvestments(){
var ss = SpreadsheetApp.getSheets().getSheetByName("Logs");
ss.appendRow(["a man", "a plan", "panama"]);
  
}

function hide1column() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D:D').activate();
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('E:E').activate();
};

function UnhideALLColumns() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A:AM').activate();
  spreadsheet.getActiveSheet().showColumns(0, 39);
};

function AddTimeStamp(e){
  //const RangeModified = e.range
  var RangeModified = e.range;
  
  console.log("A1 Notation", RangeModified.getA1Notation())
  console.log("sheetname", RangeModified.getSheet().getSheetName())
  console.log("Column Number", RangeModified.getColumn())
  console.log("Row Number", RangeModified.getRow())

  if (!RangeModified.offset(0,-2).isBlank()) return
  // if (RangeModified.getSheet().getSheetName() !== "posting") return
  if (RangeModified.getColumn() !== 3) return 
  if (RangeModified.getRow() < 3 ) return

 // RangeModified.offset(0,-3).setValue(new Date())
 // var a_Date = RangeModified.offset(0,29).getValue()
 // var a_Time = RangeModified.offset(0,30).getValue()
 // RangeModified.offset(0,-2).setValue(a_Date)
 // RangeModified.offset(0,-1).setValue(a_Time)
  
  if (RangeModified.getSheet().getSheetName() == "Bind.Fb.GcRefNum" || 
      RangeModified.getSheet().getSheetName() == "liveselling"

     ) 
  {
      RangeModified.offset(0,-2).setValue(new Date())
  }
}

function EntryRefNum(e)
{
  const range = e.range;
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('VerifyReferenceNo').getRange('D4').getValue();
  //var ws = ss.getSheetbyName('VerifyReferenceNo');
  // var FinalName = ss.getaActiveSheet.getRange('D5').getValue();
  
  if((range.getSheet().getSheetName() == 'VerifyReferenceNo') && (range.getColumn() == 3) && (range.getRow() == 8)){
 //    ss.getaActiveSheet.getRange('C6').setValue(ss)
 console.log("you came here...")
 SpreadsheetApp.getActiveSpreadsheet().getSheetByName('VerifyReferenceNo').getRange('C6').setValue(ss);
  }
 // console.log('Final Name is ',ss.getaActiveSheet.getRange('D4').getValue());
 console.log("Value of ss: ",ss)
 console.log("Sheet Name: ", range.getSheet().getSheetName())
 console.log("Column No.: ",range.getColumn())
 console.log("Row No. : ", range.getRow())
}










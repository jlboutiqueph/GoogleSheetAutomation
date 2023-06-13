function onOpen(){
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('JL Menu')
    .addItem('Update Column Format','FormatYtoText')
    .addItem('Go to Last Row','GotoLastRow')
    .addItem('Insert Row','Insert1Row')
    .addItem('Insert 5 Rows','Insert5Rows')
    .addItem('Insert 10 Rows','Insert10Rows')
    .addItem('Insert 15 Rows','Insert15Rows')
    .addItem('Insert 20 Rows','Insert20Rows')                
    .addToUi();
}

function onEdit(e) {
AddTimeStamp(e)

}

function sheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

function myFunction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const ws = ss.getSheetByName('plan2023')

  const colData = ws.getRange("J:J").getValues()

  for(let i = colData.length-1;i>=0;i--){
    if(colData[i][0]===true){
      ws.deleteRow(i+1)
    }
  }
}

function mySetSum() {
  let Shit = SpreadsheetApp.getActiveSheet();
  let totalCell = Shit.getRange('AD2');
  totalCell.setValue('hello');

  
  let sumCell = Shit.getRange('AC2');
  sumCell.setFormula("='MONEY Formula'!C2");

}

function mySetSum2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s1 = ss.getSheetByName('MONEY Formula');
  const s2 = ss.getSheetByName('hol');
  const furmila = s1.getRange('C2').getValue();
  
  let Shit = SpreadsheetApp.getActiveSheet();
  let totalCell = Shit.getRange('AD2');
  totalCell.setValue('hello');

  
  let sumCell = Shit.getRange('AC2');
  sumCell.setFormula(furmila);

}

function AddTimeStamp(e){
  //const RangeModified = e.range
  var RangeModified = e.range;
  
  console.log("A1 Notation", RangeModified.getA1Notation())
  console.log("sheetname", RangeModified.getSheet().getSheetName())
  console.log("Column Number", RangeModified.getColumn())
  console.log("Row Number", RangeModified.getRow())

  if (!RangeModified.offset(0,-4).isBlank()) return
  if (RangeModified.getColumn() !== 5) return 
  if (RangeModified.getRow() < 3 ) return
  if (RangeModified.getSheet().getSheetName() == "walkin/pickup" || 
      RangeModified.getSheet().getSheetName() == "liveselling" || 
      RangeModified.getSheet().getSheetName() == "posting" ||
      RangeModified.getSheet().getSheetName() == "loanitem/loanload" ||
      RangeModified.getSheet().getSheetName() == "scrap/cancel/bogus" ||
      RangeModified.getSheet().getSheetName() == "delivery/pickup/meetup/dropoff" ||
      RangeModified.getSheet().getSheetName() == "inventory" ||
      RangeModified.getSheet().getSheetName() == "jlb-expenses" ||
      RangeModified.getSheet().getSheetName() == "hol" ||
      RangeModified.getSheet().getSheetName() == "bills" ||
      RangeModified.getSheet().getSheetName() == "fundtransfer" ||
      RangeModified.getSheet().getSheetName() == "stock-order" ||
      RangeModified.getSheet().getSheetName() == "add-items" ||
      RangeModified.getSheet().getSheetName() == "DEBIT" ||
      RangeModified.getSheet().getSheetName() == "CREDIT" ||
      RangeModified.getSheet().getSheetName() == "shopee"
     ) 
  {
      RangeModified.offset(0,-4).setValue(new Date())
  }
}

function ifGIsBlankThenMakeItZero()
{
  var ssA = SpreadsheetApp.getActive();//changed from openById() for my convenience
  var ss = ssA.getActiveSheet();//change from getSheetByName() for my convenience
  var lastRow = ss.getLastRow();
  var range = ss.getRange(2,7,lastRow,1);//row 2 column 7 (G) lastRow 1 column 
  var data = range.getValues();//Gets all data
  for(var i=0;i<data.length;i++)//this runs over entire selected range 
  {  
    if(!data[i][0])//If true then it's blank
    {
      data[i][0]=0;//notice this is data[i][0] because there is only one column in the range.
    }
  }
  range.setValues(data);//Sets all data.  
}


function FormatYtoText() {
  var spreadsheet = SpreadsheetApp.getActive();
  //macro to format column Y to text
  spreadsheet.getRange('Y:Y').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('@');
  //macro to go to last row
  spreadsheet.getRange('A5').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
};

function GotoLastRow() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A5').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
};

function Insert1Row() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveSheet().getMaxRows(), 1);
};

function Insert5Rows() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveSheet().getMaxRows(), 5);
};

function Insert10Rows() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveSheet().getMaxRows(), 10);
};

function Insert15Rows() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveSheet().getMaxRows(), 15);
};

function Insert20Rows() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveSheet().getMaxRows(), 20);
};



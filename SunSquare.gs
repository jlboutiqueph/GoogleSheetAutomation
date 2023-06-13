function onOpen(){
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Summer')
    .addItem('Calculate Sum','mySetSum')
    .addItem('Calculate Toinks','mySetSum2')
    .addToUi();
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

function onEdit(e) {
AddTimeStampSSBC(e)

}

function AddTimeStampSSBC(e){
  var RangeModified = e.range;
  if (!RangeModified.offset(0,-4).isBlank()) return
  if (RangeModified.getColumn() !== 5) return 
  if (RangeModified.getRow() < 3 ) return
  if (RangeModified.getSheet().getSheetName() == "walkin/pickup:ssbc" || 
      RangeModified.getSheet().getSheetName() == "liveselling:ssbc" || 
      RangeModified.getSheet().getSheetName() == "posting:ssbc" ||
      RangeModified.getSheet().getSheetName() == "scrap/cancel/bogus:ssbc" ||
      RangeModified.getSheet().getSheetName() == "loanitem/loanload:ssbc" ||
      RangeModified.getSheet().getSheetName() == "delivery/meetup/dropoff:ssbc" ||
      RangeModified.getSheet().getSheetName() == "jlb-expenses:ssbc"
     ) 
  {
      RangeModified.offset(0,-4).setValue(new Date())
  }
}






function doGet(request) {
  return HtmlService.createTemplateFromFile('Page')
      .evaluate();
}
var s = SpreadsheetApp.openById('1-4BTmH57oCNUtuS-g-NHXhfjY2Ag7eUMoE7Qq_QT6DI')
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}
function find_row_to_move(){
  var number_sheet = s.getNumSheets()
  for(var i = 0; i < number_sheet; i++){
    if(s.getSheets()[i]=='Sheet2'){
      continue
    }
    s.getSheets()[i].getDataRange().setValues([[new Date]])
    console.log(s.getSheets()[i].getDataRange().getValues())
  }
}
function move(){
  var sheet1 = s.getSheetByName('Sheet1')
  console.log(sheet1[0])
  var sheet2 = s.getSheetByName('Sheet2')
  sheet2.getRange(`A${sheet2.getLastRow()+1}:E${sheet2.getLastRow()+1}`).setValues([sheet1.getRange('A1:E1').getValues()[0]])
  sheet1.getRange('A1:E1').clear()
}


function doGet(request) {
  return HtmlService.createTemplateFromFile('Page')
    .evaluate();
}
var s = SpreadsheetApp.openById('1-4BTmH57oCNUtuS-g-NHXhfjY2Ag7eUMoE7Qq_QT6DI')
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}
function find_row_to_move() {
  var number_sheet = s.getNumSheets()
  for (var i = 0; i < number_sheet; i++) {
    if (s.getSheets()[i] == 'Sheet2') {
      continue
    }
    s.getSheets()[i].getDataRange().setValues([[new Date]])
    console.log(s.getSheets()[i].getDataRange().getValues())
  }
}
function move() {
  var sheet1 = s.getSheetByName('Sheet1')
  console.log(sheet1[0])
  var sheet2 = s.getSheetByName('Sheet2')
  sheet2.getRange(`A${sheet2.getLastRow() + 1}:E${sheet2.getLastRow() + 1}`).setValues([sheet1.getRange('A1:E1').getValues()[0]])
  sheet1.getRange('A1:E1').clear()
}

function lay_cot_du_lieu() {
  var url_sheet = 'https://docs.google.com/spreadsheets/d/1UQQwb0TYRPgtiUIm1tuqyT-Oosnwfj7sOXtxkgvoWKw/edit#gid=480466906'
  var sheet = SpreadsheetApp.openByUrl(url_sheet).getDataRange().getValues()
  console.log(sheet)
}

function test_push_arr(){
  var arr = [[]]
  arr[0].push(12312,324234)
  console.log(arr)
}

function test_setvalues_nhieu_dong(){
  var url_sheet = 'https://docs.google.com/spreadsheets/d/1ywqGvEUf5UFRHHfUq0-N3auYSAAOIw8JD4FkttSmIuc/edit#gid=1815006366'
  var s = SpreadsheetApp.openByUrl(url_sheet).getSheetByName('Log')
  var arr = [
    [ '01bef834',
    new Date(),
    'ADD',
    '7,a60ab7a9,333333,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,',
    'don_hang',
    'Nghiant@hungdunghd.com.vn' ],
  [ '7b866a28',
    new Date(),
    'DELETE',
    '2,88888867,rgdfg,dfgdfg,,,dfg,fdg,gdf,gdfg,,,,,,,,,,,,,,,,,,,,,,,,,,',
    'don_hang',
    'Nghiant@hungdunghd.com.vn' ] ]
  s.getRange(`A${s.getLastRow()+1}:F${s.getLastRow()+arr.length}`).setValues(arr)
}


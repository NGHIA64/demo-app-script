function doGet(request) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate();
}
var s = SpreadsheetApp.openById('1-4BTmH57oCNUtuS-g-NHXhfjY2Ag7eUMoE7Qq_QT6DI')
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}
var thongTin = {
  name : [
    'Sheet1', 'Sheet2', 'Sheet3'
  ],
  url : {
    sheet : {
      test : 'https://docs.google.com/spreadsheets/d/1-4BTmH57oCNUtuS-g-NHXhfjY2Ag7eUMoE7Qq_QT6DI/edit#gid=0'
    }
  }
}
function testSetFormula(){
 var s = SpreadsheetApp.openByUrl(thongTin.url.sheet.test)
 var sheet = s.getSheetByName(thongTin.name[0])
 var formulas = [
["=SUM(B2:B4)", "=SUM(C2:C4)", "=SUM(D2:D4)"],
["=AVERAGE(B2:B4)", "=AVERAGE(C2:C4)", "=AVERAGE(D2:D4)"]
];

var cell = sheet.getRange("B5:D6");
cell.setFormulas(formulas);
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
  console.log(sheet[0])
}

function test_push_arr(){
  var arr = []
  arr.push([2345])
  arr.push([234])
  console.log(arr)
  var array1 = ["Vijendra", "Singh"];
var array2 = ["Singh", "Shakyatieyyyyy1"];
var array3 = ["Singh", "Shakya"];
array1 = array1.concat(array2);
array1 = array1.concat(array3);
console.log(array1);
}

function test_setvalues_nhieu_dong(){
  var url_sheet = 'https://docs.google.com/spreadsheets/d/1ywqGvEUf5UFRHHfUq0-N3auYSAAOIw8JD4FkttSmIuc/edit#gid=1815006366'
  var s = SpreadsheetApp.openByUrl(url_sheet).getSheetByName('Log')
  var arr = [
    [ '01bef834',
    new Date(),
    'ADD',
    '7,a60ab7a9,333553313,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,',
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
function test(){
  
}
function layDuLieuNhap(url, name_sheet_data, name_sheet_import) {
  // var url = 'https://docs.google.com/spreadsheets/d/1FQac1blPRir4AEPgqS9KW_OrZs6f6OIbAxcvx5J62EU/edit#gid=0'
  // var name_sheet_import = 'Nhập Giá'
  // var name_sheet_data = 'giaSP'
  var s = SpreadsheetApp.openByUrl(url)
  var sheet_import = s.getSheetByName(name_sheet_import)
  var sheet_data = s.getSheetByName(name_sheet_data)
  console.log(sheet_import.getRange(2, 1, sheet_import.getLastRow(), sheet_import.getLastColumn()).getValues())
  sheet_data.getRange(sheet_data.getLastRow() + 1, 1, sheet_import.getLastRow(), sheet_data.getLastColumn()).setValues(sheet_import.getRange(2, 1, sheet_import.getLastRow(), sheet_import.getLastColumn()).getValues())
}

function testFormula(){
  var thongTin = {
    url : {
      sheet : {
        khachHang: 'htpps://google.com',
      }
    },
    control : {
      sheetData: s.getSheetByName('khachHang'),
      sheetImport: s.getSheetByName('Nhập khách hàng')
    }
  }
  var s = SpreadsheetApp.openByUrl(thongTin.url.sheet.khachHang)
  
}

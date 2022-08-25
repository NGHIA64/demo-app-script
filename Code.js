function doGet(request) {
  return HtmlService.createTemplateFromFile('index')
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
    '7,a60ab7a9,333553313,3,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,',
    'don_hang',
    'Nghiant@hungdunghd.com.vn' ],
  [ '7b866a28',
    new Date(),
    'DELETE',
    '2,88888867,rgdfg,dfgdfg,,,dfeg,fdg,gdf,gdfg,,,,,,,,,,,,,,,,,,,,,,,,,,',
    'don_hang',
    'Nghiant@hungdunghd.com.vn' ] ]
  s.getRange(`A${s.getLastRow()+1}:F${s.getLastRow()+arr.length}`).setValues(arr)
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


IF(ISBLANK([chon_khuyen_mai]),Min(select(CTKM[so_luong_khach_can_lay],AND([id_sp]=[_THISROW].[id_sp],[ngay_bat_dau]<=date([_THISROW].[date]),
[ngay_ket_thuc]>=date([_THISROW].[date])))),
[chon_khuyen_mai])

IF(and([so_luong]>=lookup(MaxROW("Bảng giá","_RowNumber",and([IDsp]=[_THISROW].[IDsp],[Ngày]>=[Ngày bắt đầu])),
"Bảng giá","IDgia","Số lượng áp dụng giá ngoại lệ"),ISNOTBLANK(lookup(MaxROW("Bảng giá","_RowNumber",and([IDsp]=[_THISROW].[IDsp],[Ngày]>=[Ngày bắt đầu])),
"Bảng giá","IDgia","Số lượng áp dụng giá ngoại lệ"))),lookup(MaxROW("Bảng giá","_RowNumber",and([IDsp]=[_THISROW].[IDsp],[Ngày]>=[Ngày bắt đầu])),
"Bảng giá","IDgia","Giá ngoại lệ"),
if([Chọn kiểu giá]="Giá trần",lookup(MaxROW("Bảng giá","_RowNumber",and([IDsp]=[_THISROW].[IDsp],[Ngày]>=[Ngày bắt đầu])),
"Bảng giá","IDgia","Giá trần"),IF(or([Phân loại]="Đề xuất khuyến mại",[Nguồn hàng]="Nhập kho phụ"),"0",lookup(MaxROW("Bảng giá","_RowNumber",and([IDsp]=[_THISROW].[IDsp],[Ngày]>=[Ngày bắt đầu])),
"Bảng giá","IDgia","Giá khuyến cáo"))))


unique(select(khuyen_mai[id_qua],

  and(
  
  [id_sp]=[_THISROW].[id_sp],
  
  [so_luong_khach_can_lay]=[_THISROW].[chon_khuyen_mai],[loai_khuyen_mai]=[_THISROW].[loai_khuyen_mai],[tich_diem]=[_THISROW].[tich_diem],
  
  if(ISBLANK([_THISROW].[chon_kieu_gia]),true,([ap_dung_cho_loai_gia]=[_THISROW].[chon_kieu_gia])),
  
  or(
  [so_luong_khach_can_lay]<=[_THISROW].[so_luong],[tich_diem]="Có"),[ngay_bat_dau]<=date([_THISROW].[date]),
  [ngay_ket_thuc]>=date([_THISROW].[date]))))


  IF(LOOKUP(MAXROW("chi_tiet_don_hang","_RowNumber",AND(
    [id_kh]=[_THISROW].[id_kh],[id_sp]=[_THISROW].[id_sp],[phan_loai]="Đề xuất khuyến mại")),"chi_tiet_don_hang","id_ctdh","tich_diem")="Có",
    
    LOOKUP(MAXROW("khuyen_mai","_RowNumber",and(
    [id_qua]=INDEX([ma_qua],[dem_km]),
    [id_sp]=[_THISROW].[id_sp],
    [so_luong_khach_can_lay]=[_THISROW].[chon_khuyen_mai],
    [ap_dung_cho_loai_gia]=[_THISROW].[chon_kieu_gia]
    )),"khuyen_mai","id_qkm","so_luong"
    )*([so_luong]/[chon_khuyen_mai])+LOOKUP(MAXROW("chi_tiet_don_hang","_RowNumber",AND(
    [id_sp]=[_THISROW].[id_sp],[phan_loai]="Đề xuất khuyến mại",[tich_diem]="Có"
    
    )),"chi_tiet_don_hang","id_ctdh","so_luong_tich_diem"),
    
    LOOKUP(MAXROW("khuyen_mai","_RowNumber",and(
    [id_qua]=INDEX([ma_qua],[dem_km]),
    [id_sp]=[_THISROW].[id_sp],
    [so_luong_khach_can_lay]=[_THISROW].[chon_khuyen_mai],
    [ap_dung_cho_loai_gia]=[_THISROW].[chon_kieu_gia]
    )),"khuyen_mai","id_qkm","so_luong"
    )*([so_luong]/[chon_khuyen_mai]) 
    
    
    )


    FLOOR(IF(LOOKUP(MAXROW("chi_tiet_don_hang","_RowNumber",AND(
      [id_kh]=[_THISROW].[id_kh],[id_sp]=[_THISROW].[id_sp],[phan_loai]="Đề xuất khuyến mại")),"chi_tiet_don_hang","id_ctdh","tich_diem")="Có",
      
      LOOKUP(MAXROW("khuyen_mai","_RowNumber",and(
      [id_qua]=INDEX([ma_qua],[dem_km]),
      [id_sp]=[_THISROW].[id_sp],
      [so_luong_khach_can_lay]=[_THISROW].[chon_khuyen_mai],
      [ap_dung_cho_loai_gia]=[_THISROW].[chon_kieu_gia]
      )),"khuyen_mai","id_qkm","so_luong"
      )*([so_luong]/[chon_khuyen_mai])+LOOKUP(MAXROW("chi_tiet_don_hang","_RowNumber",AND(
      [id_sp]=[_THISROW].[id_sp],[phan_loai]="Đề xuất khuyến mại",[tich_diem]="Có"
      
      )),"chi_tiet_don_hang","IDctdh","so_luong tích điểm"),
      
      LOOKUP(MAXROW("khuyen_mai","_RowNumber",and(
      [id_qua]=INDEX([ma_qua],[dem_km]),
      [id_sp]=[_THISROW].[id_sp],
      [so_luong_khach_can_lay]=[_THISROW].[chon_khuyen_mai],
      [ap_dung_cho_loai_gia]=[_THISROW].[chon_kieu_gia]
      )),"khuyen_mai","id_qkm","so_luong"
      )*([so_luong]/[chon_khuyen_mai])  
      
      
      ))
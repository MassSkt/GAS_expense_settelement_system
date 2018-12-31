var sheet_id='';
var ignore_arr=['経理用','集計シート','勘定科目参照用','経費サマリ'];
var stat_sheet='集計シート';
var backup_folder_id='';
var folder_id='';


// 保存する
function save_as_pdf(){
  createPDF(folder_id, sheet_id, getTimestamp());
 
}


// 従業員の経費記入内容をオールクリアする
function clear_all_employee_contents(){
  // save backup before clearing
  createPDF(backup_folder_id, sheet_id, getTimestamp()+'_bkup');
  var sh_name=wmap_getSheetsName_NonHidden();
  for (  var i = 0;  i < sh_name.length;  i++  ) {    
    // clear contents if sheet name is about employee
    if (ignore_arr.indexOf(sh_name[i]) < 0){
      var spreadsheet = SpreadsheetApp.openById(sheet_id);
      var sheet=spreadsheet.getSheetByName(sh_name[i]);
      var lastRow=sheet.getLastRow();
      if (lastRow > 1){
        var range=sheet.getRange(2,1,lastRow-1,10);
        var content=range.clearContent();
      }
    }
  }
}

// 集計シートをオールクリアする
function clear_stat_contents(){
  var spreadsheet = SpreadsheetApp.openById(sheet_id);
  var sheet=spreadsheet.getSheetByName(stat_sheet);
  var lastRow=sheet.getLastRow();
  if (lastRow > 1){
    var range=sheet.getRange(2,1,lastRow-1,10);
    var content=range.clearContent();
  }
}


// 横ぐし計算用に、従業員のすべての記入内容を一枚のシートに記入する
function write_all_employee_stat(){
  //clear content before writing
  clear_stat_contents();
  var sh_name=wmap_getSheetsName_NonHidden();
  for (  var i = 0;  i < sh_name.length;  i++  ) {
    
    // write stat if sheet name is about employee
    if (ignore_arr.indexOf(sh_name[i]) < 0){
      var spreadsheet = SpreadsheetApp.openById(sheet_id);
      var sheet=spreadsheet.getSheetByName(sh_name[i]);
      var lastRow=sheet.getLastRow();
      if (lastRow > 1){
      var range=sheet.getRange(2,1,lastRow-1,7);
      var content=range.getValues();
      write_stat(sh_name[i],content);
      }
    }
  }
}

function write_stat(name,content_arr){
  var spreadsheet = SpreadsheetApp.openById(sheet_id);
  var statsheet=spreadsheet.getSheetByName(stat_sheet);
  var statlastRow=statsheet.getLastRow();
  var row_len=content_arr.length;
  var col_len=content_arr[0].length;
  var statrange=statsheet.getRange(statlastRow+1,2,row_len,col_len);
  var content=statrange.setValues(content_arr);
  var namerange=statsheet.getRange(statlastRow+1,1,row_len,1);
  var content_=namerange.setValue(name);

  
}


//表示されているシート名をすべて取得
function wmap_getSheetsName_NonHidden(){
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheet_names = new Array();
  
  if (sheets.length >= 1) {  
    for(var i = 0;i < sheets.length; i++)
    {
      if(!sheets[i].isSheetHidden())
      {
        sheet_names.push(sheets[i].getName());
      }
    }
  }
  return sheet_names;
}

// ファイルバックアップ
// PDF作成関数　引数は（folderid:保存先フォルダID, ssid:PDF化するスプレッドシートID, sheetid:PDF化するシートID, filename:PDFの名前）
function createPDF(folderid, ssid, filename){

  // PDFファイルの保存先となるフォルダをフォルダIDで指定
  var folder = DriveApp.getFolderById(folderid);

  // スプレッドシートをPDFにエクスポートするためのURL。このURLに色々なオプションを付けてPDFを作成
  var url = "https://docs.google.com/spreadsheets/d/SSID/export?".replace("SSID", ssid);

  // PDF作成のオプションを指定
  var opts = {
    exportFormat: "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
    format:       "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
    size:         "A4",     // 用紙サイズの指定 legal / letter / A4
    portrait:     "true",   // true → 縦向き、false → 横向き
    fitw:         "true",   // 幅を用紙に合わせるか
    sheetnames:   "true",  // シート名をPDF上部に表示するか
    printtitle:   "true",  // スプレッドシート名をPDF上部に表示するか
    pagenumbers:  "false",  // ページ番号の有無
    gridlines:    "false",  // グリッドラインの表示有無
    fzr:          "false"//,  // 固定行の表示有無
    //gid:          sheetid   // シートIDを指定 sheetidは引数で取得
  };
  
  var url_ext = [];
  
  // 上記のoptsのオプション名と値を「=」で繋げて配列url_extに格納
  for( optName in opts ){
    url_ext.push( optName + "=" + opts[optName] );
  }

  // url_extの各要素を「&」で繋げる
  var options = url_ext.join("&");

  // optionsは以下のように作成しても同じです。
  // var ptions = 'exportFormat=pdf&format=pdf'
  // + '&size=A4'                       
  // + '&portrait=true'                    
  // + '&sheetnames=false&printtitle=false' 
  // + '&pagenumbers=false&gridlines=false' 
  // + '&fzr=false'                         
  // + '&gid=' + sheetid;

  // API使用のためのOAuth認証
  var token = ScriptApp.getOAuthToken();

    // PDF作成
  var response = UrlFetchApp.fetch(url + options, {
      headers: {
        'Authorization': 'Bearer ' +  token
      }
    });

    // 
  var blob = response.getBlob().setName(filename + '.pdf');


  //　PDFを指定したフォルダに保存
  folder.createFile(blob);

}

 // タイムスタンプを返す関数
function getTimestamp () {
    var now = new Date();
    var year = now.getYear();
    var month = now.getMonth() + 1;
    var day = now.getDate();
    var hour = now.getHours();
    var min = now.getMinutes();
    // var sec = now.getSeconds();
    
    return year + "_" + month + "_" + day + "_" + hour + min;
}

var sheet_id='';
var ignore_arr=['�o���p','�W�v�V�[�g','����ȖڎQ�Ɨp','�o��T�}��'];
var stat_sheet='�W�v�V�[�g';
var backup_folder_id='';
var folder_id='';


// �ۑ�����
function save_as_pdf(){
  createPDF(folder_id, sheet_id, getTimestamp());
 
}


// �]�ƈ��̌o��L�����e���I�[���N���A����
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

// �W�v�V�[�g���I�[���N���A����
function clear_stat_contents(){
  var spreadsheet = SpreadsheetApp.openById(sheet_id);
  var sheet=spreadsheet.getSheetByName(stat_sheet);
  var lastRow=sheet.getLastRow();
  if (lastRow > 1){
    var range=sheet.getRange(2,1,lastRow-1,10);
    var content=range.clearContent();
  }
}


// �������v�Z�p�ɁA�]�ƈ��̂��ׂĂ̋L�����e���ꖇ�̃V�[�g�ɋL������
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


//�\������Ă���V�[�g�������ׂĎ擾
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

// �t�@�C���o�b�N�A�b�v
// PDF�쐬�֐��@�����́ifolderid:�ۑ���t�H���_ID, ssid:PDF������X�v���b�h�V�[�gID, sheetid:PDF������V�[�gID, filename:PDF�̖��O�j
function createPDF(folderid, ssid, filename){

  // PDF�t�@�C���̕ۑ���ƂȂ�t�H���_���t�H���_ID�Ŏw��
  var folder = DriveApp.getFolderById(folderid);

  // �X�v���b�h�V�[�g��PDF�ɃG�N�X�|�[�g���邽�߂�URL�B����URL�ɐF�X�ȃI�v�V������t����PDF���쐬
  var url = "https://docs.google.com/spreadsheets/d/SSID/export?".replace("SSID", ssid);

  // PDF�쐬�̃I�v�V�������w��
  var opts = {
    exportFormat: "pdf",    // �t�@�C���`���̎w�� pdf / csv / xls / xlsx
    format:       "pdf",    // �t�@�C���`���̎w�� pdf / csv / xls / xlsx
    size:         "A4",     // �p���T�C�Y�̎w�� legal / letter / A4
    portrait:     "true",   // true �� �c�����Afalse �� ������
    fitw:         "true",   // ����p���ɍ��킹�邩
    sheetnames:   "true",  // �V�[�g����PDF�㕔�ɕ\�����邩
    printtitle:   "true",  // �X�v���b�h�V�[�g����PDF�㕔�ɕ\�����邩
    pagenumbers:  "false",  // �y�[�W�ԍ��̗L��
    gridlines:    "false",  // �O���b�h���C���̕\���L��
    fzr:          "false"//,  // �Œ�s�̕\���L��
    //gid:          sheetid   // �V�[�gID���w�� sheetid�͈����Ŏ擾
  };
  
  var url_ext = [];
  
  // ��L��opts�̃I�v�V�������ƒl���u=�v�Ōq���Ĕz��url_ext�Ɋi�[
  for( optName in opts ){
    url_ext.push( optName + "=" + opts[optName] );
  }

  // url_ext�̊e�v�f���u&�v�Ōq����
  var options = url_ext.join("&");

  // options�͈ȉ��̂悤�ɍ쐬���Ă������ł��B
  // var ptions = 'exportFormat=pdf&format=pdf'
  // + '&size=A4'                       
  // + '&portrait=true'                    
  // + '&sheetnames=false&printtitle=false' 
  // + '&pagenumbers=false&gridlines=false' 
  // + '&fzr=false'                         
  // + '&gid=' + sheetid;

  // API�g�p�̂��߂�OAuth�F��
  var token = ScriptApp.getOAuthToken();

    // PDF�쐬
  var response = UrlFetchApp.fetch(url + options, {
      headers: {
        'Authorization': 'Bearer ' +  token
      }
    });

    // 
  var blob = response.getBlob().setName(filename + '.pdf');


  //�@PDF���w�肵���t�H���_�ɕۑ�
  folder.createFile(blob);

}

 // �^�C���X�^���v��Ԃ��֐�
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

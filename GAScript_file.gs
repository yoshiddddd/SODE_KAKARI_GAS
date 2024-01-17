const FORMID = '1vhpgbtBYuALfGScpPo51nxIhwWU_fcB725djJz6uBrM'
const FOLDERID = "1gn55O7k_KNDO8DhRoA9ItOfRZiPq9ZfB"
const TEMP_DOCID = '1pRgNHFsChSRTguv6ZRRBoxqaofITmtQFsIFBung4dww'
const SPSHEETID = '13fY4VV9gPcb9PmpRbT26lZ1T-jS79_4R4Q2TQ3azpGc'
function main() {
    var form = FormApp.openById(FORMID)
    var formResponses = form.getResponses()
    for (var i = 0; i < formResponses.length; i++) {
      var itemresponse = formResponses[i].getItemResponses();
      var dateResponse = itemresponse[0].getResponse();
    }
    const Folder = DriveApp.getFolderById(FOLDERID);
  if(itemresponse[3].getResponse() == "袖係を引き継ぎする")
  {
      ft_takeover(Folder);
      return ;
  }
    var date = dateResponse;
    // var date = Utilities.formatDate(date, "Asia/Tokyo", "MM/dd");
    const infiles = Folder.getFiles();
    const serch_filename  = '通し練習'+date;
    while(infiles.hasNext())
    {
      var file = infiles.next()
      if(file.getName() == serch_filename)
      {
        yetform(file,formResponses);
        return ;
      }
    }

    var documentid = TEMP_DOCID;
    var docCopyID = DriveApp.getFileById(documentid).makeCopy(serch_filename).getId();
    var doc = DocumentApp.openById(docCopyID);
    var body = doc.getBody()
    body.replaceText("日付", date);

    var num1 = 0;
    var num2 = 0;
    for (var i = 0; i < formResponses.length; i++) {
      for (var j = 1; j < 19; j++) {
        var mnumber = itemresponse[1].getResponse()
        if (mnumber == "M" + j) {
          num1 = j;
        }
      }
      for (var k = 1; k < 4; k++) {
        var place = itemresponse[2].getResponse()
        if (place == "上" + k) {
          num2 = k;
        }

        else if (place == "下" + k) {
          num2 = k + 3;
        }
      }

      body.replaceText("M" + num1 + "-" + num2, itemresponse[3].getResponse());
    }

    form.deleteAllResponses()
    const fileid = DriveApp.getFileById(doc.getId());
  fileid.moveTo(Folder);
}


function yetform(file,formResponses){
 var doc2 = DocumentApp.openById(file.getId());
var body= doc2.getBody();

  var num1 = 0;
  var num2 = 0;
  for (var i = 0; i < formResponses.length; i++) {
    var itemresponse = formResponses[i].getItemResponses()
    for (var j = 1; j < 19; j++) {
      var mnumber = itemresponse[1].getResponse()
      if (mnumber == "M" + j) {
        num1 = j;
      }
    }
    for (var k = 1; k < 4; k++) {
      var place = itemresponse[2].getResponse()
      if (place == "上" + k) {
        num2 = k;
      }

      else if (place == "下" + k) {
        num2 = k + 3;
      }
    }
    body.replaceText("M" + num1 + "-" + num2, itemresponse[3].getResponse());
  }
  // form.deleteAllResponses()
}

function ft_takeover(Folder)
{
    const files = Folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().includes('/')) {
      file.setTrashed(true);
    }
  }
  spreadsheetId =SPSHEETID; 
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheets()[0];
   var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    // 2行目以降の行を削除
    sheet.deleteRows(2, lastRow - 1);
  }
}

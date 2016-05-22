function myFunction() {
  var emails = 已回報清單(); //receive replies
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName("工作表2").activate();
  var sheet = SpreadsheetApp.getActiveSheet();

  sheet.getRange("D2:D").clear();

  var cell = sheet.getRange("E2");
  cell.setValue('=COUNTA(A2:A)');

  var startRow = 2;              // First row of data to process
  var startColumn = 1;
  var numRows = cell.getValue(); // Number of rows to process

  var data = sheet.getRange(startRow, startColumn, numRows).getValues(); //prepare list

  for (i in data) {
    var row = data[i];
    var emailAddress = row[0]; // email
    if (emails.indexOf(emailAddress) != -1) {
      sheet.getRange((startRow + parseInt(i)), 4).setValue('y');
      Logger.log(emailAddress + ' y');
    }
    else {
      Logger.log(emailAddress);
    }
  }
  寄信(sheet);
}

function 已回報清單() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName("表單回應1").activate();
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var d = new Date();
  var target = (d.getMonth() + 1) + "/" + d.getDate();
  //target = "12/24";
  var receiveData = []; // pick replies on the specific date
  for (i in data) {
    if (i != 0) {
      var row = data[i];
      var emailAddress = row[1];  // email
      var time = row[4];  // 填寫日期
      if (time === target) {
        receiveData.push(emailAddress);
      }
    }
  }
  return receiveData;
}

function 寄信(sheet) {
  var data = sheet.getDataRange().getValues(); //prepare list
  for (i in data) {
    if (i != 0) {
      var row = data[i];
      var is_reply = row[3];      // Column "D", y for reply
      if (is_reply === 'y') continue;

      var emailAddress = row[0];  // First column
      var message = row[1];       // Second column
      var subject = "Sending emails from a Spreadsheet";
      MailApp.sendEmail(emailAddress, subject, message);
    }
  }
}

function config() {

}

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

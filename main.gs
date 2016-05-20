function myFunction() {
  var emails = 已回報清單(); //receive replies
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var my_worksheet = ss.getSheetByName("工作表2");
  my_worksheet.activate();
  var sheet = SpreadsheetApp.getActiveSheet();

  sheet.getRange("D2:D").clear();

  var cell = sheet.getRange("E2");
  var data = sheet.getRange(2, 1, cell.getValue()).getValues(); //prepare list

  for (i in data) {
    var row = data[i];
    var emailAddress = row[0]; // email
    if (emails.indexOf(emailAddress) != -1) {
      sheet.getRange((2+parseInt(i)), 4).setValue('y');
      Logger.log(emailAddress + 'y');
    }
    else {
      Logger.log(emailAddress);
    }
  }
  寄信(sheet);
}

function 已回報清單() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var my_worksheet = ss.getSheetByName("表單回應1");
  my_worksheet.activate();
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
  //Logger.log(receiveData[0]);
  return receiveData;
}

function 寄信(sheet) {
  //var sheet = SpreadsheetApp.getActiveSheet();
  //var startRow = 2;  // First row of data to process
  //var numRows = 2;   // Number of rows to process

  // Fetch the range of cells A2:B3
  //var dataRange = sheet.getRange(startRow, 1, numRows, 1)
  // Fetch values for each row in the Range.
  //var data = dataRange.getValues();
  var data = sheet.getDataRange().getValues(); //prepare list
  for (i in data) {
    if (i != 0) {
      var row = data[i];
      var is_reply = row[3];      // y for reply
      if (is_reply === 'y') continue;
      var emailAddress = row[0];  // First column
      Logger.log(emailAddress);
      var message = row[1];       // Second column
      var subject = "Sending emails from a Spreadsheet";
      MailApp.sendEmail(emailAddress, subject, message);
    }
  }
}

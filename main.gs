function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = 2;   // Number of rows to process
  // Fetch the range of cells A2:B3
  //var dataRange = sheet.getRange(startRow, 1, numRows, 1)
  // Fetch values for each row in the Range.
  //var data = dataRange.getValues();
  var data = sheet.getDataRange().getValues();
  for (i in data) {
    if (i != 0) {
      var row = data[i];
      var emailAddress = row[0];  // First column
      var message = row[1];       // Second column
      var subject = "Sending emails from a Spreadsheet";
      MailApp.sendEmail(emailAddress, subject, message);
    }
  }
}

function lookupNoReceived() {
  // This example assumes there is a sheet named "first"
  var emails = lookupReceivedEmail();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var my_worksheet = ss.getSheetByName("sheet2");
  my_worksheet.activate();
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange("D2:D").clear();
  var cell = sheet.getRange("E2");
  var data = sheet.getRange(2, 1, cell.getValue()).getValues();
  
  for (i in data) {
    var row = data[i];
    var emailAddress = row[0];  // email
    Logger.log(emailAddress);
    if (emails.indexOf(emailAddress) != -1) {
      sheet.getRange((2+parseInt(i)), 4).setValue('y');
    }
  }
  
  //return data;
}

function lookupReceivedEmail() {
  //data = lookupNoReceived();
  //Logger.log(data);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var my_worksheet = ss.getSheetByName("sheet1");
  my_worksheet.activate();
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var d = new Date();
  var target = (d.getMonth() + 1) + "/" + d.getDate();
  target = "12/24";
  var receiveData = [];
  for (i in data) {
    if (i != 0) {
      var row = data[i];
      var emailAddress = row[11];  // email
      var time = row[1];  // 填寫日期
      if (time === target) {
        receiveData.push(emailAddress);
      }
    }
  }
  //Logger.log(receiveData[0]);
  return receiveData;
}

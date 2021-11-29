function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SEND E-MAIL')
    .addItem('Send Email', 'sendEmail')
    .addToUi();
}

function sendEmail() {
  var sheetName = "Sayfa1";
  var spreadsheet = SpreadsheetApp.openById('19k3Z2ThsKJnF4bvyK9RRWse-r7ccd5B7Fnsdj6ekOU8');
  var sheet = spreadsheet.getSheetByName(sheetName);
  ssTZ = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  console.log(ssTZ);
  var range = sheet.getRange("A1:A").getValues();
  var lastrow = range.filter(String).length;
  var lastCol = sheet.getLastColumn();

  for (var i = 1; i < lastCol + 1; i++) {
    if (sheet.getRange(1, i).getValue() == "Candidate name") { var candidateNameCol = i; }
    if (sheet.getRange(1, i).getValue() == "Swapped Name") { var swappedNameCol = i; }
    if (sheet.getRange(1, i).getValue() == "Last date to cancel") { var lastDateCol = i; }
    if (sheet.getRange(1, i).getValue() == "Swapped Ref No") { var swappedRefCol = i; }
    if (sheet.getRange(1, i).getValue() == "App Ref") { var appRefCol = i; }
    if (sheet.getRange(1, i).getValue() == "Date and time") { var dateTimeCol = i; }
    if (sheet.getRange(1, i).getValue() == "Test Centre") { var testCentreCol = i; }
    if (sheet.getRange(1, i).getValue() == "Email") { var emailCol = i; }
    if (sheet.getRange(1, i).getValue() == "Status") { var statusCol = i; }
    if (sheet.getRange(1, i).getValue() == "Confirmation Sent") { var confirmCol = i; }
  }

  for (var i = 2; i < lastrow + 1; i++) {
    if (sheet.getRange(i, statusCol).getValue().toLowerCase() == "sold" && sheet.getRange(i, confirmCol).getValue().toLowerCase() == "no") {
      if (sheet.getRange(i, swappedNameCol).getValue() == "") {
        var candidateName = sheet.getRange(i, candidateNameCol).getValue();
      } else {
        var candidateName = sheet.getRange(i, swappedNameCol).getValue();
      }
      var status = sheet.getRange(i, statusCol).getValue();
      var dateTime = (sheet.getRange(i, dateTimeCol).getValue());
      console.log(dateTime);
      var date = Utilities.formatDate(dateTime,Session.getScriptTimeZone(),'MMMM dd, yyyy');
      var date2 = Utilities.formatDate(dateTime,Session.getScriptTimeZone(),'dd MMMM yyyy');
      console.log(date);
      var time = Utilities.formatDate(dateTime,Session.getScriptTimeZone(),'HH:mm');
      console.log(time);
      var testCentre = sheet.getRange(i, testCentreCol).getValue();
      if (sheet.getRange(i, swappedRefCol).getValue() == "") {
        var refNo = sheet.getRange(i, appRefCol).getValue();
      } else {
        var refNo = sheet.getRange(i, swappedRefCol).getValue();
      }
      var lastDate = sheet.getRange(i, lastDateCol).getValue();
      var lastDate = Utilities.formatDate(lastDate,Session.getScriptTimeZone(),'dd MMMM yyyy');
      var emailAddress = sheet.getRange(i, emailCol).getValue();
      var subject = "Driving test booking confirmation: " + date2;
      console.log(candidateName, status, dateTime, refNo, lastDate, emailAddress, subject);
      var htmlOutput = HtmlService.createHtmlOutputFromFile('emailBody'); // Message is the name of the HTML file

      var message = htmlOutput.getContent()
      message = message.replace("%name", candidateName);
      message = message.replace("%date", date);
      message = message.replace("%time", time);
      message = message.replace("%testcentre", testCentre);
      message = message.replace("%refno", refNo);
      message = message.replace("%lastdate", lastDate);
      // SpreadsheetApp.getUi().alert("Approval process");
      try {
        MailApp.sendEmail(emailAddress, subject, message, { htmlBody: message });
        sheet.getRange(i, confirmCol).setValue("YES");
      }
      catch (err) {
        alert(i + " 'th cant send");
      }
    }
  }
}

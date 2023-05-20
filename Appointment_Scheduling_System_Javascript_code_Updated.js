var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var lastRow = sheet.getLastRow();
var lastColumn = sheet.getLastColumn();
var calendar = CalendarApp.getCalendarById('parvsoni2003@gmail.com');


// function to read all the values
function Submission(row){
  this.timestamp = sheet.getRange(row, 1).getValue();
  this.name = sheet.getRange(row, 2).getValue();
  this.reason = sheet.getRange(row, 3).getValue();
  this.date = new Date(sheet.getRange(row, 4).getValue());
  this.dateString = (this.date.getMonth() + 1) + '/' + 
    this.date.getDate() + '/' + this.date.getFullYear();
  this.time = sheet.getRange(row,5).getValue();
  this.timeString = this.time.toLocaleTimeString();
  this.email = sheet.getRange(row, 6).getValue();
  this.suggesteddate = new Date (sheet.getRange(row, 7).getValue());
  this.suggesteddatestring = (this.suggesteddate.getMonth()+1) + '/' + 
  this.suggesteddate.getDate() + '/' + this.suggesteddate.getFullYear();
  // Adjust time and make end time
  this.date.setHours(this.time.getHours());
  this.date.setMinutes(this.time.getMinutes());
  this.endTime = new Date(this.date);
  this.endTime.setHours(this.time.getHours() + 1);
}


// Function to know conflict of more than one event in a day
function getProblems(request) {
   var conflicts = calendar.getEvents(request.date, request.endTime);
  if (conflicts.length < 1) {
    request.status = "New";
  } else {
    request.status = "Conflict";
    sheet.getRange(lastRow, lastColumn - 1).setValue("Reject");
    sheet.getRange(lastRow, lastColumn).setValue("Sent: Conflict");
    sheet.getRange(lastRow, lastColumn-2).setValue("-------------");
  }
}


// function to know what type of e-mail id going to sent
function structEmail(request){
  request.buttonLink = "https://docs.google.com/forms/d/e/1FAIpQLSdhICQxoLbyh4c6_dXLzZigh4t2B4A6hPe42HvGHX0yf225Xg/viewform?usp=sf_link"
  request.buttonText = "New Request";
  switch (request.status) {
    case "New":
      request.subject = "Request for " + request.dateString + " Appointment Received";
      request.header = "Request Received";
      request.message = "Once the request has been reviewed you will receive an email updating you on it.";
      break;
    case "New2":
      request.email = "parvsoni2003@gmail.com";
      request.subject = "New Request for " + request.dateString;
      request.header = "Request Received";
      request.message = "A new request needs to be reviewed.";
      request.buttonLink = "https://docs.google.com/spreadsheets/d/15AbfWFRi6YvTJirZlqm3lOhu9pkTAyv8EtY9gQUcQkY/edit?usp=sharing";
      request.buttonText = "View Request";
      break;
    case "Approve":
      request.subject = "Confirmation: Appointment for " + request.dateString + " has been scheduled";
      request.header = "Confirmation";
      request.message = "Your appointment has been scheduled.";
      break;
    case "Conflict":
      request.subject = "Conflict with " + request.dateString + " Appointment Request";
      request.header = "Conflict";
      request.message = "There was a scheduling conflict. Please choose new date or time.";
      request.buttonText = "Reschedule";
      break;
    case "Reject":
      request.subject = "Update on Appointment Requested for " + request.dateString;
      request.header = "Reschedule";
      request.message = "Unfortunately the requested time does not work. Could "+
        "we reschedule?"+"Suggested Date - "+request.suggesteddatestring;
      request.buttonText = "Reschedule";
      break;
  }
}


// function to create event on google calendar
function updateCalendar(request){
  var event = calendar.createEvent(
    request.name,
    request.date,
    request.endTime
    )
}


//  mail sending function
function sendmail(request){
  MailApp.sendEmail({
    to: request.email,
    subject: request.subject,
    htmlBody: makeEmail(request)
  })
}


// request received and reviewing email sending functions 
function onFormSubmission() {
  var request = new Submission(lastRow);
  getProblems(request);
  structEmail(request);
  Logger.log(request.status);
  sendmail(request);
  if(request.status == "New"){
    request.status = "New2";
    structEmail(request);
    sendmail(request);
  }
}


// converting columns data in array to no need of referencing again to spreadsheet 
function StatusObject(){
  this.statusArray = sheet.getRange(1, lastColumn -1, lastRow, 1).getValues();
  this.notifiedArray = sheet.getRange(1, lastColumn, lastRow, 1).getValues();
  this.statusArray = [].concat.apply([], this.statusArray);
  this.notifiedArray = [].concat.apply([], this.notifiedArray);
}


// updating column values in last column after reading from last second column
function getChange(statusgetChange){
  statusgetChange.index = statusgetChange.notifiedArray.indexOf("");
  statusgetChange.row = statusgetChange.index + 1;
  if (statusgetChange.index == -1){
    return;
  } else if (statusgetChange.statusArray[statusgetChange.index] != "") {
    statusgetChange.status = statusgetChange.statusArray[statusgetChange.index];
    sheet.getRange(statusgetChange.row, lastColumn).setValue("Sent: " + statusgetChange.status);
    statusgetChange.notifiedArray[statusgetChange.index] = "update";
  } else {
    statusgetChange.status = statusgetChange.statusArray[statusgetChange.index];
    statusgetChange.notifiedArray[statusgetChange.index] = "no update";
  }
}


// function to send email on approving or rejecting and update calendar on approving any appointment
function onEdit(){
  var statusChange = new StatusObject();
  while (true){
    getChange(statusChange);
    if (statusChange.index == -1){
      return;
    } else {
      var request = new Submission(statusChange.row);
      if (statusChange.status){
        request.status = statusChange.status;
        if (statusChange.status == "Approve"){
          updateCalendar(request);
          sheet.getRange(lastRow, lastColumn-2).setValue("-------------");
        }
        structEmail(request);
        sendmail(request);
      }
    }
  }
}

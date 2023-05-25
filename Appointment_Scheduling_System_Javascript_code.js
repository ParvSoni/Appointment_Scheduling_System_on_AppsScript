var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var lastRow = sheet.getLastRow();
var lastColumn = sheet.getLastColumn();
var calendar = CalendarApp.getCalendarById('calender id');


// function to read all the values
function Submission(row){
  this.timestamp = sheet.getRange(row, 1).getValue();
  this.name = sheet.getRange(row, 2).getValue();
  this.reason = sheet.getRange(row, 3).getValue();
  this.date = new Date(sheet.getRange(row, 4).getValue());
  this.dateString = (this.date.getMonth() + 1) + '/' + 
    this.date.getDate() + '/' + this.date.getYear();
  this.time = sheet.getRange(row,5).getValue();
  this.timeString = this.time.toLocaleTimeString();
  this.email = sheet.getRange(row, 6).getValue();
  // Adjust time and make end time
  this.date.setHours(this.time.getHours());
  this.date.setMinutes(this.time.getMinutes());
  this.endTime = new Date(this.date);
  this.endTime.setHours(this.time.getHours() + 1);
}


// Function to know conflict of more than one event in a day
function getConflicts(request) {
   var conflicts = calendar.getEvents(request.date, request.endTime);
  if (conflicts.length < 1) {
    request.status = "New";
  } else {
    request.status = "Conflict";
    sheet.getRange(lastRow, lastColumn - 1).setValue("Reject");
    sheet.getRange(lastRow, lastColumn).setValue("Sent: Conflict");
  }
}


// function to know what type of e-mail id going to sent
function draftEmail(request){
  request.buttonLink = "Form link"
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
      request.buttonLink = "Google sheet link";
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
      request.message = "There was a scheduling conflict. Please reschedule.";
      request.buttonText = "Reschedule";
      break;
    case "Reject":
      request.subject = "Update on Appointment Requested for " + request.dateString;
      request.header = "Reschedule";
      request.message = "Unfortunately the request times does not work. Could "+
        "we reschedule?";
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
function sendEmail(request){
  MailApp.sendEmail({
    to: request.email,
    subject: request.subject,
    htmlBody: makeEmail(request)
  })
}


// request received and reviewing email sending functions 
function onFormSubmission() {
  var request = new Submission(lastRow);
  getConflicts(request);
  draftEmail(request);
  Logger.log(request.status);
  sendEmail(request);
  if(request.status == "New"){
    request.status = "New2";
    draftEmail(request);
    sendEmail(request);
  }
}


// converting columns data in array to no need of referncing again to spreadsheet 
function StatusObject(){
  this.statusArray = sheet.getRange(1, lastColumn -1, lastRow, 1).getValues();
  this.notifiedArray = sheet.getRange(1, lastColumn, lastRow, 1).getValues();
  this.statusArray = [].concat.apply([], this.statusArray);
  this.notifiedArray = [].concat.apply([], this.notifiedArray);
}


// updating column values in last column after reading from last second column
function getChangeIndex(statusChange){
  statusChange.index = statusChange.notifiedArray.indexOf("");
  statusChange.row = statusChange.index + 1;
  if (statusChange.index == -1){
    return;
  } else if (statusChange.statusArray[statusChange.index] != "") {
    statusChange.status = statusChange.statusArray[statusChange.index];
    sheet.getRange(statusChange.row, lastColumn).setValue("Sent: " + statusChange.status);
    statusChange.notifiedArray[statusChange.index] = "update";
  } else {
    statusChange.status = statusChange.statusArray[statusChange.index];
    statusChange.notifiedArray[statusChange.index] = "no update";
  }
}


// function to send email on approving or rejecting and update calendar on approving any appointment
function onEdit(){
  var statusChange = new StatusObject();
  while (true){
    getChangeIndex(statusChange);
    if (statusChange.index == -1){
      return;
    } else {
      var request = new Submission(statusChange.row);
      if (statusChange.status){
        request.status = statusChange.status;
        if (statusChange.status == "Approve"){
          updateCalendar(request);
        }
        draftEmail(request);
        sendEmail(request);
      }
    }
  }
}

//Code written by Matthew Gendreau
//If you need help with this contact me at [REDACTED]
//https://pastebin.com/q7buvzZt
//Compiled using code from http://blog.ouseful.info/2010/03/04/maintaining-google-calendars-from-a-google-spreadsheet/ and https://mashe.hawksey.info/2010/03/google-apps-spreadsheet-2-calendar-site/

function processEvents() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var cal=CalendarApp.getCalendarsByName("Capital District Webmaster")[0]; //Replace "Capital District Webmaster" with the name of your calendar
  var sheet=SpreadsheetApp.setActiveSheet(ss.getSheets()[0]); //need a routine to set active sheet by name?
  //go defensive
  var col_broadcast,col_itunes, col_youtube=1; //Leftover code from what I copied that I'm too afraid to get rid of, delete at your own risk
  var maxcols=sheet.getMaxColumns();

  var startRow = 2;  // First row of data to process
  var numRows = 100;   // Number of rows to process | If the calendar stops working, try increasing this number or deleting earlier entries
  var dataRange = sheet.getRange(startRow, 1, numRows, 26);
  var data = dataRange.getValues();
  
  for (i in data) { //Identifies each column by row (furthest left row is row zero)
    var row = data[i];
    var title = row[1];
    var desc=row[5];
    var added = row[0];  //Check to see if details for this event have been added to the calendar(s)
    var tstart = row[2]; //start time - I have defined the column in the spreadsheet as a Date type
    var tstop = row[3]; //start time - I have defined the column in the spreadsheet as a Date type
    var loc = row[4];
    
  if (added!="Added") { //the calendar(s) have not been updated for this event
    CalendarApp.createEvent(title, tstart, tstop, {location:loc, description:desc}); // create calendar event
  }
  var v = parseInt(i)+2; //No idea what this is for but it makes it work
  sheet.getRange(v, 1, 1, 1).setValue("Added"); //set the fact that we have updated the calendars for this event
  }
}

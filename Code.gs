//Code written by Matthew Gendreau
//If you need help with this contact me at [REDACTED]
//Compiled using code from http://blog.ouseful.info/2010/03/04/maintaining-google-calendars-from-a-google-spreadsheet/ and https://mashe.hawksey.info/2010/03/google-apps-spreadsheet-2-calendar-site/

//10/24/18 Edit: New comments are marked with *** at the beginning

function processEvents() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  //Replace "NameHere" with the name of your calendar sheet
  var cal=CalendarApp.getCalendarsByName("NameHere")[0];
  //Uses the sheet at the index of getSheets() I think?
  var sheet=SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  //Leftover code from what I copied that I'm too afraid to get rid of, delete at your own risk
  var col_broadcast,col_itunes, col_youtube=1;
  var maxcols=sheet.getMaxColumns();

  //First row of data to process
  var startRow = 2;
  //Number of rows to process | If the calendar stops working, try increasing this number or deleting earlier entries
  //*** Note that it doesn't process all of these rows at once, this is just specifying the range that the script is equipped to read.
  //*** Theoretically you could try making this number as large as possible, but I'm not sure to what extent you can do that
  //*** without breaking something.
  var numRows = 100;
  //startRow is the y coord of your top left cell, 1 is the x coord, numRows is the # of rows your range covers, 26 is # of columns
  var dataRange = sheet.getRange(startRow, 1, numRows, 26);
  var data = dataRange.getValues();
  
  //Identifies each column by row (furthest left row is row zero)
  for (i in data) {
    var row = data[i];
    var title = row[1];
    var desc=row[5];
    var added = row[0];  //Check to see if details for this event have been added to the calendar(s)
    //*** Further explanation: added basically acts as a boolean that tells the script to ignore the field once it has been successfully
    //*** added to the calendar.
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

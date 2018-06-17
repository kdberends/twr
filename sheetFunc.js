function appendToSheet(sheetName, data) {
  var ActiveSheet = SpreadsheetApp.getActive().getSheetByName(sheetName)
  var values = ActiveSheet.getDataRange().getValues();
  
  ActiveSheet.insertRowBefore(values.length + 1);
  
  newdata = []
  newdata.push(data)
  ActiveSheet.getRange(values.length + 1, 1, 1, data.length).setValues(newdata);
}


function createSheet(name) {
  var activeSpreadsheet = SpreadsheetApp.getActive()
  var yourNewSheet = activeSpreadsheet.getSheetByName(name);

    if (yourNewSheet != null) {
        activeSpreadsheet.deleteSheet(yourNewSheet);
    }

    yourNewSheet = activeSpreadsheet.insertSheet();
    yourNewSheet.setName(name);
    header = []
    header.push(["Student", "Klas", "Tijdvak", "Aanwezig", "Te laat", "Absent"])
    var cells = yourNewSheet.getRange(1, 1, 1, 6)
    cells.setValues(header).setValues(header).setBackgroundColor("#09b5a3")
    cells.setFontWeight("bold").setBorder(true, null, true, null, false, false, "Black", SpreadsheetApp.BorderStyle.SOLID_THICK)
}

/* Schedule per timeslot
 * Parses the 'planning' sheet and returns the 'schedule' array. Each
 * item in the array is an object with info on the schedule item. 
 * Number of students per timeslot is retrieved from the linked form
 *
 *
 * 
 * 
 */
function retrieveSchedule(){ 
  var schedule = [];
  
  // Open sheet and load info
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('planning');
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  // See if there is a sheet linked to form
  var sheets = ss.getSheets()
  var FormSheet = null
  for (i in sheets){ 
      var linkedform = sheets[i].getFormUrl()
      if (linkedform != null){var FormSheet = sheets[i]}
  }
  Logger.log('linked form: ' + FormSheet)
  
  
  for (var i = 1; i < values.length; i++) {
    // each row has following syntax: [day, timeslot, course1, instructor1, classroom1, course2, ....]
    var row = values[i];  
    var day = row[0]; // Date
    var time = row[1]; // Timeslot
    
    // Generate a unique name for the session (will be used as questions in the form for students)   
    var session = DAYS[day.getDay() - 1] + ", " + time; 
   
    //  parse options
    numberOfOptions = (row.length - 2) / 3;
    
    var options = []
    for (var j = 0; j < numberOfOptions; j++){
      var numberofstudents = 0
      var course = row[j * 3 + 2]
      if (FormSheet != null){
        // get first row of return form and find column of header
        sessionrow = FormSheet.getRange(1, 1, 1, 150).getValues()[0]
        column = findMatchingIdInArray(session, sessionrow) + 1
        sessioncolumn = FormSheet.getRange(1, column, 20, 1).getValues()
        numberofstudents = countIf(course, sessioncolumn)
      }
      options.push({course:course, instructor:row[j * 3 + 3], classroom:row[j * 3 + 4], numberofstudents:numberofstudents})
    }
    // Add session to schedule
    if (!schedule[session]) {
      schedule.push({name: session, date: day, time: time, options:options});
    }
  }
  
  
  return schedule  
}

/* Schedule per student
 * Parses the linked form to retrieve schedule per user
 * Returns array. Each entry is object with attributes
 * 'name' - string, 
 * 'class' - string,
 * 'email' - string,
 * 'response' array, each entry:
 *    day
 *    session
 *    choice 
 */
function retrieveStudentSchedule(){ 
  var schedule = [];
  
  // Open spreadsheet
  var ss = SpreadsheetApp.getActive();
  
  // Check whether there is a sheet linked to form
  var sheets = ss.getSheets()
  var FormSheet = null
  for (i in sheets){ 
      var linkedform = sheets[i].getFormUrl()
      if (linkedform != null){var FormSheet = sheets[i]}
  }
  Logger.log('Found linked form: ' + FormSheet)

  // Read header (choises)
  
  // Read all students (assuming no more than 998 students)
  var students = FormSheet.getDataRange().getValues();
  var header = students[0]
  for (var row = 1; row < students.length; row ++){
    schedule.push({name: students[row][1],
                   email: students[row][2],
                   class: students[row][3],
                   response: []
                   });
      for (var i = 4; i < students[row].length; i ++){
        options = {day:header[i].split(',')[0].trim(),
                   session:header[i].split(',')[1].trim(),
                   choice:students[row][i] 
                  };
        schedule.slice(-1)[0].response.push(options);
      }
    }
  
  //Logger.log(students)
  
  return schedule  
}


/* Counts only if condition is met
 * Similar to excel countif function
 */
function countIf(condition, list){
  var count = 0
  for (i in list){
    if (list[i] == condition){
      count ++
    }
  }
  return count
}

/* Will remove all falsy values: undefined, null, 0, false, NaN and "" (empty string)
*/
function cleanArray(actual) {
  var newArray = new Array();
  for (var i = 0; i < actual.length; i++) {
    if (actual[i]) {
      newArray.push(actual[i]);
    }
  }
  return newArray;
}

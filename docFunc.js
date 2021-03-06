/**
 * This function generates an overview schedule of all
 * courses given. This overview is meant for instructors
 * or schedulers. 
 * 
 * The document is generated in the OutputFolder. 
 * 
 * @param {String[][]} OutputFolder Name of the folder the document should 
 *                                 be put into
 */
function generateInstructorDocument(OutputFolder){
  
  
  // Copy the template
  // TODO: Give rooster ID instead? If multiple files exist with same name 
  // wrong one could be copied
  var docid = DriveApp.getFilesByName("RoosterTemplate")
                      .next()
                      .makeCopy(SETTINGS.title_overviewschedule)
                      .getId();
  
  // Open the document 
  var doc = DocumentApp.openById(docid)
  var body = doc.getBody();
  
  // Define a custom paragraph style to maximimze usage of document space
  var style = {};
  style[DocumentApp.Attribute.MARGIN_LEFT] = 10
  style[DocumentApp.Attribute.MARGIN_RIGHT] = 10
  style[DocumentApp.Attribute.FONT_SIZE] = 8;
  style[DocumentApp.Attribute.LINE_SPACING] = 1;
  body.setAttributes(style)
  
  // to landscape (in points https://www.google.nl/webhp#newwindow=1&q=8,27+inch+in+point)
  body.setPageHeight(595.276).setPageWidth(841.89);
  
  // Table Header
  var table = [['Tijdvak']];
  for (i in DAYS){table[0].push(DAYS[i])}
  
  // Fill table with first column and empty rest
  for (var i = 0; i < SESSIONS.length; i++){
    var row = new Array(DAYS.length + 1).join(';').split(';')
    row[0] = SESSIONS[i].name
    table.push(row)
  }
  
  // sessionsnames used for look up later
  var sessionnames = []
    for (i in SESSIONS){
      sessionnames.push(SESSIONS[i].name[0])
    }
  
  // Fill table with data from schedule
  var schedule = retrieveSchedule()
  for (i in schedule){
    columnId = schedule[i].date.getDay()
    rowId = findMatchingIdInArray(schedule[i].time, sessionnames) 
    if (columnId!=null && rowId!= null){
      for (j in schedule[i].options){
        if (schedule[i].options[j].course.length > 1){
          if (j != 0){table[rowId + 1][columnId] += "\n"}
          text = Utilities.formatString("%s, %s, %s (%s)",
                                        getIdofCourse(schedule[i].options[j].course), 
                                        schedule[i].options[j].instructor, 
                                        schedule[i].options[j].classroom,
                                        schedule[i].options[j].numberofstudents);
          table[rowId + 1][columnId] += text;
          
        }
        
      }
    }
  }

  // Fill document with text, table
  body.insertParagraph(0, Utilities.formatString('Rooster voor %s %s', INFO.weektitle, INFO.weeknumber))
      .setHeading(DocumentApp.ParagraphHeading.HEADING1).editAsText().setFontFamily("Comfortaa");
  table = body.appendTable(table);
  
  // Color first row
  for (var i = 0; i < table.getRow(0).getNumCells(); i++){table.getRow(0).getCell(i).setBackgroundColor('#107896')}
  table.getRow(0).editAsText().setBold(true).setForegroundColor('#F2F3F4');
  
  // Style other rows
  for (var i = 0; i< table.getNumRows(); i++){
    for (var j = 0; j<  table.getRow(i).getNumCells();j++){
      table.getRow(i).getCell(j).editAsText().setFontFamily("Courier New");
    }
  }
  
  // Add Legend (course id: id)
  var legend = '';
  for (i=0;i<COURSES.length;i++){
    legend += Utilities.formatString('%s: %s\n', COURSES[i].id, COURSES[i].name)
  };
  body.insertParagraph(body.getNumChildren(), legend)

  // Close file and move to given output folder
  var docId = doc.getId();
  doc.saveAndClose();
  Logger.log("Document ${docId} saved and closed")
  moveFileToAnotherFolder(docId, OutputFolder)
  return docId
}


/**
 * For a coursename, return the ID
 */
function getIdofCourse(coursename){
  for (var i=0;i<COURSES.length;i++){
    if (COURSES[i].name == coursename){
      return COURSES[i].id
    }
  }
}

/* We use triviumlogo from web. 
 * TODO: store logo locally? This may give problems if 
 * external location changes..
*/
function getTriviumMavoLogoBlob() {
  // Retrieve an image from the web.
  var resp = UrlFetchApp.fetch("https://www.triviumcollege.nl/files/images/cache/7f5bffc0abbf95f3f7c8759ad348ee3206e40fd0.jpg");
  return resp.getBlob();
}

/**
 * This function generates a preliminary document
 * based on choices of the student.
 * 
 * The document is generated in the OutputFolder. 
 * 
 * Argument:
 * OutputFolder: google driver folder ID
 * 
 * Returns:
 * Google document ID
 */
function generatePreliminaryDocument(user) {
  
  // Create and share a personalized Google Doc that shows the student's chosen schedule
  var docid = DriveApp.getFilesByName("RoosterTemplate")
                      .next()
                      .makeCopy(Utilities.formatString('Voorlopig Rooster voor %s, %s', user.name, user.class))
                      .getId();
  var doc = DocumentApp.openById(docid)
  var body = doc.getBody();
  // to landscape (in points https://www.google.nl/webhp#newwindow=1&q=8,27+inch+in+point)
  doc.getBody().setPageHeight(595.276).setPageWidth(841.89);
  
  // Define a custom paragraph style.
  var style = {};
  style[DocumentApp.Attribute.MARGIN_BOTTOM] = 10
  style[DocumentApp.Attribute.MARGIN_TOP] = 10
  style[DocumentApp.Attribute.MARGIN_LEFT] = 10
  style[DocumentApp.Attribute.MARGIN_RIGHT] = 10
  style[DocumentApp.Attribute.FONT_SIZE] = 8;
  body.setAttributes(style)
  
  // Build table out of student response
  // =======================================
  // Header
  var table = [['Tijdvak']];
  for (i in DAYS){table[0].push(DAYS[i])}
  
  // Fill table with first column and empty rest
  for (var i = 0; i < SESSIONS.length; i++){
    var row = new Array(DAYS.length + 1).join('; ').split(';')
    row[0] = SESSIONS[i].name
    table.push(row)
  }
  
  // sessionsnames for look up
  var sessionnames = []
    for (i in SESSIONS){
      sessionnames.push(SESSIONS[i].name[0])
    }
  
  // Loop through student response and add to schedule
  for (i in user.response){
    columnId = findMatchingIdInArray(user.response[i].day, DAYS)
    rowId = findMatchingIdInArray(user.response[i].session, sessionnames) 
    Logger.log("row: " + rowId, " Column; " + columnId)
    if (columnId!=null && rowId!= null){
      table[rowId + 1][columnId + 1] = user.response[i].choice
    }
  }
  
  // Fill document
  body.insertParagraph(0, "Voorlopig programma")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(Utilities.formatString("%s | %s, %s", user.name, user.class, user.email));
  body.appendParagraph("Dit is je voorlopige persoonlijk programma." 
                     + "Dit programma is gevuld met jouw keuzes. Let op:"
                     + "dit rooster is niet definitief en kan nog veranderen.");
  table = body.appendTable(table);
  
  // Background color first row
  for (var i = 0; i < table.getRow(0).getNumCells(); i++){table.getRow(0).getCell(i).setBackgroundColor('#107896')}
  
  table.getRow(0).editAsText().setBold(true).setForegroundColor('#F2F3F4');
  var docId = doc.getId();
  doc.saveAndClose();
  
  return docId
}


/**
 * This function generates the final document
 * based on choices of the student.
 * 
 * The document is generated in the OutputFolder. 
 * 
 * Argument:
 * OutputFolder: google driver folder ID
 * 
 * Returns:
 * Google document ID
 *
 * TODO: merge generatePreliminaryDocument and 
 *       generateFinalScheduleForUser into one
 *       more general func?
 *
 */
function generateFinalScheduleForUser(user) {
  
  // Create and share a personalized Google Doc that shows the student's chosen schedule
  var docid = DriveApp.getFilesByName("RoosterTemplate").next().makeCopy(Utilities.formatString('Definitief Rooster voor %s, %s', user.name, user.class)).getId();
  var doc = DocumentApp.openById(docid)
  var body = doc.getBody();
  // to landscape (in points https://www.google.nl/webhp#newwindow=1&q=8,27+inch+in+point)
  doc.getBody().setPageHeight(595.276).setPageWidth(841.89);
  
  // Define a custom paragraph style.
  var style = {};
  style[DocumentApp.Attribute.MARGIN_BOTTOM] = 10
  style[DocumentApp.Attribute.MARGIN_TOP] = 10
  style[DocumentApp.Attribute.MARGIN_LEFT] = 10
  style[DocumentApp.Attribute.MARGIN_RIGHT] = 10
  style[DocumentApp.Attribute.FONT_SIZE] = 8;
  body.setAttributes(style)
  
  // Build table out of student response
  // =======================================
  
  // Header
  var table = [['Tijdvak']];
  for (i in DAYS){table[0].push(DAYS[i])}
  
  // Fill table with first column and empty rest
  for (var i = 0; i < SESSIONS.length; i++){
    var row = new Array(DAYS.length + 1).join('; ').split(';')
    row[0] = SESSIONS[i].name
    table.push(row)
  }
  
  // sessionsnames for look up
  var sessionnames = []
    for (i in SESSIONS){
      sessionnames.push(SESSIONS[i].name[0])
    }
  
  // Loop through student response and add to schedule
  for (i in user.response){
    columnId = findMatchingIdInArray(user.response[i].day, DAYS)
    rowId = findMatchingIdInArray(user.response[i].session, sessionnames) 
    Logger.log("row: " + rowId, " Column; " + columnId)
    if (columnId!=null && rowId!= null){
      var details = GetDetailsForCourse(user.response[i].day, user.response[i].session, user.response[i].choice)
      table[rowId + 1][columnId + 1] = Utilities.formatString("%s\n%s, %s", user.response[i].choice, details.instructor, details.classroom)
    }
  }
  
  // Fill document
  body.insertParagraph(0, "Rooster voor " + INFO.weektitle + " " + INFO.weeknumber)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(Utilities.formatString("%s | %s, %s", user.name, user.class, user.email));
  table = body.appendTable(table);
  
  // Background color first row
  for (var i = 0; i < table.getRow(0).getNumCells(); i++){table.getRow(0).getCell(i).setBackgroundColor('#107896')}
  
  table.getRow(0).editAsText().setBold(true).setForegroundColor('#F2F3F4');
  var docId = doc.getId();
  doc.saveAndClose();
  
  return docId
}


/* Generate attendance form. Currently, just
 * copies templates and filles with list of names
 * of students registered for course. 
 * TODO: make more fancy
 */
function generateAttendanceForm(course, day, time) {
  // Create and share a personalized Google Doc that shows the student's chosen schedule
  
  var docid = DriveApp.getFilesByName("RoosterTemplate")
                      .next()
                      .makeCopy(Utilities.formatString('Presentielijst voor %s, %s, %s', course, day, time))
                      .getId();
  var doc = DocumentApp.openById(docid)
  var body = doc.getBody();

  // to landscape (in points https://www.google.nl/webhp#newwindow=1&q=8,27+inch+in+point)
  doc.getBody().setPageHeight(595.276).setPageWidth(841.89);
  
  // Define a custom paragraph style.
  var style = {};
  style[DocumentApp.Attribute.MARGIN_BOTTOM] = 10
  style[DocumentApp.Attribute.MARGIN_TOP] = 10
  style[DocumentApp.Attribute.MARGIN_LEFT] = 10
  style[DocumentApp.Attribute.MARGIN_RIGHT] = 10
  style[DocumentApp.Attribute.FONT_SIZE] = 8;
  body.setAttributes(style)
  
  // Build table out of student response
  // =======================================
  var students = retrieveStudentsPerCourseSession(course, day, time)
  Logger.log(students)
  var msg = ''
  for (var i=0;i<students.length;i++){
    msg += Utilities.formatString("%s, %s\n", students[i].name, students[i].class)
  }
  
  body.insertParagraph(0, "Aanwezigheidslijst")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(msg)
  
  var docId = doc.getId();
  doc.saveAndClose();
  
  return docId
}

/* For given day, timeslot and coursename, returns
 * which instructor and classroom are listed. 
 */
function GetDetailsForCourse(day, session, course){
  Logger.log('requestion details for '+day+session+course)
  var schedule = retrieveSchedule()
  for (var i=0; i < schedule.length; i++){
    if (DAYS[schedule[i].date.getDay() - 1]==day){
      if (schedule[i].time==session){
        for (var j=0; j < schedule[i].options.length;j++){
          if (schedule[i].options[j].course == course){
            return {instructor:schedule[i].options[j].instructor, classroom:schedule[i].options[j].classroom}
          }
        }
      }
    }
  }
  // If this line is executed, an invalid combination is found. This may
  // be due to a typo after correction of the preliminary schedules. this
  // will throw an error message. 
  Browser.msgBox(Utilities.formatString('FOUTMELDING: Er is een ongeldige combinatie gevonden: %s, %s, %s', day, session, course))
}
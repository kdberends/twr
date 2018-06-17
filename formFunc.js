/** =============================================================================================
 * Creates a Google Form that allows students to select which 
 * sessions they would like to attend.
 *
 * @param {Spreadsheet} ss The spreadsheet that contains the conference data.
 * @param {String[][]} values Cell values for the spreadsheet range.
 */
function setUpForm_(ss, values) {
  // Group the sessions by date and time so that they can be passed to the form.
  var schedule = retrieveSchedule()
  var parentFolderId = getParentFolderOfFile(ss)
  
  // Create the form from template
  var formid = DriveApp.getFilesByName("FormTemplate").next().makeCopy('Inschrijfformulier').getId();
  
  // New files are created at root. move to same folder as spreadsheet
  moveFileToAnotherFolder(formid, parentFolderId);

  // Open the form
  var form = FormApp.openById(formid);
     
  // Build form
  form.addSectionHeaderItem().setTitle('Triviumweek-planner');
  form.addTextItem().setTitle('Name').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);
  var item = form.addListItem().setTitle('Klas').setChoiceValues(CLASSES).setRequired(true)
  
  // Activiteiten op de maandag, aparte sectie, eveneens dinsdag, woensdag ....
  Logger.log('starting loop to fill form')
  for (i = 0; i < DAYS.length; i++){
    form.addPageBreakItem().setTitle(DAYS[i]); 
    for (j in schedule) {
      // Loop over days
      if (schedule[j].date.getDay() == i + 1) {
        var courseOptions = [] 
        Logger.log(schedule[j])
        // Loop over timeslots
        for (var k = 0; k < schedule[j].options.length; k++){
          if (schedule[j].options[k].course.length > 1){
            courseOptions.push(schedule[j].options[k].course)
          }
        }
        Logger.log('Options: ' + courseOptions)
        if (courseOptions.length > 1){
        form.addListItem()
        .setTitle(schedule[j].name)
        .setChoiceValues(courseOptions);
        }
      }
    } 
  }
  Logger.log("Form created")

  // Responses should go to current sheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  
  return form.getId()
}


/**
 * Function executed on submission of form:
 *  - Adds student answers to course-specific sheet
 *  - Emails students 
 *
 *
 * @param {Object} e The event parameter for form submission to a spreadsheet;
 *     see https://developers.google.com/apps-script/understanding_events
 */
function onFormSubmit(e) {
  Logger.log("A user has submitted a form")
  
  // schedule from Trivium_planning
  var schedule = retrieveSchedule()
  
  // contact info of student
  var user = {name: e.namedValues["Name"][0], 
              email: e.namedValues["Email"][0],
              class:e.namedValues["Klas"][0], response: []};
  Logger.log('Started logging for ' + user.name)

  // Fills response for user
  for (var i = 0; i < schedule.length; i++){
    Logger.log(schedule[i].name)
    chosenCourse = e.namedValues[schedule[i].name]
    Logger.log(chosenCourse)
    if (chosenCourse[0].length > 0){
    //appendToSheet(chosenCourse[0], [user.name, user.class, schedule[i].name])
    user.response.push({day:schedule[i].name.split(',')[0].trim(), 
                        session: schedule[i].name.split(',')[1].trim(), 
                        choice: e.namedValues[schedule[i].name]})
    }
  }
  
  // Generate personalised schedule for student
  Logger.log('generating preliminary schedule for '+user.name)
  var ss = SpreadsheetApp.getActive();
  var docId = generatePreliminaryDocument(user);

  // Move personalised schedule to designated subfolder
  moveFileToAnotherFolder(docId, getSubfolder(SETTINGS.schedules_folder_preliminary, getParentFolderOfFile(ss)));

  // Email student
  Logger.log('sendign email to ' + user.email)
  var attachment = DriveApp.getFileById(docId);
  //var htmlBody = HtmlService.createHtmlOutputFromFile('email_prelim.html').getContent();
  var template = HtmlService.createTemplateFromFile('email_prelim.html')
  template.name = user.name;
  htmlBody = template.evaluate().getContent();
  MailApp.sendEmail({
    to: user.email,
    subject: "Jouw voorlopige Triviumweekrooster",
    htmlBody: htmlBody,
    attachments:  [attachment.getAs(MimeType.PDF)]
  });
}


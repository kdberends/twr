/** =============================================================================================
 * Creates a Google Form that allows respondents to select which conference
 * sessions they would like to attend, grouped by date and start time.
 *
 * @param {Spreadsheet} ss The spreadsheet that contains the conference data.
 * @param {String[][]} values Cell values for the spreadsheet range.
 */
function setUpForm_(ss, values) {
  // Group the sessions by date and time so that they can be passed to the form.
  
  var schedule = retrieveSchedule()
  var parentFolderId = getParentFolderOfFile(ss)
  
  // Create the form
  var formid = DriveApp.getFilesByName("FormTemplate").next().makeCopy('Inschrijfformulier').getId();
  moveFileToAnotherFolder(formid, parentFolderId);
  var form = FormApp.openById(formid);
  //var form = FormApp.create("Triviumtestweek").setDescription("Stel hier zelf je rooster samen");
  
  // Move the newly created form to 'Trivium' folder
  
   
  // Responses should go to current sheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  
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
      if (schedule[j].date.getDay() == i + 1) {
        var courseOptions = [] 
        Logger.log(schedule[j])
        for (var k = 0; k < schedule[j].options.length; k++){
          if (schedule[j].options[k].course.length > 1){
            courseOptions.push(schedule[j].options[k].course)
          }
        }
        Logger.log('Options: ' + courseOptions)
        form.addListItem().setTitle(schedule[j].name).setChoiceValues(courseOptions);
      }
    } 
  }
  Logger.log("form set-up")
  return form.getId()
}


/**
 * A trigger-driven function that sends out calendar invitations and a
 * personalized Google Docs itinerary after a user responds to the form.
 *
 * @param {Object} e The event parameter for form submission to a spreadsheet;
 *     see https://developers.google.com/apps-script/understanding_events
 */
function onFormSubmit(e) {
  Logger.log("form submitted by a user")
  
  // schedule from Trivium_planning
  var schedule = retrieveSchedule()
  Logger.log("Retrieved Schedules")
  // contact info of student
  var user = {name: e.namedValues["Name"][0], 
              email: e.namedValues["Email"][0],
              class:e.namedValues["Klas"][0], response: []};
  Logger.log('Started logging for '+user.name)
  // Add student to teacher overview
  for (var i = 0; i < schedule.length; i++){
    Logger.log(schedule[i].name)
    chosenCourse = e.namedValues[schedule[i].name]
    Logger.log(chosenCourse)
    if (chosenCourse[0].length > 0){
    appendToSheet(chosenCourse[0], [user.name, user.class, schedule[i].name])
    user.response.push({day:schedule[i].name.split(',')[0].trim(), 
                        session: schedule[i].name.split(',')[1].trim(), 
                        choice: e.namedValues[schedule[i].name]})
    }
  }
  Logger.log(user.response)
  Logger.log('generating voorlopig programma for '+user.name)
  var ss = SpreadsheetApp.getActive();
  var docId = generatePreliminaryDocument(user);
  moveFileToAnotherFolder(docId, getSubfolder("Leerlingroosters", getParentFolderOfFile(ss)));
  
}


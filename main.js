/** README
* ------------------------------------------------------
* This script is written to automatically create a form 
* for students to choose their own schedule for a given
* week, mail students their preliminary schedule and 
* provide instructors an overview of students per course.
*
* The spreadsheet this script is attached to should have 
* a sheet called 'planning'. The 'planning' sheet has the
* following format:
* 
*   column 1: "Date"
*   column 2: "Tijdvak"
*   columns 3 and onward: each column a possible choice 
*                         for students
*
* This script adds the menu item "Trivium". 
* Based on: https://developers.google.com/apps-script/quickstart/forms
*
* -------------------------------------------------------
* Contact: mariekebloo@triviumcollege.com 
*
* Authors
*    - Marieke Berends-Bloo
*    - Koen Berends
*
* -------------------------------------------------------
* Copyright 2018 
*
*
*/

// global variables
var COURSES = getCOURSES()
var SESSIONS = getSESSIONS()
var INFO = getINFO()
var CLASSES = ['M1A', 'M1B', 'M2A', 'M2B', 'M3A', 'M3B', 'M4A', 'M4B']
var DAYS = ['Maandag', 'Dinsdag', 'Woensdag', 'Donderdag', 'Vrijdag']
var DAYS = ['Maandag', 'Dinsdag'];
var SETTINGS = {schedules_folder_preliminary: "VoorlopigeRoosters", schedules_folder_final: "DefinitieveRoosters"}

function getINFO(){
  var ss = SpreadsheetApp.getActive();
  var weektitle = ss.getRange("Middelen!C2").getValues()
  var weeknumber = ss.getRange("Middelen!C3").getValues()
  return {weektitle:weektitle, weeknumber:weeknumber}
}

function getCOURSES(){
  var ss = SpreadsheetApp.getActive();
  
  // names
  var range = ss.getRange("Middelen!F8:F42");
  var range_input = range.getValues()
  var names = [i for each (i in range_input)if (isNaN(i))];
  
  // ids
  var range = ss.getRange("Middelen!E8:E42");
  var range_input = range.getValues()
  var ids = [i for each (i in range_input)if (isNaN(i))];

  // output
  courses = []
  for (i in names){
    courses.push({id:ids[i], name:names[i]})
  }
  return courses
}

function getSESSIONS(){
  var ss = SpreadsheetApp.getActive();
  // names
  var range = ss.getRange("Middelen!C8:C25");
  var range_input = range.getValues()
  var names = [i for each (i in range_input)if (isNaN(i))];
  
  // ids
  var range = ss.getRange("Middelen!B8:B25");
  var range_input = range.getValues()
  var ids = [i for each (i in range_input)if (isNaN(i))];
  
  // output
  sessions = []
  for (i in names){
    sessions.push({id:i, name:names[i]})
  }
  return sessions
}

/**
 * A special function that inserts the custom menu when the spreadsheet opens.
 */
function onOpen() {
  
  var menuEntries = [];
  
  // Menu item to create form
  menuEntries.push({name: 'Genereer overzichtsrooster', functionName: 'genInstructorDocMenu'});
  
  // Menu item to create form for students
  menuEntries.push({name: 'Genereer studentinschrijfformulier', functionName: 'setUpTrivium_'});
  
  // Menu item to create final schedules after students have filled in forms
  menuEntries.push({name: 'Maak definitieve studentroosters', functionName: 'generateSchedulesFromSheet_'});
  menuEntries.push({name: 'Stuur definitieve studentroosters', functionName: 'sendSchedulesFromSheet_'});
  
  // Menu item to display information on how to use app
  menuEntries.push({name: 'Reset sheet', functionName: 'resetApp_'});
  
  // Menu item to display information on how to use app
  menuEntries.push({name: 'Show help sidebar', functionName: 'showHelp_'});
  
  SpreadsheetApp.getActive().addMenu('Trivium', menuEntries);
}

/**
 * A set-up function that uses the planning in the spreadsheet to create
 * a Google Form, and a trigger that allows the script
 * to react to form responses.
 */
function setUpTrivium_() {
  // Get current active Spreadsheet file
  var ss = SpreadsheetApp.getActive();
  
  // Get schedule from planning
  var sheet = ss.getSheetByName('planning');
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  // Reset sheets
  resetApp_()
  
  // Create a sheet for each course
  for (var i = 0; i<COURSES.length; i++) {
    createSheet(COURSES[i].name)
  }
  
  // Create form for students
  formid = setUpForm_(ss, values);
  Browser.msgBox('Het invulformulier is aangemaakt')
  Logger.log("Form created")
  
  // Trigger for form submit. Note: current spreadsheet MUST be set as destination for
  // this form. Otherwise, the results won't log properly
  deleteTriggers()
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  Logger.log("Triggers assigned")
  
  return true
}

/**
 * Resets the spreadsheet to pristine state. 
 * The user is warned with a popup before command
 * is executed
 */
function resetApp_(){
  var ss = SpreadsheetApp.getActive();
  // Warn user that this action will delete all current information in the sheet
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Let op! Als u hiermee doorgaat wordt het document gereset.'
                        + 'Alle informatie, met uitzondering de sheets "Middelen" en '
                        + '"planning" gaat verloren. Weet u zeker dat u dat wilt?', ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    Logger.log('The user clicked "Yes."');
  

  // Delete all sheet except planning
  var sheets = ss.getSheets()
  for (i in sheets){
    
  if (sheets[i].getSheetName() != "planning" && sheets[i].getSheetName() != "Middelen") {
    // If sheet already has linked form, unlink and delete it
    var linkedform = sheets[i].getFormUrl()
    if (linkedform != null){FormApp.openByUrl(linkedform).removeDestination()}
    // Delete sheet
    Logger.log(Utilities.formatString("Deleting sheet %s", sheets[i].getSheetName()))
    ss.deleteSheet(sheets[i])
      
    }
  }
  
  Logger.log("Sheets deleted") 
  } else {
    Logger.log('The user clicked "No" or the close button in the dialog\'s title bar.');
    return null
  }
}

/**
 * Creates an overview schedule for instructors. If students
 * have already filled out the form, displays how many students
 * have registered for a course
 */
function genInstructorDocMenu(){
  var ss = SpreadsheetApp.getActive();
  Logger.log('Sending command to make instr. doc')
  var docId = generateInstructorDocument(getParentFolderOfFile(ss));
  moveFileToAnotherFolder(docId, getParentFolderOfFile(ss));
  Browser.msgBox("Overzichtsrooster is aangemaakt")
}

function generateSchedulesFromSheet_(){
  Browser.msgBox('Deze functie is nog niet ingebouwd')
}

function sendSchedulesFromSheet_(){
  Browser.msgBox('Deze functie is nog niet ingebouwd')
}

function showHelp_(){
  // Display a modal dialog box with custom HtmlService content.
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Tivium Rooster App Help')
      //.setWidth(300); // lijkt niets te doen..
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);
}

// =============================================================================================
// 
// =============================================================================================


function deleteTriggers() {
  // Deletes all triggers in the current project.
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
 
}

function findMatchingIdInArray(value, array){
  for (i in array){
    if (array[i] == value){return parseInt(i)}
  }
  return null
}


// =============================================================================================
// end of script
// =============================================================================================
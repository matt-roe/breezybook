function doPost(e) {
    try {
        var ss = SpreadsheetApp.openById("1LIgR-YYyEXfWdRZ7IAixpBroKAOCx76fHGnG3jfG8lw");
        var sheet = ss.getSheetByName('bookings');
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var holderArray = [];
        for (var x = 0; x < headers.length; x++) {
            var tempValue = !e.parameter[headers[x]] ? ' ' : e.parameter[headers[x]];
            holderArray.push(tempValue);
        }
        sheet.appendRow(holderArray);
        var results = {
            "data": e.parameter
            , "holder": holderArray
        }
        var jsonData = JSON.stringify(results)
        return ContentService.createTextOutput(jsonData).setMimeType(ContentService.MimeType.JSON)
    }
    catch (e) {
        var error = {
            "error": e
        }
        var jsonError = JSON.stringify(error)
        return ContentService.createTextOutput(jsonError).setMimeType(ContentService.MimeType.JSON)
    }
}

function doGet(e) {
    //return ContentService.createTextOutput('Hello World');
    try {
        var ss = SpreadsheetApp.openById("1LIgR-YYyEXfWdRZ7IAixpBroKAOCx76fHGnG3jfG8lw");
        var sheet = ss.getSheetByName('booked');
        var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
        var jsonData = JSON.stringify(data)
        return ContentService.createTextOutput(jsonData).setMimeType(ContentService.MimeType.JSON)
    }
    catch (e) {
        var error = {
            "error": e
        }
        var jsonError = JSON.stringify(error)
        return ContentService.createTextOutput(jsonError).setMimeType(ContentService.MimeType.JSON)
    }
}

//modified from starting https://stackoverflow.com/questions/40845359/google-apps-script-get-an-array-of-dates-between-two-dates?rq=1
var DAY_MILLIS = 24 * 60 * 60 * 1000;

function createDateSpan(startDate, endDate) {
if (startDate === "") {
return ""
}
if (endDate === "") {
var curDate = new Date(startDate.getTime());
curDate = new Date(curDate.getTime() + DAY_MILLIS);
var singleDay = Utilities.formatDate(curDate, "MST", 'MM/dd/yyyy')
return singleDay
}
if (startDate.getTime() > endDate.getTime()) {
throw Error('Start is later than end');
}

var dates = [];

var curDate = new Date(startDate.getTime());
curDate = new Date(curDate.getTime() + DAY_MILLIS);
endDate = new Date(endDate.getTime() + DAY_MILLIS);
while (!dateCompare(curDate, endDate)) {
dates.push(Utilities.formatDate(curDate, "MST", 'MM/dd/yyyy'));
curDate = new Date(curDate.getTime() + DAY_MILLIS);
}
dates.push(Utilities.formatDate(endDate, "MST", 'MM/dd/yyyy'));
var dateArray = dates.toString().split(",").join();
return dateArray;
}
function dateCompare(a, b) {
return a.getFullYear() === b.getFullYear() &&
a.getMonth() === b.getMonth() &&
a.getDate() === b.getDate();
}

// Script to synchronize a calendar to a spreadsheet and vice versa.
//
// See https://github.com/Davepar/gcalendarsync for instructions on setting this up.
//

// Set this value to match your calendar!!!
// Calendar ID can be found in the "Calendar Address" section of the Calendar Settings.
var calendarId = 'b76lop0cmgce83ijbcp7ls5pbk@group.calendar.google.com';
var calendar2Id = 'nb5f1bqlo36h06oemhkb92r6hie8onrp@import.calendar.google.com';

// Set the beginning and end dates that should be synced. beginDate can be set to Date() to use
// today. The numbers are year, month, date, where month is 0 for Jan through 11 for Dec.
var firstDate = new Date(); // Today
var secondDate = new Date();
secondDate.setDate(firstDate.getDate()+300); // Today + 10 months

var beginDate = firstDate;
var endDate = secondDate;

// Date format to use in the spreadsheet.
var dateFormat = 'mm/dd/yyy';

var titleRowMap = {
'title': 'Title',
'description': 'Payment',
'starttime': 'Checkin',
'endtime': 'Checkout',
'guests': 'Email',
'color': 'Color',
'id': 'Id'
};
var titleRowKeys = ['title', 'description', 'starttime', 'endtime', 'guests', 'color', 'id'];
var requiredFields = ['id', 'title', 'starttime', 'endtime'];

// This controls whether email invites are sent to guests when the event is created in the
// calendar. Note that any changes to the event will cause email invites to be resent.
var SEND_EMAIL_INVITES = true;

// Setting this to true will silently skip rows that have a blank start and end time
// instead of popping up an error dialog.
var SKIP_BLANK_ROWS = false;

// Updating too many events in a short time period triggers an error. These values
// were tested for updating 40 events. Modify these values if you're still seeing errors.
var THROTTLE_THRESHOLD = 1;
var THROTTLE_SLEEP_TIME = 300;

// Adds the custom menu to the active spreadsheet.
function onOpen() {
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var ui = SpreadsheetApp.getUi();
ui.createMenu("Auto Trigger")
.addItem("Run","runAuto")
.addToUi();
var menuEntries = [
{
  name: "Update from Calendar",
  functionName: "syncFromCalendar"
}, {
  name: "Update to Calendar",
  functionName: "syncToCalendar"
}
];
spreadsheet.addMenu('Calendar Sync', menuEntries);
}

// https://github.com/benlcollins/apps_script/blob/master/auto_trigger/code.gs

function runAuto() {

// resets the loop counter if it's not 0
refreshUserProps();

// create trigger to run program automatically
createTrigger();
}


// -----------------------------------------------------------------------------
// reset loop counter to 0 in properties
// -----------------------------------------------------------------------------
function refreshUserProps() {
var userProperties = PropertiesService.getUserProperties();
userProperties.setProperty('loopCounter', 0);
}

// -----------------------------------------------------------------------------
// create trigger to run addNumber every minute
// -----------------------------------------------------------------------------
function createTrigger() {
// Trigger every 5 minute
ScriptApp.newTrigger('syncToCalendar')
  .timeBased()
  .everyMinutes(30)
  .create();
}

// -----------------------------------------------------------------------------
// function to delete triggers
// -----------------------------------------------------------------------------
function deleteTrigger() {

// Loop over all triggers and delete them
var allTriggers = ScriptApp.getProjectTriggers();

for (var i = 0; i < allTriggers.length; i++) {
ScriptApp.deleteTrigger(allTriggers[i]);
}

}

// Creates a mapping array between spreadsheet column and event field name
function createIdxMap(row) {
var idxMap = [];
for (var idx = 0; idx < row.length; idx++) {
var fieldFromHdr = row[idx];
for (var titleKey in titleRowMap) {
  if (titleRowMap[titleKey] == fieldFromHdr) {
    idxMap.push(titleKey);
    break;
  }
}
if (idxMap.length <= idx) {
  // Header field not in map, so add null
  idxMap.push(null);
}
}
return idxMap;
}

// Converts a spreadsheet row into an object containing event-related fields
function reformatEvent(row, idxMap, keysToAdd) {
var reformatted = row.reduce(function(event, value, idx) {
if (idxMap[idx] != null) {
  event[idxMap[idx]] = value;
}
return event;
}, {});
for (var k in keysToAdd) {
reformatted[keysToAdd[k]] = '';
}
return reformatted;
}

// Converts a calendar event to a psuedo-sheet event.
function convertCalEvent(calEvent) {
convertedEvent = {
'id': calEvent.getId(),
'title': calEvent.getTitle(),
'description': calEvent.getDescription(),
'location': calEvent.getLocation(),
'guests': calEvent.getGuestList().map(function(x) {return x.getEmail();}).join(','),
'color': calEvent.getColor()
};
if (calEvent.isAllDayEvent()) {
convertedEvent.starttime = calEvent.getAllDayStartDate();
var endtime = calEvent.getAllDayEndDate();
if (endtime - convertedEvent.starttime === 24 * 3600 * 1000) {
  convertedEvent.endtime = '';
} else {
  convertedEvent.endtime = endtime;
  if (endtime.getHours() === 0 && endtime.getMinutes() == 0) {
    convertedEvent.endtime.setSeconds(endtime.getSeconds() - 1);
  }
}
} else {
convertedEvent.starttime = calEvent.getStartTime();
convertedEvent.endtime = calEvent.getEndTime();
}
return convertedEvent;
}

// Converts calendar event into spreadsheet data row
function calEventToSheet(calEvent, idxMap, dataRow) {
convertedEvent = convertCalEvent(calEvent);

for (var idx = 0; idx < idxMap.length; idx++) {
if (idxMap[idx] !== null) {
  dataRow[idx] = convertedEvent[idxMap[idx]];
}
}
}

// Returns empty string or time in milliseconds for Date object
function getEndTime(ev) {
return ev.endtime === '' ? '' : ev.endtime.getTime();
}

// Tests whether calendar event matches spreadsheet event
function eventMatches(cev, sev) {
var convertedCalEvent = convertCalEvent(cev);
return convertedCalEvent.title == sev.title &&
convertedCalEvent.description == sev.description &&
convertedCalEvent.location == sev.location &&
convertedCalEvent.starttime.toString() == sev.starttime.toString() &&
getEndTime(convertedCalEvent) === getEndTime(sev) &&
convertedCalEvent.guests == sev.guests &&
convertedCalEvent.color == ('' + sev.color);
}

// Determine whether required fields are missing
function areRequiredFieldsMissing(idxMap) {
return requiredFields.some(function(val) {
return idxMap.indexOf(val) < 0;
});
}

// Returns list of fields that aren't in spreadsheet
function missingFields(idxMap) {
return titleRowKeys.filter(function(val) {
return idxMap.indexOf(val) < 0;
});
}

// Set up formats and hide ID column for empty spreadsheet
function setUpSheet(sheet, fieldKeys) {
sheet.getRange(1, fieldKeys.indexOf('starttime') + 1, 999).setNumberFormat(dateFormat);
sheet.getRange(1, fieldKeys.indexOf('endtime') + 1, 999).setNumberFormat(dateFormat);
sheet.hideColumns(fieldKeys.indexOf('id') + 1);
}

// Display error alert
function errorAlert(msg, evt, ridx) {
var ui = SpreadsheetApp.getUi();
if (evt) {
ui.alert('Skipping row: ' + msg + ' in event "' + evt.title + '", row ' + (ridx + 1));
} else {
ui.alert(msg);
}
}

// Updates a calendar event from a sheet event.
function updateEvent(calEvent, sheetEvent){
sheetEvent.sendInvites = SEND_EMAIL_INVITES;
if (sheetEvent.endtime === '') {
calEvent.setAllDayDate(sheetEvent.starttime);
} else {
calEvent.setTime(sheetEvent.starttime, sheetEvent.endtime);
}
calEvent.setTitle(sheetEvent.title);
calEvent.setDescription(sheetEvent.description);
calEvent.setLocation(sheetEvent.location);
// Set event color
if (sheetEvent.color > 0 && sheetEvent.color < 12) {
calEvent.setColor('' + sheetEvent.color);
}
var guestCal = calEvent.getGuestList().map(function (x) {
return {
  email: x.getEmail(),
  added: false
};
});
var sheetGuests = sheetEvent.guests || '';
var guests = sheetGuests.split(',').map(function (x) {
return x ? x.trim() : '';
});
// Check guests that are already invited.
for (var gIx = 0; gIx < guestCal.length; gIx++) {
var index = guests.indexOf(guestCal[gIx].email);
if (index >= 0) {
  guestCal[gIx].added = true;
  guests.splice(index, 1);
}
}
guests.forEach(function (x) {
if (x) calEvent.addGuest(x);
});
guestCal.forEach(function (x) {
if (!x.added) {
  calEvent.removeGuest(x.email);
}
});
}

// Synchronize from calendar to spreadsheet.
function syncFromCalendar() {
// Get calendar and events
var calendar = CalendarApp.getCalendarById(calendarId);
var calEvents = calendar.getEvents(beginDate, endDate);

// Get spreadsheet and data
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet = spreadsheet.getActiveSheet();
var range = sheet.getDataRange();
var data = range.getValues();
var eventFound = new Array(data.length);

// Check if spreadsheet is empty and add a title row
var titleRow = [];
for (var idx = 0; idx < titleRowKeys.length; idx++) {
titleRow.push(titleRowMap[titleRowKeys[idx]]);
}
if (data.length < 1) {
data.push(titleRow);
range = sheet.getRange(1, 1, data.length, data[0].length);
range.setValues(data);
setUpSheet(sheet, titleRowKeys);
}

if (data.length == 1 && data[0].length == 1 && data[0][0] === '') {
data[0] = titleRow;
range = sheet.getRange(1, 1, data.length, data[0].length);
range.setValues(data);
setUpSheet(sheet, titleRowKeys);
}

// Map spreadsheet headers to indices
var idxMap = createIdxMap(data[0]);
var idIdx = idxMap.indexOf('id');

// Verify header has all required fields
if (areRequiredFieldsMissing(idxMap)) {
var reqFieldNames = requiredFields.map(function(x) {return titleRowMap[x];}).join(', ');
Logger.log('Spreadsheet must have ' + reqFieldNames + ' columns');
return;
}

// Array of IDs in the spreadsheet
var sheetEventIds = data.slice(1).map(function(row) {return row[idIdx];});

// Loop through calendar events
for (var cidx = 0; cidx < calEvents.length; cidx++) {
var calEvent = calEvents[cidx];
var calEventId = calEvent.getId();

var ridx = sheetEventIds.indexOf(calEventId) + 1;
if (ridx < 1) {
  // Event not found, create it
  ridx = data.length;
  var newRow = [];
  var rowSize = idxMap.length;
  while (rowSize--) newRow.push('');
  data.push(newRow);
} else {
  eventFound[ridx] = true;
}
// Update event in spreadsheet data
calEventToSheet(calEvent, idxMap, data[ridx]);
}

// Remove any data rows not found in the calendar
var rowsDeleted = 0;
for (var idx = eventFound.length - 1; idx > 0; idx--) {
//event doesn't exists and has an event id
if (!eventFound[idx] && sheetEventIds[idx - 1]) {
  data.splice(idx, 1);
  rowsDeleted++;
}
}

// Save spreadsheet changes
range = sheet.getRange(1, 1, data.length, data[0].length);
range.setValues(data);
if (rowsDeleted > 0) {
sheet.deleteRows(data.length + 1, rowsDeleted);
}
}

// Synchronize from calendar 2 to spreadsheet.
function syncFromCalendar2() {
// Get calendar and events
var calendar = CalendarApp.getCalendarById(calendar2Id);
var calEvents = calendar.getEvents(beginDate, endDate);

// Get spreadsheet and data
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet = spreadsheet.getActiveSheet();
var range = sheet.getDataRange();
var data = range.getValues();
var eventFound = new Array(data.length);

// Check if spreadsheet is empty and add a title row
var titleRow = [];
for (var idx = 0; idx < titleRowKeys.length; idx++) {
titleRow.push(titleRowMap[titleRowKeys[idx]]);
}
if (data.length < 1) {
data.push(titleRow);
range = sheet.getRange(1, 1, data.length, data[0].length);
range.setValues(data);
setUpSheet(sheet, titleRowKeys);
}

if (data.length == 1 && data[0].length == 1 && data[0][0] === '') {
data[0] = titleRow;
range = sheet.getRange(1, 1, data.length, data[0].length);
range.setValues(data);
setUpSheet(sheet, titleRowKeys);
}

// Map spreadsheet headers to indices
var idxMap = createIdxMap(data[0]);
var idIdx = idxMap.indexOf('id');

// Verify header has all required fields
if (areRequiredFieldsMissing(idxMap)) {
var reqFieldNames = requiredFields.map(function(x) {return titleRowMap[x];}).join(', ');
Logger.log('Spreadsheet must have ' + reqFieldNames + ' columns');
return;
}

// Array of IDs in the spreadsheet
var sheetEventIds = data.slice(1).map(function(row) {return row[idIdx];});

// Loop through calendar events
for (var cidx = 0; cidx < calEvents.length; cidx++) {
var calEvent = calEvents[cidx];
var calEventId = calEvent.getId();

var ridx = sheetEventIds.indexOf(calEventId) + 1;
if (ridx < 1) {
  // Event not found, create it
  ridx = data.length;
  var newRow = [];
  var rowSize = idxMap.length;
  while (rowSize--) newRow.push('');
  data.push(newRow);
} else {
  eventFound[ridx] = true;
}
// Update event in spreadsheet data
calEventToSheet(calEvent, idxMap, data[ridx]);
}

// Remove any data rows not found in the calendar
var rowsDeleted = 0;
for (var idx = eventFound.length - 1; idx > 0; idx--) {
//event doesn't exists and has an event id
if (!eventFound[idx] && sheetEventIds[idx - 1]) {
  data.splice(idx, 1);
  rowsDeleted++;
}
}

// Save spreadsheet changes
range = sheet.getRange(1, 1, data.length, data[0].length);
range.setValues(data);
if (rowsDeleted > 0) {
sheet.deleteRows(data.length + 1, rowsDeleted);
}
}

// Synchronize from spreadsheet to calendar.
function syncToCalendar() {
// Get calendar and events
var calendar = CalendarApp.getCalendarById(calendarId);
if (!calendar) {
Logger.log('Cannot find calendar. Check instructions for set up.');
}
var calEvents = calendar.getEvents(beginDate, endDate);
var calEventIds = calEvents.map(function(val) {return val.getId();});

// Get spreadsheet and data
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet = spreadsheet.getActiveSheet();
var range = sheet.getDataRange();
var data = range.getValues();
if (data.length < 2) {
Logger.log('Spreadsheet must have a title row and at least one data row');
return;
}

// Map headers to indices
var idxMap = createIdxMap(data[0]);
var idIdx = idxMap.indexOf('id');
var idRange = range.offset(0, idIdx, data.length, 1);
var idData = idRange.getValues()

// Verify header has all required fields
if (areRequiredFieldsMissing(idxMap)) {
var reqFieldNames = requiredFields.map(function(x) {return titleRowMap[x];}).join(', ');
Logger.log('Spreadsheet must have ' + reqFieldNames + ' columns');
return;
}

var keysToAdd = missingFields(idxMap);

// Loop through spreadsheet rows
var numChanges = 0;
var numUpdated = 0;
var changesMade = false;
for (var ridx = 1; ridx < data.length; ridx++) {
var sheetEvent = reformatEvent(data[ridx], idxMap, keysToAdd);

// If enabled, skip rows with blank/invalid start and end times
if (SKIP_BLANK_ROWS && !(sheetEvent.starttime instanceof Date) &&
    !(sheetEvent.endtime instanceof Date)) {
  continue;
}

// Do some error checking first
if (!sheetEvent.title) {
  Logger.log('must have title', sheetEvent, ridx);
  continue;
}
if (!(sheetEvent.starttime instanceof Date)) {
  Logger.log('start time must be a date/time', sheetEvent, ridx);
  continue;
}
if (sheetEvent.endtime !== '') {
  if (!(sheetEvent.endtime instanceof Date)) {
    Logger.log('end time must be empty or a date/time', sheetEvent, ridx);
    continue;
  }
  if (sheetEvent.endtime < sheetEvent.starttime) {
    Logger.log('end time must be after start time for event', sheetEvent, ridx);
    continue;
  }
}

// Ignore events outside of the begin/end range desired.
if (sheetEvent.starttime > endDate) {
  continue;
}
if (sheetEvent.endtime === '') {
  if (sheetEvent.starttime < beginDate) {
    continue;
  }
} else {
  if (sheetEvent.endtime < beginDate) {
    continue;
  }
}

// Determine if spreadsheet event is already in calendar and matches
var addEvent = true;
if (sheetEvent.id) {
  var eventIdx = calEventIds.indexOf(sheetEvent.id);
  if (eventIdx >= 0) {
    calEventIds[eventIdx] = null;  // Prevents removing event below
    addEvent = false;
    var calEvent = calEvents[eventIdx];
    if (!eventMatches(calEvent, sheetEvent)) {
      // Update the event
      updateEvent(calEvent, sheetEvent);

      // Maybe throttle updates.
      numChanges++;
      if (numChanges > THROTTLE_THRESHOLD) {
        Utilities.sleep(THROTTLE_SLEEP_TIME);
      }
    }
  }
}
if (addEvent) {
  var newEvent;
  sheetEvent.sendInvites = SEND_EMAIL_INVITES;
  if (sheetEvent.endtime === '') {
    newEvent = calendar.createAllDayEvent(sheetEvent.title, sheetEvent.starttime, sheetEvent);
  } else {
    newEvent = calendar.createEvent(sheetEvent.title, sheetEvent.starttime, sheetEvent.endtime, sheetEvent);
  }
  // Put event ID back into spreadsheet
  idData[ridx][0] = newEvent.getId();
  changesMade = true;

  // Set event color
  if (sheetEvent.color > 0 && sheetEvent.color < 12) {
    newEvent.setColor('' + sheetEvent.color);
  }

  // Maybe throttle updates.
  numChanges++;
  if (numChanges > THROTTLE_THRESHOLD) {
    Utilities.sleep(THROTTLE_SLEEP_TIME);
  }
}
}

// Save spreadsheet changes
if (changesMade) {
idRange.setValues(idData);
}

// Remove any calendar events not found in the spreadsheet
var numToRemove = calEventIds.reduce(function(prevVal, curVal) {
if (curVal !== null) {
  prevVal++;
}
return prevVal;
}, 0);
if (numToRemove > 0) {
//var ui = SpreadsheetApp.getUi();
//var response = ui.Button.YES;
if (numToRemove > numUpdated) {
//  response = ui.alert('Delete ' + numToRemove + ' calendar event(s) not found in spreadsheet?',
//      ui.ButtonSet.YES_NO);
//}
//if (response == ui.Button.YES) {
  calEventIds.forEach(function(id, idx) {
    if (id != null) {
      calEvents[idx].deleteEvent();
     Utilities.sleep(20);
    }
  });
}
}
Logger.log('Updated %s calendar events', numChanges);
syncFromCalendar();
syncFromCalendar2();
archive();
}

// Set up a trigger to automatically update the calendar when the spreadsheet is
// modified. See the instructions for how to use this.
function createSpreadsheetEditTrigger() {
var ss = SpreadsheetApp.getActive();
ScriptApp.newTrigger('syncToCalendar')
  .forSpreadsheet(ss)
  .onEdit()
  .create();
}

// Delete the trigger. Use this to stop automatically updating the calendar.
function deleteTrigger() {
// Loop over all triggers.
var allTriggers = ScriptApp.getProjectTriggers();
for (var idx = 0; idx < allTriggers.length; idx++) {
if (allTriggers[idx].getHandlerFunction() === 'syncToCalendar') {
  ScriptApp.deleteTrigger(allTriggers[idx]);
}
}
}

/**
* Moves row of data to another spreadsheet based on criteria in column 6 to sheet with same name as the value in column 4.
*/

function archive() {
// see Sheet event objects docs
// https://developers.google.com/apps-script/guides/triggers/events#google_sheets_events
var ss = SpreadsheetApp.openById("1LIgR-YYyEXfWdRZ7IAixpBroKAOCx76fHGnG3jfG8lw");
var s = ss.getSheetByName('bookings');
var r = s.getDataRange()

// to let you modify where the action and move columns are in the form responses sheet
var actionCol = 5;
var nameCol = 4;

// Get the number of columns in the active sheet.
// -1 to drop our action/status column
var colNumber = s.getLastColumn();
var rowNumber = s.getLastRow();
for (var w=1; w <= rowNumber; w++) {
for (var c=1; c <= colNumber; c++) {
// if our action/status col is changed to ok do stuff
if (c == actionCol && s.getRange(w, c).getValue().valueOf() < new Date().valueOf()) {
// get our target sheet name - in this example we are using the priority column
var targetSheet = "archive"
// if the sheet exists do more stuff
if (ss.getSheetByName(targetSheet)) {
  // set our target sheet and target range
  var targetSheet = ss.getSheetByName(targetSheet);
  var targetRange = targetSheet.getRange(targetSheet.getLastRow()+1, 1, 1, colNumber);
  // get our source range/row
  var sourceRange = s.getRange(w, 1, 1, colNumber);
  // new sheets says: 'Cannot cut from form data. Use copy instead.'
  sourceRange.copyTo(targetRange);
  // ..but we can still delete the row after
  s.deleteRow(w);
  // or you might want to keep but note move e.g. r.setValue("moved");
  Logger.log("moved row " + w);
}
}
}
}
var archiveSheet = ss.getSheetByName("archive")
var data = archiveSheet.getDataRange().getValues();
var newData = [];
for (var i in data) {
var row = data[i];
var duplicate = false;
for (var j in newData) {
  if(row[0] == newData[j][0] && row[1] == newData[j][1]){
    duplicate = true;
  }
}
if (!duplicate) {
  newData.push(row);
}
}
archiveSheet.clearContents();
archiveSheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);

}

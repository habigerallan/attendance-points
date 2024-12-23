function DATE_TO_KEY(dateObj) {
  var month = ('0' + (dateObj.getMonth() + 1)).slice(-2);
  var day = ('0' + dateObj.getDate()).slice(-2);
  var year = dateObj.getFullYear();
  return `${month}/${day}/${year}`;
}

var POINT_VALUES = {
  "late": -0.5,
  "missed": -1
};

var ATTENDANCE_SHEET_NAME = ""
var MAX_POINTS = ;

var API_URL = "";
var W2W_KEY = "";

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Attendance')
      .addItem('Update Data', 'update')
      .addToUi();
}

function update() {

  Logger.log("Update Started.");

  var headers = { "key": W2W_KEY };
  var options = { "headers": headers };

  var today = new Date();
  var oneYearAgo = new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());

  var urls = [];
  for (var currentStartDate = oneYearAgo; currentStartDate <= today; ) {
    
    // Creates all of the GET request URLs in 30 day increments

    var currentEndDate = new Date(currentStartDate);
    currentEndDate.setDate(currentStartDate.getDate() + 30);

    if (currentEndDate > today) { 

      currentEndDate = today;

    }

    var startDate = DATE_TO_KEY(currentStartDate);
    var endDate = DATE_TO_KEY(currentEndDate);

    urls.push(API_URL + "?start_date=" + encodeURIComponent(startDate) + "&end_date=" + encodeURIComponent(endDate));

    currentStartDate = new Date(currentEndDate);
    currentStartDate.setDate(currentStartDate.getDate() + 1);

  }

  try {

    // Attempting to recieve and process data

    var responses = UrlFetchApp.fetchAll(urls.map(url => ({ "url": url, ...options })));
    Logger.log("All data fetched successfully.");

    archive();
    Logger.log("Current data archived successfully.")

    const data = process(responses);
    Logger.log("All data processed successfully.");

    write(data);
    Logger.log("All data written successfully.");

  } catch (error) {

    Logger.log("Error fetching data: " + error.message + ".");

  }

  Logger.log("Update Ended.");

}

function archive() {

  Logger.log("Archiving Started.");

  /* Duplicates the current Attendance points sheet and
     archives it with the current timestamp as the name */

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME);
  var archive = sheet.copyTo(spreadsheet);

  const archiveName = "Archive: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd | HH:mm:ss");

  var sheetName = archiveName;

  archive.setName(sheetName)
  archive.hideSheet();

  Logger.log("Archiving Ended.");

}


function process(responses) {

  Logger.log("Processing Started.");

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME);

  var lastRow = sheet.getLastRow();
  var namesRange = sheet.getRange(2, 1, lastRow, 1);

  var lastChangedDate = new Date();
  lastChangedDate.setFullYear(lastChangedDate.getFullYear() - 1);


  // Localizing data

  var names = namesRange.getValues().flat();
  var points = Array(lastRow).fill(MAX_POINTS);
  var lastChanged = Array(lastRow).fill(DATE_TO_KEY(lastChangedDate));


  // Mapping names to their index for constant time lookup

  var nameToIndex = names.reduce((obj, name, index) => {

    obj[name] = index;
    return obj;

  }, {});

  var allShifts = [];
  responses.forEach((response, idx) => {

    try {

      // Attempts to localize all shifts from each API response

      var data = JSON.parse(response.getContentText());
      var shifts = data.AssignedShiftList.map(shift => ({

        START_DATE: shift.START_DATE, 
        FIRST_NAME: shift.FIRST_NAME, 
        LAST_NAME: shift.LAST_NAME, 
        CATEGORY_SHORT: shift.CATEGORY_SHORT

      }));

      allShifts = allShifts.concat(shifts);

    } catch (error) {

      Logger.log("Error processing response at index " + idx + ": " + error.message + ".");

    }

  });


  /* Maps all shifts on a specific day to a 
     dictionary with the date as the key */

  allShifts.sort((a, b) => new Date(a.START_DATE) - new Date(b.START_DATE));

  var shiftsByDay = {};
  allShifts.forEach(shift => {

    var dayKey = DATE_TO_KEY(new Date(shift.START_DATE));
    if (!shiftsByDay[dayKey]) {

      shiftsByDay[dayKey] = [];

    }
    
    shiftsByDay[dayKey].push(shift);

  });

  var today = new Date();
  var oneYearAgo = new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());

  for (var d = new Date(oneYearAgo.getTime()); d <= today; d.setDate(d.getDate() + 1)) {

    /* Modifies localized attendance points based upon
       shift data for all days within the year */

    var currentDay = new Date(d);
    var currentDayKey = DATE_TO_KEY(currentDay);

    var dayShifts = shiftsByDay[currentDayKey] || [];
    var processedIndices = new Set();
    
    Logger.log("Processing Day: " + currentDayKey + ".");

    dayShifts.forEach(shift => {

      var fullName = `${shift.FIRST_NAME} ${shift.LAST_NAME}`.trim();
      var i = nameToIndex[fullName];

      if (i !== undefined) {

        processedIndices.add(i);

        var category = shift.CATEGORY_SHORT;
        if (Object.prototype.hasOwnProperty.call(POINT_VALUES, category)) {

          points[i] += POINT_VALUES[category];
          lastChanged[i] = currentDayKey;

          Logger.log("Shift Processed | Name: " + names[i] + " | Points: " + points[i] + " | Last Changed: " + lastChanged[i] + " | Category: " + category + ".");

        }

      }

    });

    for (var i = 0; i < names.length; i++) {

      // Checks for perfect attendance for all employees each day

      if (processedIndices.has(i)) continue;

      var lastChange = new Date(lastChanged[i]);
      var daysDifference = (currentDay - lastChange) / (1000 * 60 * 60 * 24);

      if (daysDifference >= 30) {

        points[i] = Math.min(points[i] + 1, 9);
        lastChanged[i] = currentDayKey;

        Logger.log("Point Rewarded | Name: " + names[i] + " | Points: " + points[i] + " | Last Changed: " + lastChanged[i] + ".");
        
      }

    }

  }

  const data = names.map((name, index) => ({

    name: name,
    points: points[index],
    lastChanged: lastChanged[index]

  }));

  Logger.log("Processing Completed.");

  return data;
}

function write(data) {

  Logger.log("Writing Started.");

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME);

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nameCol = headers.indexOf("Name") + 1;
  var pointsCol = headers.indexOf("Current Points") + 1;
  var lastChangedCol = headers.indexOf("Last Changed") + 1;
  var lastRanCol = headers.indexOf("Last Ran") + 1;

  var today = new Date();
  var formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd/yyyy");


  // Writing data into row format

  var rows = data
    .filter(item => item.name)
    .map(function(item) {

      var row = [];
      row[nameCol - 1] = item.name;
      row[pointsCol - 1] = item.points;
      row[lastChangedCol - 1] = item.lastChanged;
      row[lastRanCol - 1] = formattedToday;
      return row;
      
  });

  var startRow = 2;
  for (var i = 0; i < rows.length; i++) {

    // Writes data into sheet

    if (rows[i][nameCol - 1]) {

      sheet.getRange(startRow + i, nameCol, 1, 1).setValue(rows[i][nameCol - 1]);

    }

    if (rows[i][pointsCol - 1]) {

      sheet.getRange(startRow + i, pointsCol, 1, 1).setValue(rows[i][pointsCol - 1]);

    }

    if (rows[i][lastChangedCol - 1]) {

      sheet.getRange(startRow + i, lastChangedCol, 1, 1).setValue(rows[i][lastChangedCol - 1]);

    }

    sheet.getRange(startRow + i, lastRanCol, 1, 1).setValue(formattedToday);

  }

  Logger.log("Writing Ended.");

}

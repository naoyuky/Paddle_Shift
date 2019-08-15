var startColumn = 3; // Position of Shift Name
//Define Relative position for each object from Col C
var shiftName = 0;
var startHr = 1;
var startMin = 2;
var endHr = 3;
var endMin = 4;
var isUpdate = 5;

var startRow = 2; // 1st row is for the title.

var calID = ""; //please add your calendar ID here.

function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheetUp = spreadsheet.getSheetByName('Schedule');
  var sheetCompare = spreadsheet.getSheetByName('Recent');
    
  var shift = sheetUp.getRange(startRow, startColumn, 31, 6).getValues(); //get until update column

  var monthRange = sheetUp.getRange(2, 9, 1, 2).getValues();
  var month = monthRange[[0]][0]-1;
  var year = monthRange[[0]][1];
  
  var recentMonth = sheetCompare.getRange(2, 9, 1, 2).getValues();
  
  var lastDay = sheetUp.getRange(2, 11).getValue();
  
  var cal = CalendarApp.getCalendarById(calID);
  
  if(monthRange[[0]][0] !== recentMonth[[0]][0] || monthRange[[0]][1] !== recentMonth[[0]][1]){
    // if new years or month, update schedules for entire month.

    for(var i = 0; i < lastDay; i++) {
      addCal(shift[i], year, month, i+1, cal);
    }
  } else if (monthRange[[0]][0] === recentMonth[[0]][0] && monthRange[[0]][1] === recentMonth[[0]][1]){
      // if the same years and months, update the schedule with a checkbox enabled.
    for(var i = 0; i < lastDay; i++){
      if(shift[i][isUpdate]){
        addCal(shift[i], year, month, i+1, cal);
      }
    }
  }
  SpreadsheetApp.getActive().deleteSheet(sheetCompare);
  sheetUp.activate();
  SpreadsheetApp.getActive().duplicateActiveSheet().setName("Recent");
  
  Browser.msgBox("completed adding schedule.");
}

// add the schedule to the calendar.
function addCal(shiftArray, year, month, day, calendar){
  if(shiftArray[startHr]!=shiftArray[endHr] && (shiftArray[startHr] != "" && shiftArray[endHr] != "")){
    // if no shift is added to the array, no schedule will be added.

    var title = "Shift" + shiftArray[shiftName];
    var shiftStart, shiftEnd;
    
    if(shiftArray[startHr] >= 24){
      shiftStart = new Date(year, month, day+1, shiftArray[startHr] - 24, shiftArray[startMin]);
    }else{
      shiftStart = new Date(year, month, day, shiftArray[startHr], shiftArray[startMin]);
    }
    if(shiftArray[endHr] >= 24){
      shiftEnd = new Date(year, month, day+1, shiftArray[endHr] - 24, shiftArray[endMin]);
    }else{
      shiftEnd = new Date(year, month, day, shiftArray[endMin], shiftArray[endMin]);
    }
    var options = {
      description : "",
      location : "",
    };
    calendar.createEvent(title, shiftStart, shiftEnd, options);
  }
}
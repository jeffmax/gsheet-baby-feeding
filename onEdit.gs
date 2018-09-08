var EDIT_ROW = 2;
var PREV_ROW = 3;
var BREAST_COL = 2;
var FORMULA_COL = 3;
var START_COL = 1;
var DONE_COL = 5;
var TOTAL_COL = 6;
var NOTES_COL = 7;
var DAY_COL = 8;
var END_COL = 12;
var DAY_COL_LETTER='H';
var DISABLED_ROW_COLOR = "#D3D3D3";



// From https://stackoverflow.com/questions/11301438/return-index-of-greatest-value-in-an-array
function indexOfMin(arr) {
    if (arr.length === 0) {
        return -1;
    }

    var min = arr[0];
    var minIndex = 0;

    for (var i = 1; i < arr.length; i++) {
        if (arr[i] < min) {
            minIndex = i;
            min = arr[i];
        }
    }

    return minIndex;
}

function rowsMatchingDayFromRow(matchingDay, row, babySheet){
   row = row+1;
   var nextDate = new Date(matchingDay.toLocaleDateString());
   nextDate.setDate(nextDate.getDate()+1);
   var dates = babySheet.getRange(DAY_COL_LETTER+row+":"+DAY_COL_LETTER).getValues();
   var rows = [];
   for (var i = 0; i < dates.length; i++){
      var iDate = dates[i][0];
      if (iDate < matchingDay){
         break;
      }
      if (iDate < nextDate){
        rows.push(i+row);
      }
   }
   return rows;
}

function pad(n){return (n<10?'0':'')+n};

function dateToTimeString(d){
  var hours = d.getHours();
  return (pad(hours%12 || 12) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds())+ ' ' + (hours<12? 'AM' :'PM'));
}

function onEdit(e) {
  var sheetName = e.source.getActiveSheet().getName();
 
  if (sheetName != "Log"){
    return;
  }
  // Only one execution of this function at once
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  var inserted_time = false;
 
  var d = new Date();
  var lastDate;
  
  var instructions = "ENTER -->";
  var currentCell = e.source.getActiveSelection();
  var currentCellCol = currentCell.getColumn();
  var currentCellRow = currentCell.getRow();
  var babySheet = e.source.getActiveSheet();
  var startTime = babySheet.getRange(currentCellRow, START_COL);
  var endDateTime = babySheet.getRange(currentCellRow, END_COL);
  var runningTotal = babySheet.getRange(currentCellRow, TOTAL_COL);
  var rowDay = babySheet.getRange(currentCellRow, DAY_COL);
  
  // Set the formula that calculates the running total, this will work when someone manually inserts row
  if (runningTotal.getValue()==""){
    runningTotal.setValue("=SUMIF(INDIRECT(\""+DAY_COL_LETTER+"\"&ROW()&\":"+DAY_COL_LETTER+"\"),INDIRECT(\""+DAY_COL_LETTER+"\"&ROW()),INDIRECT(\"INTERNAL!A\"&ROW()&\":A\"))")
  }
 
  // Populate the date (handles editing the top row and when someone manually inserts a row)
  var day = babySheet.getRange(currentCellRow, DAY_COL);
  if (day.getValue() == ""){
      if (currentCellRow==EDIT_ROW){
           day.setValue(d.toLocaleDateString());
      }else{
        // Use the date of the cell below. 
        var belowDate = babySheet.getRange(currentCellRow+1, DAY_COL).getValue();
        babySheet.getRange(currentCellRow, DAY_COL).setValue(belowDate.toLocaleDateString());
      }
  }   
  
  // Did they just edit the top row for the first time? 
  // If a row has been inserted manually, rangeWidth will be greater than one
  if (currentCellRow==EDIT_ROW && e.range.getWidth() == 1 & (startTime.getDisplayValue() == "" || startTime.getDisplayValue() == instructions)){
      startTime.setValue(dateToTimeString(d));
      inserted_time = true;
      lastDate = babySheet.getRange(PREV_ROW, DAY_COL).getValue();
      if (lastDate.toLocaleDateString() != d.toLocaleDateString()){
        // Add a border to the top of the last row to show that it was yeterday
        var yesterdayRow = babySheet.getRange(PREV_ROW,1,1, babySheet.getLastColumn());
        yesterdayRow.setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
  }else if (currentCellRow==EDIT_ROW && e.range.getWidth() > 1 & startTime.getValue() == ""){
      // If a row was inserted manually at the top, set instructions, eitherwise leave it blank
      startTime.setValue(instructions);
  }
  
  // Go through all prior days feeding times and compare edited times time
  if (inserted_time || (currentCellCol == START_COL && e.oldValue != e.value)){ 
    var priorDay = new Date(startTime.getDisplayValue() + " " + rowDay.getDisplayValue());
    // If there is actually a date in here, continue
    if (priorDay instanceof Date && !isNaN(priorDay)){
        priorDay.setDate(priorDay.getDate()-1);  
        var priorDaysRows = rowsMatchingDayFromRow(new Date(priorDay.toLocaleDateString()), currentCellRow, babySheet);
        if (priorDaysRows.length > 0){
            var times = babySheet.getRange("A"+priorDaysRows[0]+":A"+priorDaysRows[priorDaysRows.length-1]).getValues();
            var offsets=[];
            for (var i=0; i<times.length;i++){
               var iDate = new Date(dateToTimeString(times[i][0]) + " " + priorDay.toLocaleDateString());
               offsets.push(Math.abs(priorDay.getTime() - iDate.getTime()));
            }
            // Find the closest time yesterday to the current feeding time
            var index = indexOfMin(offsets);
            if (index != -1){
               if (offsets[index] < 7200000){ // If not within two hours, don't bother 
                   var closestTotalYesterday = babySheet.getRange(priorDaysRows[index], TOTAL_COL).getDisplayValue();
                   var closestTimeYesterday = dateToTimeString(times[index][0]);
                   // Set the note
                   babySheet.getRange(currentCellRow, TOTAL_COL).setNote("At the closest time yesterday ("+closestTimeYesterday+"), the total feed was "+ closestTotalYesterday + " oz");
               }else{
                  // No valid match was found, if prior note there, clear it
                  babySheet.getRange(currentCellRow, TOTAL_COL).clearNote();
               }
            }else{
                // No dates were found, if note exists, clear it
                babySheet.getRange(currentCellRow, TOTAL_COL).clearNote();
            }
         }
    }else{
          // Date is invalid, clear out any existing note
          babySheet.getRange(currentCellRow, TOTAL_COL).clearNote();
    }
    
  }
  
  // Did they just hit done on the top row?
  var done = babySheet.getRange(currentCellRow, DONE_COL);
  var breast = babySheet.getRange(currentCellRow, BREAST_COL).getValue();
  var formula = babySheet.getRange(currentCellRow, FORMULA_COL).getValue();
  var timestamp = babySheet.getRange(currentCellRow, START_COL).getValue();
  
  if (currentCellRow==EDIT_ROW && done.getValue() == true && (breast != "" || formula != "") && timestamp != "" && timestamp != instructions){
    endDateTime.setValue(d.toLocaleString());
    babySheet.getRange(EDIT_ROW, TOTAL_COL).setFontWeight("bold");
    lastDate = babySheet.getRange(PREV_ROW, DAY_COL).getValue();
    if (lastDate.toLocaleDateString() != d.toLocaleDateString()){
      babySheet.getRange(PREV_ROW, TOTAL_COL).setFontWeight("bold");
    }else{
      babySheet.getRange(PREV_ROW, TOTAL_COL).setFontWeight("normal");
    }
    babySheet.getRange(EDIT_ROW,1,1, babySheet.getLastColumn()).setBackground(DISABLED_ROW_COLOR);
    babySheet.insertRowBefore(EDIT_ROW);
    babySheet.getRange(EDIT_ROW, START_COL).setValue(instructions);
    babySheet.getRange(EDIT_ROW,1,1,babySheet.getLastColumn()).setBackground("white");
    babySheet.getRange(EDIT_ROW, TOTAL_COL).setFontWeight("normal")
  }else if (currentCellRow==EDIT_ROW && done.getValue() == true){
    done.setValue(false);
  }
  SpreadsheetApp.flush();
  lock.releaseLock();
}

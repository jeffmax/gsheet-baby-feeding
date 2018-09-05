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
  
  var EDIT_ROW = 2;
  var PREV_ROW = 3;
  var BREAST_COL = 2;
  var FORMULA_COL = 3;
  var START_COL = 1;
  var DONE_COL = 4;
  var TOTAL_COL = 5;
  var NOTES_COL = 6;
  var DAY_COL = 7;
  var END_COL = 10;
  var ROW_TOTAL_COL=11;
  var DISABLED_ROW_COLOR = "#D3D3D3";
  var d = new Date();
  var lastDate;
  
  var instructions = "ENTER -->";
  var currentCell = e.source.getActiveSelection();
  var currentCellRow = currentCell.getRow();
  var babySheet = e.source.getActiveSheet();
  var startTime = babySheet.getRange(currentCellRow, START_COL);
  var endDateTime = babySheet.getRange(currentCellRow, END_COL);
  var runningTotal = babySheet.getRange(currentCellRow, TOTAL_COL);
  var rowTotal = babySheet.getRange(currentCellRow, ROW_TOTAL_COL);
  // Set the formula that calculates the running total, this will work when someone manually inserts row
  if (runningTotal.getValue()==""){
    runningTotal.setValue("=SUMIF(INDIRECT(\"G\"&ROW()&\":G\"),INDIRECT(\"G\"&ROW()),INDIRECT(\"INTERNAL!A\"&ROW()&\":A\"))")
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
  if (currentCellRow==EDIT_ROW && e.range.getWidth() == 1 & (startTime.getValue() == "" || startTime.getValue() == instructions)){
      startTime.setValue(dateToTimeString(d));
      lastDate = babySheet.getRange(PREV_ROW, DAY_COL).getValue();
      if (lastDate.toLocaleDateString() != d.toLocaleDateString()){
        // Add a border to the top of the last row to show that it was yeterday
        var yesterdayRow = babySheet.getRange(PREV_ROW,1,1, babySheet.getLastColumn());
        yesterdayRow.setBorder(true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
      // Set the note on this row's total to show what the closest feeding to yesterday's total was.
      // Create date objects with midnigh time
      var today = new Date(d.toLocaleDateString());
      var yesterday = new Date(d.toLocaleDateString());
      yesterday.setDate(d.getDate() - 1);
      // Locate all rows for yesterday's date
      var dates = babySheet.getRange("G2:G").getValues();
      var yesterdayRows = [];
      for (var i = 0; i < dates.length; i++){
         var iDate = dates[i][0];
         if (iDate < yesterday){
            break;
         }
         if (iDate < today){
          yesterdayRows.push(i+2);
         }
      }
      if (yesterdayRows.length > 0){
        // Go through all yesterday's feeding times and compare to today's time
        var thisTimeYesterday = new Date(d);
        thisTimeYesterday.setDate(d.getDate()-1);
        var times = babySheet.getRange("A"+yesterdayRows[0]+":A"+yesterdayRows[yesterdayRows.length-1]).getValues();
        var offsets=[];
        for (var i=0; i<times.length;i++){
           var iDate = new Date(dateToTimeString(times[i][0]) +" " + yesterday.toLocaleDateString());
           offsets.push(Math.abs(thisTimeYesterday.getTime() - iDate.getTime()));
        }
        
        // Find the closest time yesterday to the current feeding time
        var index = indexOfMin(offsets);
        if (index != -1){
          if (offsets[index] < 7200000){ // If not within two hours, don't bother 
            var closestTotalYesterday = babySheet.getRange(yesterdayRows[index], TOTAL_COL).getDisplayValue();
            var closestTimeYesterday = dateToTimeString(times[index][0]);
            // Set the note
            babySheet.getRange(EDIT_ROW, TOTAL_COL).setNote("At the closest time yesterday ("+closestTimeYesterday+"), the total feed was "+ closestTotalYesterday + " oz");
          }
        }
      }
  }else if (currentCellRow==EDIT_ROW && e.range.getWidth() > 1 & startTime.getValue() == ""){
      // If a row was inserted manually at the top, set instructions, eitherwise leave it blank
      startTime.setValue(instructions);
  }
  
  // Did they just hit done on the top row?
  var done = babySheet.getRange(currentCellRow, DONE_COL);
  var breast = babySheet.getRange(currentCellRow, BREAST_COL).getValue();
  var formula = babySheet.getRange(currentCellRow, FORMULA_COL).getValue();
  var timestamp = babySheet.getRange(currentCellRow, START_COL).getValue();
  
  if (currentCellRow==EDIT_ROW && done.getValue() == true && (breast != "" || formula != "") && timestamp != "" && timestamp != instructions){
    endDateTime.setValue(d.toLocaleString());
    babySheet.getRange(EDIT_ROW, TOTAL_COL).setFontWeight("bold");
    //babySheet.getRange(EDIT_ROW, TOTAL_COL).clearNote();
    lastDate = babySheet.getRange(PREV_ROW, DAY_COL).getValue();
    if (lastDate.toLocaleDateString() != d.toLocaleDateString()){
      babySheet.getRange(PREV_ROW, TOTAL_COL).setFontWeight("bold");
    }else{
      babySheet.getRange(PREV_ROW, TOTAL_COL).setFontWeight("normal");
    }
    babySheet.getRange(EDIT_ROW,1,1, babySheet.getLastColumn()).setBackground(DISABLED_ROW_COLOR);
    babySheet.insertRowBefore(EDIT_ROW);
    babySheet.getRange(EDIT_ROW, START_COL).setValue(instructions);
    babySheet.getRange(EDIT_ROW,1,1, babySheet.getLastColumn()).setBackground("white");
    babySheet.getRange(EDIT_ROW, TOTAL_COL).setFontWeight("normal")
  }else if (currentCellRow==EDIT_ROW && done.getValue() == true){
    done.setValue(false);
  }
  SpreadsheetApp.flush();
  lock.releaseLock();
}

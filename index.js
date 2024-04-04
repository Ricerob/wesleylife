function getNextEmptyRow(sheet) {
  if (!sheet) {
    throw new Error("Sheet is undefined. Make sure the sheet exists and is properly referenced.");
  }

  var range = sheet.getRange("B:B");
  var values = range.getValues();
  for (var i = 2; i < values.length; i++) {
    if (!values[i][0]) {
      return i + 1; 
    }
  }
  return values.length + 1; // If no empty row found, return the next row number
}


function getMostRecentWeek(sheet) {
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange("A4:A" + lastRow).getValues(); // Get all values in column A
  
  // Iterate backwards to find the last row with a date value
  for (var i = values.length - 1; i >= 0; i--) {
    var value = values[i][0];
    if (value instanceof Date) {
      return value; // Return the date object if found
    }
  }
  
  // If no date value is found, return null or handle the scenario as required
  return null;
}

function getMostRecentHarvest(sheet, ui) {
  var values = sheet.getRange("A4:A").getValues();

  for (var i = values.length - 1; i >= 0; i--) {
    var value = values[i][0];
    if (value) {
      return parseInt(value);
    }
  }

  return null;
}

function transferSeedingToTransplant() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var row = SpreadsheetApp.getActiveRange().getRow();
  var currentDate = new Date();


  let week = sheet.getRange(row, 1).getValue();

  if(!week) {
    // In the future, handle by finding row before this with a valid week. For now, return
    return null;
  }

  let sPeople = sheet.getRange(row, 4).getValue();
  let color = sheet.getRange(row, 5).getValue();

  var transplantSheet = SpreadsheetApp.getActive().getSheetByName('Transplant')
  if (transplantSheet === null) {
    return null;
  }


  // Add values based on if new week, new day, etc.

  var transplantRecentWeek = getMostRecentWeek(transplantSheet);
  var weekAfter = new Date(transplantRecentWeek);
  var MMDD = (currentDate.getMonth() + 1) + '/' + currentDate.getDate();
  var MMDDYYYY = (currentDate.getMonth() + 1) + '/' + currentDate.getDate() + '/' + currentDate.getFullYear();
  var transplantNextEmptyRow = getNextEmptyRow(transplantSheet)

  // Increment weekAfter by 7 days
  weekAfter.setDate(weekAfter.getDate() + 7);
  let newWeek = false;
  if (currentDate > weekAfter) {
    newWeek = true;
  }

  if(newWeek || transplantNextEmptyRow === 4) {
    transplantSheet.getRange(transplantNextEmptyRow, 1).setValue(MMDDYYYY)
  }

  transplantSheet.getRange(transplantNextEmptyRow, 2).setValue(MMDD)
  transplantSheet.getRange(transplantNextEmptyRow, 3).setValue(week)
  transplantSheet.getRange(transplantNextEmptyRow, 4).setValue(color)
  transplantSheet.getRange(transplantNextEmptyRow, 6).setValue(sPeople);

  return null;
}

function transferTransplantToHarvest() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[1];
  var row = SpreadsheetApp.getActiveRange().getRow();
  var currentDate = new Date();
  var MMDDYYYY = (currentDate.getMonth() + 1) + '/' + currentDate.getDate() + '/' + currentDate.getFullYear();

  let transplantRow = sheet.getRange(row, 7).getValue();
  let transplantPeople = sheet.getRange(row, 5).getValue();
  let seedingPeople = sheet.getRange(row, 6).getValue();
  let transplantColor = sheet.getRange(row, 4).getValue();

  // Grab previous color. If same, add under same harvest number. Else, new

  var harvestSheet = SpreadsheetApp.getActive().getSheetByName('Harvest')
  if (harvestSheet === null) {
    return null;
  }

  var harvestPrevRow = harvestSheet.getLastRow();
  var lastColor = harvestSheet.getRange(harvestPrevRow, 7).getValue();
  var newHarvest = false;
  if(transplantColor !== lastColor) {
    newHarvest = true;
  }

  var nextHarvestNum;

  if(newHarvest) {

    let lastHarvestNum = getMostRecentHarvest(harvestSheet, ui)

    if(!lastHarvestNum) {
      nextHarvestNum = 1
    } else {
      nextHarvestNum = parseInt(lastHarvestNum) + 1;
    }

    nextHarvestNum = nextHarvestNum.toString().padStart(3, '0')
    harvestSheet.getRange(harvestPrevRow + 1, 1).setValue(nextHarvestNum)
  }

  // Set values
  harvestSheet.getRange(harvestPrevRow + 1, 2).setValue(MMDDYYYY);
  harvestSheet.getRange(harvestPrevRow + 1, 4).setValue(transplantRow);
  harvestSheet.getRange(harvestPrevRow + 1, 5).setValue(transplantPeople);
  harvestSheet.getRange(harvestPrevRow + 1, 6).setValue(seedingPeople);
  harvestSheet.getRange(harvestPrevRow + 1, 7).setValue(transplantColor)

  return null;
}

function transferHarvestToDistribution() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var distroSheet = ss.getSheets()[3];
  var row = SpreadsheetApp.getActiveRange().getRow();
  var currentDate = new Date();

  var harvestNum = sheet.getRange(row, 1).getValue();
  var harvestDate = sheet.getRange(row, 2).getValue();

  var result = ui.prompt("How many lbs are being distributed?", ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var lbs;
  if (button == ui.Button.OK) { 
    lbs = result.getResponseText();
  } else {
    return
  }

  var result1 = ui.prompt("Scan customer num:", ui.ButtonSet.OK_CANCEL);

  var button = result1.getSelectedButton();
  var customerNum;
  if (button == ui.Button.OK) { 
    customerNum = result1.getResponseText();
  } else {
    return
  }

  var data = distroSheet.getRange('K4:L').getValues(); // Get values of columns K and L
  var found = false;
  var customerName = '';

  // Loop through each row in the data
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == customerNum) {
      customerName = data[i][1];
      found = true;
    }
  }

  if(!found) {
    ui.alert("Couldn't find customer mapping.")
    return;
  }

  // Set values
  var lastRow = getNextEmptyRow(distroSheet)
  distroSheet.getRange(lastRow, 1).setValue(harvestNum);
  distroSheet.getRange(lastRow, 2).setValue(harvestDate);
  distroSheet.getRange(lastRow, 4).setValue(lbs);
  distroSheet.getRange(lastRow, 5).setValue(customerName);
  distroSheet.getRange(lastRow, 6).setValue(currentDate);
  
}

// Used to detect if a change on cell B1 happened
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var editedColumn = range.getColumn();
  var editedRow = range.getRow()
  var userEmail = Session.getActiveUser().getEmail();
  var currentDate = new Date();
  var ui = SpreadsheetApp.getUi();
  
  if (editedColumn == 2 && editedRow == 1) { 
    var scannedValue = e.value;
    
    // New seed row
    if (scannedValue === "barcode1") {
      var nextWeek = getMostRecentWeek(sheet);
      var weekAfter = new Date(nextWeek);

      // Increment weekAfter by 7 days
      weekAfter.setDate(weekAfter.getDate() + 7);

      let newWeek = false;
      if (currentDate > weekAfter) {
        newWeek = true;
      }

      const nextEmptyRow = getNextEmptyRow(sheet);
      var MMDD = (currentDate.getMonth() + 1) + '/' + currentDate.getDate();
      var MMDDYYYY = (currentDate.getMonth() + 1) + '/' + currentDate.getDate() + '/' + currentDate.getFullYear();

      if(!newWeek) {
        // Set new MMDD of row
        sheet.getRange(nextEmptyRow, 2).setValue(MMDD)

        // Set color value
        const color = sheet.getRange(nextEmptyRow - 1, 5).getValue();
        sheet.getRange(nextEmptyRow, 5).setValue(color)
      }
      else {
        // Set week & day
        sheet.getRange(nextEmptyRow, 1).setValue(MMDDYYYY);
        sheet.getRange(nextEmptyRow, 2).setValue(MMDD);
      }
      
    }
  }


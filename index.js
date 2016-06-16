// Mover Orders though a simple process of 'To Be Ordered' -> 'On Order' -> 'Received'
// Each Department can have it's own flow so they know the status of thier own books
//
// To move data though the flow sheets must be named in the following format:
//   * 'To Be Ordered - YOUR_NAME_HERE'
//   * 'On Order - YOUR_NAME_HERE'
//   * 'Received'
//
// To move an order each sheet expects a specific status to be entered in 'Order Status' feild
// Entering the correct status causes the order to be moved to the next step
//   * To Be Ordered -> Ordered
//   * On Order -> Received

function onOpen() {
 var ss = SpreadsheetApp.getActiveSpreadsheet(),
     options = [
      {name:"Process Changes", functionName:"approveRequests"},
      {name:"Toggle Active Sheet Lock", functionName:"toggleLockSheet"}
     ];
 ss.addMenu("Library Actions", options);
}

function approveRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Logs') || ss.insertSheet('Logs');
  const sheets = ss.getSheets();
  const startTime = new Date();
  //var protections = [];
  
  const LastRowInLogSheet = logSheet.getLastRow();
  const LastColumnInLogSheet = logSheet.getLastColumn();
  
  if(LastRowInLogSheet && LastColumnInLogSheet) {
    const LastLogSheetValue = logSheet.getRange(LastRowInLogSheet, LastColumnInLogSheet).getValue();
    
    if (LastLogSheetValue !== 'SCRIPT FINISHED') {
      const ui = SpreadsheetApp.getUi(); // Same variations.
      
      const result = ui.alert(
        'STOP',
        'You are trying to run the script more then once. \r\n Please check the Log tab to see when the script is finished. \r\n\r\n If you are 100% sure the script is not runing delete the Logs tab and run script again.',
        ui.ButtonSet.OK);
      return;
    }
  }
  
  logSheet.clear();
  logSheet.appendRow(['STARTING SCRIPT ' + startTime]);
  
/*
  logSheet.appendRow(['Locking all sheets in spreadsheet']);
  
  sheets.forEach(function(sheet, index) {
    protections[index] = sheet.protect().setDescription('Library Actions Script Running');
    protections[index].removeEditors(protections[index].getEditors());
    
    logSheet.appendRow(['Sheet "'+ sheet.getName() +'" restricted to ' + protections[index].getEditors()]);
    
    if (protections[index].canDomainEdit()) {
      protections[index].setDomainEdit(false);
    }
  });
  
  logSheet.appendRow(['All Sheets are locked']);
 */
  
  // Sheet Types
  const TO_BE_ORDERED = 'To Be Ordered';
  const ON_ORDER = 'On Order';
  const RECEIVED = 'Received';
  
  const SHEET_TYPES = [TO_BE_ORDERED, ON_ORDER, RECEIVED];
  const ACTIONABLE_ORDER_STATUS = ['ordered','received'];
  
  sheets.forEach(function(sheet, index) {
    const sheetName = sheet.getName();
    const sheetNameValues = sheetName.split(' - ');
    
    if(sheetNameValues.length !== 2 || SHEET_TYPES.indexOf(sheetNameValues[0]) === -1) {
      Logger.log(sheetName + ' will be skipped.');
      //logSheet.appendRow([sheetName + ' will be skipped.']);
      return;
    }
    
    logSheet.appendRow(['Looking for books to move on: ' + sheetName]);
    
    const sheetType = sheetNameValues[0];
    const sheetTypeIndex = SHEET_TYPES.indexOf(sheetType);
    const sheetCategory = sheetNameValues[1];
    const sheetOrderStatus = ACTIONABLE_ORDER_STATUS.find(function(status, index) { return index === sheetTypeIndex });
    
    const nextSheetType = SHEET_TYPES[sheetTypeIndex + 1];
    const nextSheetName = nextSheetType === 'Received' ? nextSheetType : nextSheetType + ' - ' + sheetCategory;
    const nextSheet = ss.getSheetByName(nextSheetName);
    
    if(!nextSheet) {
      Logger.log('ERROR: Next sheet name is not found: ' + sheetName);
      logSheet.appendRow(['ERROR: Next sheet name is not found: ' + sheetName]);
      return;
    }
    
    const range = sheet.getDataRange();
    const rawValues = range.getValues();
    const headerValues =  rawValues[0];
    const rowValues =  rawValues.slice(1,rawValues.length);
    const ORDER_STATUS_INDEX = headerValues.indexOf('Order Status');
    
    if(ORDER_STATUS_INDEX === -1) {
      Logger.log('ERROR: Order Status column does not exist: ' + sheetName);
      logSheet.appendRow(['ERROR: Order Status column does not exist: ' + sheetName]);
      return;
    }
    
    // Move correct rows to next sheet
    var rowsToDelete = [];
    const orderedRows = rowValues.filter(function(row, index) {
      const orderStatus = row[ORDER_STATUS_INDEX].toLowerCase();
      if(row[ORDER_STATUS_INDEX].toLowerCase() == sheetOrderStatus) {
        // + 2 because our index starts at 0 and the first row
        // of headers is removed from this array
        rowsToDelete.push(index + 2);
        nextSheet.appendRow(row);
        
        logSheet.appendRow(['Moved "' + row[1] + '" to sheet: ' + nextSheetName]);
      }
    });
    
    logSheet.appendRow(['Found ' + rowsToDelete.length + ' books to move on: ' + sheetName]);
    
    // clean up old rows from highest index to lowest
    // otherwise the index of the row will change
    rowsToDelete.reverse().forEach(function(rowIndex) {
      sheet.deleteRow(rowIndex);
    });
    
  });
  
  /*
  protections.forEach(function(protection, index) {
      protection.remove();
  });
  logSheet.appendRow(['All Sheets are unlocked']);
  */
  
  const endTime = new Date();
  logSheet.appendRow(['SCRIPT TOOK: ' + ((endTime - startTime) / 1000) + ' SECONDS']);
  logSheet.appendRow(['SCRIPT FINISHED']);
  logSheet.autoResizeColumn(1);
}



function toggleLockSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [ss.getActiveSheet()];
  var protections = [];
  
  sheets.forEach(function(sheet, index) {
    var p = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET) || [];
    
    p.forEach(function(v) {
      protections.push(v);
    });
  });
  
  if(protections.length) {
    protections.forEach(function(protection, index) {
      protection.remove();
    });
  } else {
  
    sheets.forEach(function(sheet, index) {
      protections[index] = sheet.protect().setDescription('LOCK TOGGLED ON');
      protections[index].removeEditors(protections[index].getEditors());
      
      if (protections[index].canDomainEdit()) {
        protections[index].setDomainEdit(false);
      }
    });
  
  }

}

//
// Pollyfill useful functions
//

if (!Array.prototype.find) {
  Array.prototype.find = function(predicate) {
    if (this === null) {
      throw new TypeError('Array.prototype.find called on null or undefined');
    }
    if (typeof predicate !== 'function') {
      throw new TypeError('predicate must be a function');
    }
    var list = Object(this);
    var length = list.length >>> 0;
    var thisArg = arguments[1];
    var value;

    for (var i = 0; i < length; i++) {
      value = list[i];
      if (predicate.call(thisArg, value, i, list)) {
        return value;
      }
    }
    return undefined;
  };
}

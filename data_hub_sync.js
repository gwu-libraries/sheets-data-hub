
function loadParams(localss) {
  
  /*
  This function loads metadata about the local (currently open) spreadsheet from a sheet (tab) called "params." The metadata describe, for each distinct sheet (tab/table) on the local spreadsheet,
  the following:
    + the name of each local sheet/table/tab (sheetName, string)
    + the corresponding name of the sheet/table/tab in the data hub (remoteSheetName, string)
    + the name of the identifier column, which should contain a unique numeric ID for each row (primaryKey, string)
    + the names of the columns "owned" by the local spreadsheet (localColumns, string)
    + the names of the columns not "owned" by the local spreadsheet (foreignColumns, string),
    + whether the local spreadsheet can add new rows to the data hub (canAddKeys, Boolean)
    
    Returns keyMap = {sheetName = { remoteSheetName: ['name of remote data source']::string, 
                                    primaryKey: ['column name of primary key column']::string,
                                    localColumns: ['column name owned locally']::[Array of string],
                                    foreignColumns: ['column name owned remotely']::[Array of string],
                                    canAddKey: boolean } 
                      }                  
                                   
  */
  
  // should correspond to column headers on the 'params' sheet/tab/table
  
  var paramHeaders = ['remoteSheetName', 'primaryKey','localColumns','foreignColumns', 'canAddKeys']; 
    
  
  // load the sheet (table) name(s) and relevant columns from this spreadsheet 
    var paramSheet = localss.getSheetByName('params'),
       paramData = paramSheet.getDataRange().getValues();
  
   
    var keyMap = paramData.slice(1).reduce(function (obj, row) {
      //first cell of each row holds the name of a table to be updated
      // we make an object where the table name is the key pointing to an object that holds parameter values
      // e.g., {'script_test': {'remoteSheetName': 'individual_subs', 'primaryKey': 'po_number', 'localColumns': ['bib_id', 'amount']}}
      
      // first time seeing this sheetName?
      if (!(row[0] in obj)) { 
        
        // create the inner obj
        obj[row[0]] = row.slice(1).reduce(function (innerObj, cell, i) { 
          
          // each column name after sheetName becomes a key of the inner obj
          var key = paramHeaders[i];
          
          // some keys have string values
          if ((key == 'remoteSheetName') || (key == 'primaryKey') || (key == 'canAddKeys')) { innerObj[key] = cell; }
          
          // others take an array
          else { innerObj[key] = [cell];   }                                                      
          
          return innerObj;
       }, {});  
      }
      else {
        // sheetName should already be in the keyMap
        row.slice(1).forEach(function (cell, i) {
          // for each row of parameter values
          if (cell) { 
            // for each row after the first (associated with this sheetName), the other columns' values will be added to the arrays associated with those columns
            var key = paramHeaders[i];
            
            obj[row[0]][key].push(cell); // loading the inner obj
          }
        });
      }
      return obj;
     }, {});
  
  return keyMap;

}

function dataObj(keyMap, sheetKey, localss) {
 
  /* sheetKey = the name of the table/tab/sheet on the local spreadsheet to load data from
     keyMap = an object holding parameters (column names) to retrieve from the local and remote data sources
     localss = a reference to the local spreadsheet (to which the script is attached)
     
     This function/class is called/instantiated twice for each table/tab/sheet on the local (currently active) spreadsheet. 
     Each instance holds either local or remote data associated with a single table/tab/sheet
     
     Most of the Google Sheets API is engaged in this function; this minimizes the number of distinct calls out to the API. 
     The other functions process this data in memory (for the most part), before writing it back to the spreadsheets at the end.
  
     [local|foreign]Data = [[A1, B1, C1, D1...],
                            [A2, B2, C2, D2...],
                            [A3, B3, C3, D3...]
                            ....]
  
  */
  
  // unique identifier for the remote source (aka the data hub)
  this.foreignssKey = '1r1AiWTgjH881fJ4Ymf-UUuNZGNEazr5-iR1cqDbgecs'; // unique spreadsheet key
  
  // get a reference to the local table/tab/sheet
  this.localSheet = localss.getSheetByName(sheetKey);
  // retrieve the active data range (where the value of any cell is not undefined)
  this.localRange = this.localSheet.getDataRange();
  // pull the data into a 2X2 array
  this.localData = this.localRange.getValues(); 
    
  // loading the relevant data from the remote table
   this.foreignss = SpreadsheetApp.openById(this.foreignssKey);
   this.foreignSheet = this.foreignss.getSheetByName(keyMap.remoteSheetName); 
   this.foreignRange = this.foreignSheet.getDataRange();
   this.foreignData = this.foreignRange.getValues();
  
  return this; // return an instance of the object itself
  
 }

  function rangeToObj (values, keyMap, ui) {
    
    /* 
    values = 2 X 2 array from the API's sheet.getDataRange().getValues() functions
    keyMap = nested object holding the relevant columns names for each table/tab/sheet on the local data source
    
    This function converts the API's native data object -- really, a 2 X 2 array -- into a nested object with properties, for ease of indexing
    
    Returns objStore = {[primary key]::int: { rowIndex: [outer index of the 2X2 array corresponding to a row number on the spreadsheet]::int,
                                                ['column name']::string: [value in that cell]::[int, string, float]
                                                ...}                                                
                       ... }
    
    */
     
    var objStore = {}, 
        // get the column indices (indices into the inner array of the 2 X 2 data array) for those columns specifiied in params, local + remote
        // the column names are stored in two arrays, one for locally owned data, one for foreign or remotely owned data
        // this object will be used to link column names (as properties) and their values in the objStore
        keyIndices = keyMap.localColumns.concat(keyMap.foreignColumns)
                           .reduce(function (prev, curr) {   
                                // if an element in the column array is undefined, skip it
                                if (!curr) { return prev; }
                                // The first row of the data array holds the first row on the table/tab/sheet, which should have column headers
                                // Find the index of this column in the row
                                var idx = values[0].indexOf(curr);  
                                // If the table/tab/sheet has extra columns not listed in the params, we don't care.
                                // But if one of our parameter columns isn't represented, that poses a problem -- skip it, but alert the user!
                                // To fix this error: Make sure that everything listed under foreignColumns and localColumns, for this sheet/table/tab, is present
                                // in both the local spreadsheet and in the data hub
                                if (idx < 0) { ui.alert('Warning: columns out of sync with data hub! Column: ' + curr); }
                                
                                // Assign to each column name it's index, e.g, {'amount': 1, 'title': 2 ...} 
                                else prev[curr] = idx;   
          
                                return prev;
        }, {});
   
    keyIndices[keyMap.primaryKey] = values[0].indexOf(keyMap.primaryKey);
    // for storing rows found without primary keys, so we can create keys for them in sequential order (even if the table is out of sequence)
    var newRows = [];
     
    // for the rest of the rows (data array[1:]), add the cell values to the object store

    values.slice(1).forEach(function (row, i) { 
      // each row is a record internal to objStore
      var record = {};
      // iterate over each of the columns
      Object.keys(keyIndices)
            .forEach(function (k) {
            // assign the value in that cell (in the data table) to its column key in the record
              record[k] = row[keyIndices[k]];   
            });

      record.rowIndex = i + 1;

      // the primary key (unique row identifier) should be one of the columns, as specified on the params sheet
      var pk = record[keyMap.primaryKey]; 
      // if there are rows newly added by the user, they won't have primary keys
      if ((pk == '') || (typeof pk == 'undefined')) { 
        // add them to the list of rows requiring new keys
        newRows.push(record); 
        
        return;
      }
      
      // associate each record with its unique ID as the key for quick lookup     
      objStore[pk] = record; 
    });
    
    // are there new, unkeyed rows?
    if (newRows.length > 0) { 
      // we need to get a sequential list of primary keys
      // unfortunately, Javascript stores keys as strings, so we need to convert them to integers before sorting
      var keys = Object.keys(objStore)
      .map(function (k) { 
        var intK = parseInt(k);
        // Javascript puts NaN's at the end, but we don't want that -- keep them at the beginning
        if (!isFinite(intK)) { intK = -1; } 
        return intK;
      })
      .sort(function (a, b) {
        return a - b;
      }),
          // get the last entry in the sequential list of primary keys
          lastKey = keys[keys.length-1];
    
   
      newRows.forEach(function (row) { 
        // increment the last key value for each new row
        lastKey++;
        // assign the new row (record) its key
        row[keyMap.primaryKey] = lastKey;
        // assign that record to its associated key as an entry in the objStore
        objStore[lastKey] = row;        
      });
    }
    return objStore;  
  }  

// iterate over the rows on each data table (sheet), updating local data with remotely owned values
  // to be done on load
function updateOnLoad () {
  
  update('foreign', 'local');
  
  }
// iterate over the rows on each data table (sheet), updating remote data with locally owned values
  // to be done on user command
function updateOnCommand () { 
  
  update('local', 'foreign');

} 

function update (source, target) {
    // local table container
    var localss = SpreadsheetApp.getActiveSpreadsheet(),
        ui = SpreadsheetApp.getUi();

    
    var keyMap = loadParams(localss),
        log = new changeLog(localss);  // initialize the logger with the local spreadsheet
        
        
    //iterate over each table/tab/sheet
    Object.keys(keyMap)
          .forEach(function (sheetKey) {
      
              var d = new dataObj(keyMap[sheetKey], sheetKey, localss),
  
               // convert the data source as an object indexed with values from the primary key property
                sourceObj = rangeToObj(d[source + 'Data'], keyMap[sheetKey], ui),
                targetData = d[target + 'Data'], 
                targetHeader = targetData[0],
                // build an object for quick lookup of the columns for updating by their index number on the target sheet (the sheet to be updated)
                sourceCols = keyMap[sheetKey][source + 'Columns']
                                                          .reduce(function (prev, curr) { 
                                                              if (curr) {
                                                                prev[targetHeader.indexOf(curr)] = curr;
                                                              }
                                                              return prev;
                                                          }, {}),
                pKey = keyMap[sheetKey].primaryKey,
                pKeyIdx = targetHeader.indexOf(pKey), // the index of the primary key column on the target sheet
                newData = [];

              sourceCols[pKeyIdx] = pKey;
              
              // iterate over the 2-d array of values from the target sheet 
              for (var i=1; i<targetData.length; i++) {
            
                 var pk = targetData[i][pKeyIdx],  // find the record corresponding to this row from the source data set
                   sourceRecord = sourceObj[pk];
                 
                 if (!sourceRecord) { // if we don't have a match, just skip it --> TO DO: Alert the user?
                   continue;
                 }

                  var updateRow = targetData[i].map(function (cell, j) { 
                   
                     if (j == pKeyIdx) { return cell; }
                    
                     var colKey = sourceCols[j]; 
                   
                     if (!colKey) { return cell; }  // if this isn't a column for updating, ignore it
                      
                     // if this is a cell from a column for updating, and if its value doesn't match the value in the source data set for this row, col, 
                     // we assign the new value to this cell and log the change
                     if (cell != sourceRecord[colKey]) { 
                        log.updateLog({timestamp: new Date(), 
                                      column_name: colKey, 
                                      table: target, 
                                      data_key: pk, 
                                      old_value: cell, 
                                      new_value: sourceRecord[colKey]}); 
                      // in the event of blank cells on the source, we don't want to display "undefined"
                      return sourceRecord[colKey] || '';
                    }
                   return cell;
                  });
                 // update the 2-d array 
                 targetData[i] = updateRow;
                 // mark that we've already seen this record in the source data set
                 delete sourceObj[pk];
               }
             // broadcast the updated 2-D array back to the target sheet
             d[target + 'Range'].setValues(targetData);
           
             // any remaining records from the source = new rows for the target sheet
             Object.keys(sourceObj).forEach(function (sourceKey) { 
               newData.push(sourceObj[sourceKey]);
             });
             // some users don't have permission to add new rows; in those cases, we don't want to update the data hub (e.g., source == 'local')
             if ((newData.length > 0) && ((source == 'foreign') || (keyMap[sheetKey].canAddKeys))) { 
          
               keyMap[sheetKey][target + 'Columns'].forEach(function (col) { 
                    if (col) {
                      sourceCols[targetHeader.indexOf(col)] = col;
                    }
               });
               appendNewValues(newData, d[target + 'Sheet'], sourceCols); // if there are new rows of data, add these to the end of the target sheet
               addNewKeys(newData, d[source + 'Range'], d[source + 'Data'], pKey); // new rows (on the source) will have new primary keys -- need to write them to the source sheet
               
               if (target != 'foreign') { 
                 ui.alert('Alert: ' + newData.length + ' rows of new data added to the bottom of sheet "' + sheetKey + '"');
               }
             }
              
           log.writeLog();  
              
           });
    
    SpreadsheetApp.flush()    
    
  }

function addNewKeys (newData, sourceRange, sourceData, primaryKey) { 

  pKeyIdx = sourceData[0].indexOf(primaryKey);
  
  // for each of the source records, we have recorded its row index on the source sheet
  newData.forEach(function (d) { 
    var rowIndex = d.rowIndex,
        key = d[primaryKey];
    // assign the new keys to the appropriate cells (on the source sheet)
    sourceData[rowIndex][pKeyIdx] = key;
  });
  // broadcast to the spreadsheet
  sourceRange.setValues(sourceData);
 
}

function appendNewValues (newData, sheet, targetCols) { 
 // add new values onto the end of the range
  var columns = Object.keys(targetCols)
                      .sort(function (a, b) { 
                        return a - b;  // sort the columns according to their index on the sheet
                      }),
      lastRow = sheet.getLastRow(), // find the last row with content in it
      newRange = sheet.getRange(lastRow + 1, 1, newData.length, columns.length),  // get the range that would cover the number of additional cells
      newDataArray = newData.reduce(function (arry, rowObj) {    
        arry.push(columns.map(function (c) { 
          var col = targetCols[c];
          return (rowObj[col]) ? rowObj[col] : '';  // if the value in this cell isn't null, add it to the data array, using the empty string for nulls (to avoid 'undefined')
        }));
        return arry;
      }, new Array());
  // broadcast to the spreadsheet
  newRange.setValues(newDataArray);
  

}

      
// Add menu items for update functions and to open doc pane
function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createMenu("Serials Data");//.createAddonMenu(); // Or DocumentApp or FormApp.
      // Add a normal menu item (works in all authorization modes).
  menu.addItem('Update Data Hub', 'updateOnCommand')
       .addItem('Get Latest Data', 'updateOnLoad')
       .addItem('Show Documentation', 'showSidebar')
       .addToUi();
  
  updateOnLoad();
  
  showSidebar();

} 



function changeLog (spreadsheet) {
  // holds changes for display in a separate sheet, change_log
  // change_log should have the following columns: timestamp, column_name, old_value, new_value
  
   this.columns = ['timestamp', 'table', 'data_key', 'column_name', 'old_value', 'new_value'];
  
   this.sheet = spreadsheet.getSheetByName('change_log'); // Open the "change log" sheet in the given spreadsheet --> TO DO: error checking for no change 
   this.lastRow = this.sheet.getLastRow(); // the last row with content
   this.changes = []; // Array to hold the changes, one per row
      
   this.updateLog = function (changeObj) {
     // accepts an object mapping data to column names on log sheet
     // converts this object to an array for setting cell values
     this.changes.push(this.columns.map(function (elem) { 
       return changeObj[elem];
     }));
   }
   // write new changes to the end of log sheet
   this.writeLog = function () {
     if (this.changes.length > 0) {
       this.sheet.getRange(this.lastRow + 1, 1, this.changes.length, this.columns.length)
              .setValues(this.changes);
     }
   } 
   
  
}


function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar_doc.html')
      .setWidth(300)
      .setTitle('documentation');
      
  SpreadsheetApp.getUi() 
      .showSidebar(html);
}

// this function to be run on the data hub, tied to a timed trigger (for weekly execution)
function backupSheet () {
   var ss = SpreadsheetApp.getActive(),
     today = new Date(),
     newSS = ss.copy('data_hub_backup_' + (today.getMonth() + 1) + '-' + today.getDate() + '-' + + today.getFullYear()).getId(), // create a copy of the current spreadsheet
     backupFolder = DriveApp.getRootFolder().getFoldersByName('backups').next(), // get the folder where the backups are stored
     newFile = DriveApp.getFileById(newSS); // get a reference to the file holding this spreadsheet
     
   backupFolder.addFile(newFile); // add the backup file to the folder

}

/**
 * ServerLogic.gs - Backend functions for the Attendance Utility Web App.
 * Uses Drive API and SpreadsheetApp to manage user sheets and attendance data.
 */

// Global constant for attendance tracking, assuming date is in the first row.
const HEADER_ROW = 1;

/**
 * Entry point for the Web App.
 * @return {GoogleAppsScript.HTML.HtmlOutput} The served HTML content.
 */
function doGet() {
  const htmlOutput = HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Attendance Marker');
  
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  return htmlOutput;
}

/**
 * Uses the Drive API to fetch only the Spreadsheets the user owns,
 * prioritizing performance and relevance.
 * NOTE: Requires the Drive API to be enabled in Apps Script Services.
 * @return {Array<Object>} An array of sheet objects {id, name}.
 */
function getSpreadsheets() {
  
  // *** STEP 1: Use the correct v3 query term: 'me' in writers ***
  // This query requests files that are Spreadsheets AND not trashed AND the current user ('me') 
  // is listed as a writer (editor) on the file.
  const query = "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false and 'me' in writers";
  
  const files = Drive.Files.list({
    q: query,
    // CRITICAL: Request the capabilities field to verify edit access locally.
    fields: 'files(id, name, capabilities)', 
    corpora: 'user', 
    maxResults: 100 
  });

  const sheetList = [];
  
  if (files.files) {
    files.files.forEach(file => {
      // **STEP 2: Local Filtering**
      // Check the returned file object's capabilities property for edit permission.
      // This is the definitive safety check.
      const canEdit = file.capabilities && file.capabilities.canEdit;
      
      if (canEdit) {
        sheetList.push({
          id: file.id,
          name: file.name
        });
      }
    });
  }
  
  // Note: If the query fails due to an API restriction, the error handler will catch it,
  // but if it just returns an incomplete list, this is the most the API will reliably provide.
  return sheetList;
}

/**
 * ALWAYS inserts a new column after the last data column for a new attendance session.
 * This supports multiple classes/sessions on the same date.
 * @param {string} sheetId - ID of the target spreadsheet.
 * @param {string} sheetName - Name of the specific sheet/tab within the spreadsheet.
 * @param {string} dateString - The date selected by the user (YYYY-MM-DD format).
 * @return {number} The column index (1-based) where attendance should be written.
 */
function findOrCreateAttendanceColumn(sheetId, sheetName, dateString) {
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet/Tab named "${sheetName}" not found in the spreadsheet.`);
  }

  const timezone = ss.getSpreadsheetTimeZone();
  const lastCol = sheet.getLastColumn();
  
  // Create the target date using Utilities.parseDate for timezone correctness.
  const targetDate = Utilities.parseDate(dateString, timezone, 'yyyy-MM-dd');

  // --- NEW LOGIC: Always insert a column after the last data column ---
  const newColIndex = lastCol + 1;
  
  // Use 1 if the sheet is completely empty, otherwise use lastCol
  const columnToInsertAfter = (lastCol === 0) ? 0 : lastCol; 
  
  sheet.insertColumnAfter(columnToInsertAfter);
  
  // Set the new column header to the target date
  const dateCell = sheet.getRange(HEADER_ROW, newColIndex);
  
  // Write the correctly parsed date to the sheet
  dateCell.setValue(targetDate);
  
  // Set date format for visual clarity in the sheet (e.g., "Tue, Nov 11")
  dateCell.setNumberFormat('ddd, MMM d'); 
  
  // Optional: Also write the time just below the date to distinguish the session
  // const timeCell = sheet.getRange(HEADER_ROW + 1, newColIndex);
  // timeCell.setValue(new Date());
  // timeCell.setNumberFormat('h:mm a'); // E.g., 9:30 AM
  
  return newColIndex;
}


/**
 * Reads student names and current attendance status for the given sheet.
 * @param {string} sheetId - ID of the target spreadsheet.
 * @param {string} sheetName - Name of the specific sheet/tab within the spreadsheet.
 * @param {number} nameColIndex - Column index (1-based) where names are.
 * @param {number} startRowIndex - Row index (1-based) where the first name is.
 * @param {number} attendanceColIndex - The column index (1-based) for today's attendance.
 * @return {Array<Object>} Roster data: {name, rowIndex, status}.
 */
function getRosterData(sheetId, sheetName, nameColIndex, startRowIndex, attendanceColIndex) {
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(sheetName);
  
  const lastRow = sheet.getLastRow();
  if (lastRow < startRowIndex) {
    throw new Error("Roster has no students starting at the specified row.");
  }
  
  // Get all names and corresponding attendance marks
  const numRows = lastRow - startRowIndex + 1;
  const nameRange = sheet.getRange(startRowIndex, nameColIndex, numRows, 1).getDisplayValues();
  const attendanceRange = sheet.getRange(startRowIndex, attendanceColIndex, numRows, 1).getDisplayValues();
  
  const roster = [];
  for (let i = 0; i < numRows; i++) {
    const name = nameRange[i][0].trim();
    if (name) { // Skip blank rows
      roster.push({
        name: name,
        rowIndex: startRowIndex + i, // Store the actual row index (1-based)
        status: attendanceRange[i] ? attendanceRange[i][0].toUpperCase() : '' // Current mark, if any
      });
    }
  }
  return roster;
}

/**
 * Writes the attendance mark back to the sheet.
 * @param {string} sheetId - ID of the target spreadsheet.
 * @param {string} sheetName - Name of the specific sheet/tab within the spreadsheet.
 * @param {number} rowIndex - The row index (1-based) of the student.
 * @param {number} colIndex - The column index (1-based) for today's attendance.
 * @param {string} status - 'P', 'A', or 'L'.
 */
function markAttendance(sheetId, sheetName, rowIndex, colIndex, status) {
  try {
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    if (!sheet) {
        throw new Error(`Sheet/Tab named "${sheetName}" not found.`);
    }
    
    // Writes the status (P, A, L, etc.) to the specific cell
    sheet.getRange(rowIndex, colIndex).setValue(status);
    
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

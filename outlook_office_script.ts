/**
 * Office Script: Transfer Outlook Emails to Excel
 * This script receives email data from Power Automate and writes it to Excel
 * 
 * HOW IT WORKS:
 * 1. Power Automate gets emails from your Outlook folder
 * 2. Power Automate calls this script and passes the email data
 * 3. This script writes the emails to rows in Excel
 */

/**
 * Main function that Power Automate will call
 * This function receives an array of email objects and writes them to the sheet
 * 
 * @param workbook - The Excel workbook (this is provided automatically)
 * @param emailDate - The date the email was received
 * @param emailFrom - Who sent the email
 * @param emailSubject - The subject line of the email
 * @param emailBody - Preview of the email body (first 500 characters)
 * @param emailFolder - Which folder the email came from
 */
function main(
  workbook: ExcelScript.Workbook,
  emailDate: string,
  emailFrom: string,
  emailSubject: string,
  emailBody: string,
  emailFolder: string
) {
  // Get the first worksheet in the workbook
  let sheet = workbook.getActiveWorksheet();
  
  // Check if this is the first time running - if so, add headers
  // We check if cell A1 is empty to know if we need headers
  let firstCell = sheet.getRange("A1").getValue();
  
  if (firstCell === "" || firstCell === null) {
    // This is the first run, so create headers
    createHeaders(sheet);
  }
  
  // Find the next empty row to add our data
  // We look in column A to find the first empty cell
  let nextRow = findNextEmptyRow(sheet);
  
  // Create the row of data from the email information
  let rowData = [
    emailDate,      // Column A: Date received
    emailFrom,      // Column B: From address
    emailSubject,   // Column C: Subject
    emailBody,      // Column D: Body preview
    emailFolder     // Column E: Folder name
  ];
  
  // Write the data to the next empty row
  // We start at column A (index 0) and write 5 columns (A through E)
  let targetRange = sheet.getRangeByIndexes(
    nextRow - 1,  // Row index (subtract 1 because Excel uses 0-based indexing)
    0,            // Column A (0 = A, 1 = B, etc.)
    1,            // Number of rows (just 1 row)
    5             // Number of columns (5 columns: A, B, C, D, E)
  );
  
  // Set the values in the range
  targetRange.setValues([rowData]);
  
  // Auto-resize the columns so everything fits nicely
  sheet.getUsedRange().getFormat().autofitColumns();
  
  // Return a success message (you can see this in Power Automate's run history)
  return `Email added successfully to row ${nextRow}`;
}

/**
 * Helper function: Create header row
 * This sets up the column headers in the first row
 * 
 * @param sheet - The worksheet to add headers to
 */
function createHeaders(sheet: ExcelScript.Worksheet) {
  // Define the header names
  let headers = ["Date Received", "From", "Subject", "Body Preview", "Folder"];
  
  // Write headers to row 1
  sheet.getRange("A1:E1").setValues([headers]);
  
  // Make the headers bold so they stand out
  sheet.getRange("A1:E1").getFormat().getFont().setBold(true);
  
  // Add a light background color to the headers (optional but looks nice)
  sheet.getRange("A1:E1").getFormat().getFill().setColor("#D3D3D3");
}

/**
 * Helper function: Find the next empty row
 * This looks in column A to find where to add the next email
 * 
 * @param sheet - The worksheet to search
 * @returns The row number of the next empty row
 */
function findNextEmptyRow(sheet: ExcelScript.Worksheet): number {
  // Get all the used cells in the sheet
  let usedRange = sheet.getUsedRange();
  
  // If the sheet is empty, return row 2 (row 1 will be for headers)
  if (usedRange === undefined) {
    return 2;
  }
  
  // Get the last row that has data
  let lastRow = usedRange.getRowIndex() + usedRange.getRowCount();
  
  // Return the next row after the last used row
  return lastRow + 1;
}

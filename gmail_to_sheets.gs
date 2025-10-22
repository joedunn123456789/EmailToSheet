/**
 * Gmail to Google Sheets Transfer Script
 * This script extracts emails from Gmail and writes them to a Google Sheet
 */

/**
 * Main function that transfers emails from Gmail to Google Sheets
 * This is the function you'll run to start the transfer
 */
function transferEmailsToSheets() {
  // Get the active spreadsheet (the one that's currently open)
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Clear any existing content in the sheet to start fresh
  sheet.clear();
  
  // Create headers for our columns so we know what each column represents
  var headers = ["Date", "From", "Subject", "Body Preview", "Labels"];
  sheet.appendRow(headers);
  
  // Define your search query - this determines which emails to get
  // This searches for emails with the "Job Hunting" label
  var searchQuery = "label:Job Hunting";
  
  // Set how many emails you want to transfer at once
  var maxEmails = 100; // Change this number to get more or fewer emails
  
  // Search Gmail for threads (email conversations) matching your query
  var threads = GmailApp.search(searchQuery, 0, maxEmails);
  
  // Loop through each thread (conversation) we found
  for (var i = 0; i < threads.length; i++) {
    // Get all individual messages in this thread
    var messages = threads[i].getMessages();
    
    // Loop through each individual message in the thread
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      
      // Extract the date the email was received
      var date = message.getDate();
      
      // Extract who sent the email
      var from = message.getFrom();
      
      // Extract the email subject line
      var subject = message.getSubject();
      
      // Extract the email body (plain text version)
      // We limit it to 500 characters so it doesn't get too long
      var body = message.getPlainBody().substring(0, 500);
      
      // Get all labels (folders/categories) attached to this thread
      var labels = threads[i].getLabels().map(function(label) {
        return label.getName();
      }).join(", "); // Join multiple labels with commas
      
      // Create a row with all the data we extracted
      var rowData = [date, from, subject, body, labels];
      
      // Add this row to the spreadsheet
      sheet.appendRow(rowData);
    }
  }
  
  // Format the date column to look nice
  var dateRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  dateRange.setNumberFormat("yyyy-mm-dd hh:mm:ss");
  
  // Auto-resize all columns so the content fits nicely
  sheet.autoResizeColumns(1, 5);
  
  // Show a success message
  SpreadsheetApp.getUi().alert('Job Hunting email transfer complete! ' + 
    (sheet.getLastRow() - 1) + ' emails have been added to the sheet.');
}

/**
 * Optional: Function to transfer only unread emails
 * This is useful if you only want to see new emails
 */
function transferUnreadEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  
  var headers = ["Date", "From", "Subject", "Body Preview", "Labels"];
  sheet.appendRow(headers);
  
  // This query specifically searches for unread emails
  var searchQuery = "is:unread";
  var maxEmails = 50;
  
  var threads = GmailApp.search(searchQuery, 0, maxEmails);
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var date = message.getDate();
      var from = message.getFrom();
      var subject = message.getSubject();
      var body = message.getPlainBody().substring(0, 500);
      var labels = threads[i].getLabels().map(function(label) {
        return label.getName();
      }).join(", ");
      
      var rowData = [date, from, subject, body, labels];
      sheet.appendRow(rowData);
    }
  }
  
  var dateRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  dateRange.setNumberFormat("yyyy-mm-dd hh:mm:ss");
  sheet.autoResizeColumns(1, 5);
  
  SpreadsheetApp.getUi().alert('Unread email transfer complete! ' + 
    (sheet.getLastRow() - 1) + ' unread emails have been added.');
}

/**
 * Optional: Function to transfer emails from a specific sender
 * Modify the email address in the searchQuery variable
 */
function transferEmailsFromSender() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  
  var headers = ["Date", "From", "Subject", "Body Preview", "Labels"];
  sheet.appendRow(headers);
  
  // Change "example@gmail.com" to the email address you want to search for
  var searchQuery = "from:example@gmail.com";
  var maxEmails = 50;
  
  var threads = GmailApp.search(searchQuery, 0, maxEmails);
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var date = message.getDate();
      var from = message.getFrom();
      var subject = message.getSubject();
      var body = message.getPlainBody().substring(0, 500);
      var labels = threads[i].getLabels().map(function(label) {
        return label.getName();
      }).join(", ");
      
      var rowData = [date, from, subject, body, labels];
      sheet.appendRow(rowData);
    }
  }
  
  var dateRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  dateRange.setNumberFormat("yyyy-mm-dd hh:mm:ss");
  sheet.autoResizeColumns(1, 5);
  
  SpreadsheetApp.getUi().alert('Email transfer complete!');
}

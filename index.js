function doPost(e) {
  // CORS configuration helper - Apps Script Web App redirects are automatically handled by browser,
  // returning JSON ContentService automatically sets proper CORS headers for the client browser.
  try {
    // Check if postData content exists
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse('error', 'No post data received in request.');
    }
    // Parse incoming JSON payload
    var data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (parseError) {
      return createJsonResponse('error', 'Invalid JSON payload received.');
    }
    // Extract fields
    var name = data.name ? String(data.name).trim() : '';
    var number = data.number ? String(data.number).trim() : '';
    var email = data.email ? String(data.email).trim() : '';
    var subjects = data.subjects ? String(data.subjects).trim() : '';
    // Validate inputs server-side
    if (!name || !number || !email || !subjects) {
      return createJsonResponse('error', 'All fields (Name, Number, Email, Subjects) are required.');
    }
    // Phone number verification (digits only)
    if (!/^\d+$/.test(number)) {
      return createJsonResponse('error', 'Phone number must contain digits only.');
    }
    // Email address basic verification
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return createJsonResponse('error', 'Invalid email address format.');
    }
    // Get the active spreadsheet and sheet
    // Note: If you are using a standalone script (not bound to a sheet),
    // you can replace this with: SpreadsheetApp.openById("YOUR_SPREADSHEET_ID")
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    // Check if sheet is empty and initialize headers if needed
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Name', 'Phone Number', 'Email', 'A Level Subjects']);
      // Apply clean formatting to header row
      var headerRange = sheet.getRange(1, 1, 1, 5);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#002147'); // Aura Navy
      headerRange.setFontColor('#FFFFFF');   // White Text
      headerRange.setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
    }
    // Create a local timestamp
    var timestamp = new Date();
    // Append form data as a new row
    sheet.appendRow([
      timestamp,
      name,
      number,
      email,
      subjects
    ]);
    // Return success response to the client
    return createJsonResponse('success', 'Application submitted successfully.');
  } catch (error) {
    Logger.log('Error processing form: ' + error.toString());
    return createJsonResponse('error', 'Internal server error: ' + error.toString());
  }
}
/**
 * Helper function to generate standardized JSON responses
 */
function createJsonResponse(status, message) {
  var output = {
    status: status,
    message: message
  };
  
  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}
/**
 * Handle OPTIONS preflight requests (optional, but good practice for robustness)
 */
function doOptions(e) {
  var output = ContentService.createTextOutput('');
  return output.setMimeType(ContentService.MimeType.TEXT);
}
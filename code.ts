const SCRIPT_PROP = PropertiesService.getScriptProperties();

function setup() {
  const doc = SpreadsheetApp.getActiveSpreadsheet();
  SCRIPT_PROP.setProperty('key', doc.getId());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty('key') || SpreadsheetApp.getActiveSpreadsheet().getId());
    
    // Parse the incoming JSON payload
    let data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      data = e.parameter;
    }

    const type = data.type; // "RSVP" or "Wishes"
    
    if (!type) {
      return ContentService
        .createTextOutput(JSON.stringify({ result: 'error', error: 'Missing type field' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    let sheet = doc.getSheetByName(type);
    
    // Auto-generate sheet and headers if it doesn't exist
    if (!sheet) {
      sheet = doc.insertSheet(type);
      if (type === "RSVP") {
        sheet.appendRow(["Timestamp", "Full Name", "Guests", "Dietary Notes"]);
      } else if (type === "Wishes") {
        sheet.appendRow(["Timestamp", "Name", "Message"]);
      }
      // Make headers bold
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nextRow = sheet.getLastRow() + 1;
    const newRow = [];

    // Map incoming data to correct columns based on headers
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      if (header === 'Timestamp') {
        newRow.push(new Date());
      } else if (type === "RSVP") {
        if (header === 'Full Name') newRow.push(data.fullName || '');
        else if (header === 'Guests') newRow.push(data.guests || '');
        else if (header === 'Dietary Notes') newRow.push(data.dietaryNotes || '');
        else newRow.push('');
      } else if (type === "Wishes") {
        if (header === 'Name') newRow.push(data.name || '');
        else if (header === 'Message') newRow.push(data.message || '');
        else newRow.push('');
      }
    }

    sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', row: nextRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Handle GET requests (e.g., to verify the web app is running)
function doGet(e) {
  return ContentService.createTextOutput("Web App is running. Use POST to submit data.");
}

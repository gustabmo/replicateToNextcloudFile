/**
 * Google Apps Script to replicate specific tabs from the current spreadsheet to an existing Nextcloud file.
 * 
 * by Gustavo Exel guexel@gmail.com
 * 
 * Overwrites an existing Nextcloud file with google sheet data.
 * The Nextcloud share link and file ID will remain unchanged.
 * See "How to use" at the end of this file for instructions.
 */
function replicateToNextcloudFile(spreadheetId) {
  const ss = SpreadsheetApp.openById(spreadheetId);
  const config = ss.getSheetByName("Config");
  
  // Get config values from column B
  // B1 must be a WebDAV URL, e.g.: https://cloud.example.com/remote.php/dav/files/username/path/to/file.xlsx
  const ncUrl = config.getRange("B1").getValue();
  const username = config.getRange("B2").getValue();
  const appPassword = config.getRange("B3").getValue();
  const tabsToSync = config.getRange("B4").getValue().split(',').map(s => s.trim());

  if (!ncUrl.includes('/remote.php/dav/') && !ncUrl.includes('/dav/')) {
    throw new Error(
      'B1 must be a Nextcloud WebDAV URL, not a share link.\n' +
      'Expected format: https://cloud.example.com/remote.php/dav/files/USERNAME/path/to/file.xlsx\n' +
      'Current value: ' + ncUrl
    );
  }

  // 1. Create a temporary spreadsheet for a clean export (values/formatting only)
  const tempSS = SpreadsheetApp.create("Temp_Export_replicateToNextcloudFile");
  const defaultSheet = tempSS.getSheets()[0];
  defaultSheet.setName("__TEMP_THIS_WILL_BE_DELETED__");
  const tempId = tempSS.getId();

  tabsToSync.forEach(tabName => {

    const sourceSheet = ss.getSheetByName(tabName);
    if (!sourceSheet) return;

    const sourceRange = sourceSheet.getDataRange();
    const values = sourceRange.getDisplayValues();
    const formulas = sourceRange.getFormulas();

    const newSheet = tempSS.insertSheet(tabName);

    // Write values only (no formulas at all)
    newSheet.getRange(1, 1, values.length, values[0].length).setValues(values);

    // Remove problematic metadata
    newSheet.clearConditionalFormatRules();

    // Remove filters
    const filter = newSheet.getFilter();
    if (filter) filter.remove();

    // Remove protections
    newSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .forEach(p => p.remove());

    newSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .forEach(p => p.remove());
  });

  tempSS.deleteSheet(defaultSheet); // remove Sheet1 that is created by default

  // 2. Export temp spreadsheet as an .xlsx blob
  const exportUrl = `https://docs.google.com/spreadsheets/d/${tempId}/export?format=xlsx`;
  const token = ScriptApp.getOAuthToken();
  const excelBlob = UrlFetchApp.fetch(exportUrl, {
    headers: { 'Authorization': 'Bearer ' + token }
  }).getBlob();

  // 3. Perform the PUT request to overwrite the existing file
  const authHeader = 'Basic ' + Utilities.base64Encode(username + ':' + appPassword);
  const options = {
    method: "put",
    payload: excelBlob.getBytes(),
    headers: { 
      "Authorization": authHeader,
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(ncUrl, options);
  
  // Final Cleanup
  DriveApp.getFileById(tempId).setTrashed(true);

  if (response.getResponseCode() < 300) {
    // success!!!
  } else {
    Logger.log("HTTP " + response.getResponseCode() + ": " + response.getContentText());
    throw new Error("Upload failed: HTTP " + response.getResponseCode() + " — check Logs for details.");
  }
}



// How to use:
//
// these functions should be added to the script of a spreadsheet that will be the "source" for the data to replicate to nextcloud. It will add a menu item to trigger the replication.
// here, in a standalone script, they won't do anything, but they are included here for convenience.
function replicateThisSheet () {
  replicateToNextcloudFile.replicateToNextcloudFile(SpreadsheetApp.getActiveSpreadsheet().getId());
}

function onOpen() {
    SpreadsheetApp.getUi()
    .createMenu("Replicate")
    .addItem("To Nextcloud", "replicateThisSheet")
    .addToUi();
}
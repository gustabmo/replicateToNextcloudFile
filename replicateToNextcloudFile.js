/**
 * Google Apps Script to overwrite an existing Nextcloud file with google sheet data from specific tabs
 * Only values and formatting are replicated, no formulas, filters, protections or conditional formatting.
 * See "How to use" at the end of this file for instructions.
 * 
 * by Gustavo Exel guexel@gmail.com 2026-05-09
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

    // Copy full sheet with formatting
    const newSheet = sourceSheet.copyTo(tempSS).setName(tabName);

    const sourceRange = sourceSheet.getDataRange();
    const values = sourceRange.getDisplayValues();

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


/**
 * How to use:
 * 
 * 1. Create a google apps script project with this file in it and call it "replicateToNextcloudFile" or whatever you like.
 * 2. Create or locate a new google spreadsheet that will be the "source" of the data to replicate to nextcloud.
 * 3. In that spreadsheet, create a tab named "Config" and fill in the following values in column B:
 *    B1: WebDAV URL of the target Nextcloud file, e.g. https://nextcloud.example.com/remote.php/dav/files/username/path/to/file.xlsx
 *    B2: Nextcloud username (or email)
 *    B3: Nextcloud app password (you must create an app password in your Nextcloud account settings, your regular password will not work here)
 *    B4: Comma-separated list of tab names to replicate, e.g. "Sheet1,Data,Summary"
 * 4. In the same spreadsheet, open the Apps Script editor and copy the functions "replicateThisSheet" and "onOpen" from the end of this comment into the script editor of that spreadsheet.
 * 5. In the Apps Script editor, go to "Resources" > "Libraries", and add the library with the project key of your "replicateToNextcloudFile" script project. Select the latest version and give it a short identifier (e.g. "replicateToNextcloudFile").
 * 6. Save everything and reload the spreadsheet. You should see a new menu item "Replicate" with an option "To Nextcloud". Click it to start the replication.
 * 7. Check your Nextcloud file to see if it has been updated with the data from the specified tabs. The file ID and share link should remain unchanged, only the content will be overwritten.
 * 
 * Note: This script only replicates values and formatting. Formulas, filters, protections and conditional formatting will not be replicated.
 * The script creates a temporary spreadsheet to prepare the data for export, then exports it as an .xlsx file and uploads it to Nextcloud using a PUT request. Finally, it cleans up the temporary file.
 * Make sure to handle your Nextcloud credentials securely and do not share the script with others if it contains sensitive information.
 * 
 * If you want to automate this process, you can set up a time-driven trigger in the Apps Script editor to run the "replicateThisSheet" function at regular intervals (e.g. daily).
 * 
 * For any issues or improvements, feel free to contact me at guexel@gmail.com

function replicateThisSheet () {
  replicateToNextcloudFile.replicateToNextcloudFile(SpreadsheetApp.getActiveSpreadsheet().getId());
}

function onOpen() {
    SpreadsheetApp.getUi()
    .createMenu("Replicate")
    .addItem("To Nextcloud", "replicateThisSheet")
    .addToUi();
}

*
*
*/


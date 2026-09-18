/**
 * NSPPD UK Medical Unit — Volunteer Declaration backend
 *
 * This is already deployed and live at:
 * https://script.google.com/macros/s/AKfycbw8YZ4RPhchG2mvmoSvLIrdiPqLV_chm0GVD6Q0ZG9puz_M6jfBpWY1ONq63xdK6fbo/exec
 * (that URL is already plugged into index.html in this folder)
 *
 * Kept here only as a reference copy of what's running in your Google
 * Apps Script project. Re-paste and redeploy only if you need to change
 * its behaviour.
 *
 * SETUP
 * 1. Create (or open) the Google Sheet you want submissions saved to.
 * 2. In the Sheet, go to Extensions > Apps Script.
 * 3. Delete any starter code and paste in this whole file.
 * 4. In row 1 of your sheet, add the headers: Timestamp | Name | Volunteer ID | Phone Number
 * 5. Click Deploy > New deployment.
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Click Deploy, authorize the script when prompted, and copy the
 *    Web app URL it gives you (it ends in /exec).
 * 7. Paste that URL into the SCRIPT_URL constant near the top of the
 *    <script> section in the e-learning HTML file (search for
 *    "PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE").
 *
 * This script rejects a submission if the Volunteer ID already exists
 * in the sheet (case-insensitive, trimmed), so no one can submit twice.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    var data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      return jsonOutput({ status: "error", message: "Could not read submission data." });
    }

    var name = (data.name || "").toString().trim();
    var volunteerId = (data.volunteerId || "").toString().trim();
    var phone = (data.phone || "").toString().trim();

    if (!name || !volunteerId) {
      return jsonOutput({ status: "error", message: "Name and Volunteer ID are both required." });
    }

    // Make sure a header row exists
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "Name", "Volunteer ID", "Phone Number"]);
    }

    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      var existingId = (values[i][2] || "").toString().trim().toLowerCase();
      if (existingId === volunteerId.toLowerCase()) {
        return jsonOutput({
          status: "error",
          message: "This Volunteer ID has already submitted the declaration. Each volunteer may only submit once."
        });
      }
    }

    sheet.appendRow([new Date(), name, volunteerId, phone]);
    return jsonOutput({ status: "success", message: "Declaration recorded. Thank you." });

  } finally {
    lock.releaseLock();
  }
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

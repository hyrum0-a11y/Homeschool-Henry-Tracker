/**
 * Google Apps Script — Email Verification Code Sender
 *
 * SETUP INSTRUCTIONS:
 * 1. Go to https://script.google.com and create a new project
 * 2. Paste this entire file into Code.gs (replace the default code)
 * 3. Click Deploy → New deployment
 * 4. Select type: "Web app"
 * 5. Set "Execute as": Me (your Google account)
 * 6. Set "Who has access": Anyone
 * 7. Click Deploy and authorize when prompted
 * 8. Copy the Web app URL (looks like: https://script.google.com/macros/s/XXXXX/exec)
 * 9. Add to your .env file: APPS_SCRIPT_URL=https://script.google.com/macros/s/XXXXX/exec
 * 10. Restart the server
 *
 * HOW IT WORKS:
 * The Sovereign HUD server calls this web app with ?email=...&code=...&name=...
 * This script sends an email with the verification code to the specified address.
 * The code is NEVER stored in the Google Sheet — only in the server's memory.
 */

function doGet(e) {
  var email = e.parameter.email;
  var code = e.parameter.code;
  var name = e.parameter.name || "User";

  if (!email || !code) {
    return ContentService.createTextOutput("Missing email or code parameter")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  var subject = "Your login code is " + code;
  var body = "Hi " + name + ",\n\n"
    + "Your code: " + code + "\n\n"
    + "Enter this on the login page. Expires in 10 minutes.";

  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      name: name
    });
    return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput("Error: " + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

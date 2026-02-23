/**
 * Google Apps Script — Email Sender for Sovereign HUD
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
 * FUNCTIONS:
 * doGet  — Sends plain-text verification code emails (called with ?email=...&code=...&name=...)
 * doPost — Sends HTML emails (called with JSON body: { action: "sendHtml", email, subject, body })
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

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.action === "sendHtml") {
      if (!data.email || !data.subject || !data.body) {
        return ContentService.createTextOutput(JSON.stringify({ status: "Error", message: "Missing email, subject, or body" }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      MailApp.sendEmail({
        to: data.email,
        subject: data.subject,
        htmlBody: data.body,
        name: "Sovereign HUD"
      });

      return ContentService.createTextOutput(JSON.stringify({ status: "OK" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "Error", message: "Unknown action: " + data.action }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "Error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

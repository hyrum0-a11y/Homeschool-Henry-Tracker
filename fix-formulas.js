/**
 * One-time script to update Sectors sheet formulas from column-index VLOOKUPs
 * to header-resilient INDEX/MATCH formulas.
 *
 * Before running:
 *   - Service account must have Editor access on the sheet
 *
 * Usage: node fix-formulas.js
 */

const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");

// Load .env
try {
  const envFile = fs.readFileSync(path.join(__dirname, ".env"), "utf8");
  for (const line of envFile.split("\n")) {
    const t = line.trim();
    if (!t || t.startsWith("#")) continue;
    const eq = t.indexOf("=");
    if (eq === -1) continue;
    const k = t.slice(0, eq).trim();
    const v = t.slice(eq + 1).trim();
    if (!process.env[k]) process.env[k] = v;
  }
} catch {}

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const CREDENTIALS_PATH = process.env.GOOGLE_CREDENTIALS_PATH || "./credentials.json";

async function main() {
  const creds = JSON.parse(fs.readFileSync(path.resolve(__dirname, CREDENTIALS_PATH), "utf8"));
  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  const sheets = google.sheets({ version: "v4", auth });

  // Find last row with data in Sectors column A
  const colA = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sectors!A:A",
  });
  const lastRow = colA.data.values ? colA.data.values.length : 1;
  const dataRows = lastRow - 1; // excluding header
  console.log(`Found ${dataRows} data rows in Sectors sheet.`);

  if (dataRows < 1) {
    console.log("No data rows found. Nothing to update.");
    return;
  }

  // Build fully header-based formulas for rows 2..lastRow
  // Every column is found by MATCH on header name — no column letters for data.
  //
  // Pattern:
  //   INDEX($A:$Z, ROW(), MATCH("Impact(1-3)", $1:$1, 0))          ← Impact from current row
  //   INDEX($A:$Z, ROW(), MATCH("Sector", $1:$1, 0))               ← Sector from current row
  //   INDEX(Definitions!$A:$Z, , MATCH("Sector", Definitions!$1:$1, 0)) ← Sector col in Definitions
  //   MATCH("INTELLIGENCE", Definitions!$1:$1, 0)                   ← target stat col in Definitions
  //
  const statNames = ["INTELLIGENCE", "STAMINA", "TEMPO", "REPUTATION"];
  const formulas = [];
  for (let row = 2; row <= lastRow; row++) {
    formulas.push(
      statNames.map(
        (stat) =>
          `=INDEX($A:$Z,ROW(),MATCH("Impact(1-3)",$1:$1,0))*INDEX(Definitions!$A:$Z,MATCH(INDEX($A:$Z,ROW(),MATCH("Sector",$1:$1,0)),INDEX(Definitions!$A:$Z,,MATCH("Sector",Definitions!$1:$1,0)),0),MATCH("${stat}",Definitions!$1:$1,0))`
      )
    );
  }

  // Write formulas to E2:H{lastRow}
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `Sectors!E2:H${lastRow}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: formulas },
  });

  console.log(`Updated formulas in Sectors!E2:H${lastRow} to fully header-based format.`);
  console.log("All column references now use MATCH on header names — no column letters.");
  console.log("Verify in Google Sheets that the values are unchanged.");
}

main().catch((err) => {
  console.error("Error:", err.message);
  process.exit(1);
});

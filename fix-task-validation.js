// ---------------------------------------------------------------------------
// fix-task-validation.js — One-time script to:
//   Remove data validation (dropdowns) from the Task column in the Sectors sheet.
//
// Usage: node fix-task-validation.js
// ---------------------------------------------------------------------------

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");

// Find project root by walking up from __dirname until we find .env
function findProjectRoot(startDir) {
  let dir = startDir;
  while (dir !== path.dirname(dir)) {
    if (fs.existsSync(path.join(dir, ".env"))) return dir;
    dir = path.dirname(dir);
  }
  return startDir;
}
const PROJECT_ROOT = findProjectRoot(__dirname);

// Load .env
try {
  const envFile = fs.readFileSync(path.join(PROJECT_ROOT, ".env"), "utf8");
  for (const line of envFile.split("\n")) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eqIdx = trimmed.indexOf("=");
    if (eqIdx === -1) continue;
    const key = trimmed.slice(0, eqIdx).trim();
    const val = trimmed.slice(eqIdx + 1).trim();
    if (!process.env[key]) process.env[key] = val;
  }
} catch {}

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const CREDENTIALS_PATH = process.env.GOOGLE_CREDENTIALS_PATH || "./credentials.json";

async function main() {
  const creds = JSON.parse(fs.readFileSync(path.resolve(PROJECT_ROOT, CREDENTIALS_PATH), "utf8"));
  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  const sheets = google.sheets({ version: "v4", auth });

  // Step 1: Find the Task column and Sectors sheet ID
  console.log("Reading Sectors headers...");
  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sectors!1:1",
  });
  const headers = (headerRes.data.values || [[]])[0];
  const taskIdx = headers.indexOf("Task");

  if (taskIdx === -1) {
    console.error('ERROR: "Task" column not found in Sectors headers.');
    process.exit(1);
  }
  console.log(`  Task column: ${String.fromCharCode(65 + taskIdx)} (index ${taskIdx})`);

  // Get sheet ID for Sectors
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: "sheets.properties",
  });
  const sectorsSheet = meta.data.sheets.find(
    (s) => s.properties.title === "Sectors"
  );
  if (!sectorsSheet) {
    console.error("ERROR: Sectors sheet not found.");
    process.exit(1);
  }
  const sheetId = sectorsSheet.properties.sheetId;

  // Step 2: Clear data validation on the Task column (row 2 onward)
  console.log("Clearing data validation on Task column...");
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          setDataValidation: {
            range: {
              sheetId,
              startRowIndex: 1, // skip header
              startColumnIndex: taskIdx,
              endColumnIndex: taskIdx + 1,
            },
            // No "rule" property = clears existing validation
          },
        },
      ],
    },
  });

  console.log("\n========================================");
  console.log("Done! Data validation removed from Task column.");
  console.log("Cells should no longer show as dropdowns.");
  console.log("========================================");
}

main().catch((err) => {
  console.error("Failed:", err.message);
  process.exit(1);
});

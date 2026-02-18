// ---------------------------------------------------------------------------
// backfill-task.js — One-time script to:
//   Populate the "Task" column in the Sectors worksheet with matching
//   Task values from the Master Catalog, using Sector|Boss|Minion as the key.
//
// Prerequisites:
//   - "Task" column must already exist in both the Catalog and Sectors sheets
//   - Service account must have Editor access on both spreadsheets
//
// Usage: node backfill-task.js
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
  return startDir; // fallback to script dir
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
const CATALOG_SPREADSHEET_ID = process.env.CATALOG_SPREADSHEET_ID;
const CREDENTIALS_PATH = process.env.GOOGLE_CREDENTIALS_PATH || "./credentials.json";

async function main() {
  const creds = JSON.parse(fs.readFileSync(path.resolve(PROJECT_ROOT, CREDENTIALS_PATH), "utf8"));
  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  const sheets = google.sheets({ version: "v4", auth });

  // =========================================================================
  // Step 1: Read catalog data and build Task lookup
  // =========================================================================
  console.log("Reading Master Catalog...");
  const catalogRes = await sheets.spreadsheets.values.get({
    spreadsheetId: CATALOG_SPREADSHEET_ID,
    range: "Catalog!A:Z",
  });
  const catValues = catalogRes.data.values || [];
  const catHeaders = catValues[0] || [];
  const catSectorIdx = catHeaders.indexOf("Sector");
  const catBossIdx = catHeaders.indexOf("Boss");
  const catMinionIdx = catHeaders.indexOf("Minion");
  const catTaskIdx = catHeaders.indexOf("Task");

  if (catTaskIdx === -1) {
    console.error('ERROR: "Task" column not found in Catalog. Add it first.');
    process.exit(1);
  }

  // Build lookup: Sector|Boss|Minion -> Task
  const taskLookup = {};
  let catalogWithTask = 0;
  for (let i = 1; i < catValues.length; i++) {
    const row = catValues[i];
    const sector = row[catSectorIdx] || "";
    const boss = row[catBossIdx] || "";
    const minion = row[catMinionIdx] || "";
    const task = row[catTaskIdx] || "";
    if (sector && minion) {
      const key = `${sector}|${boss}|${minion}`;
      taskLookup[key] = task;
      if (task) catalogWithTask++;
    }
  }
  console.log(`  ${catValues.length - 1} catalog rows, ${catalogWithTask} with Task content.`);
  console.log(`  ${Object.keys(taskLookup).length} unique keys in lookup.`);

  // =========================================================================
  // Step 2: Read Sectors sheet and find Task column
  // =========================================================================
  console.log("\nReading Sectors worksheet...");
  const sectorsRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sectors",
  });
  const secValues = sectorsRes.data.values || [];
  const secHeaders = secValues[0] || [];
  const secSectorIdx = secHeaders.indexOf("Sector");
  const secBossIdx = secHeaders.indexOf("Boss");
  const secMinionIdx = secHeaders.indexOf("Minion");
  const secTaskIdx = secHeaders.indexOf("Task");

  if (secTaskIdx === -1) {
    console.error('ERROR: "Task" column not found in Sectors. Add it first.');
    process.exit(1);
  }

  const taskColLetter = String.fromCharCode(65 + secTaskIdx);
  console.log(`  ${secValues.length - 1} sector rows.`);
  console.log(`  Task column: ${taskColLetter} (index ${secTaskIdx})`);

  // =========================================================================
  // Step 3: Match and build updates
  // =========================================================================
  console.log("\nMatching Sectors rows to Catalog...");
  const updates = [];
  let matched = 0;
  let alreadyFilled = 0;
  let noMatch = 0;

  for (let i = 1; i < secValues.length; i++) {
    const row = secValues[i];
    const sector = row[secSectorIdx] || "";
    const boss = row[secBossIdx] || "";
    const minion = row[secMinionIdx] || "";
    const existingTask = row[secTaskIdx] || "";

    if (!sector || !minion) continue;

    const key = `${sector}|${boss}|${minion}`;
    const catalogTask = taskLookup[key];

    if (catalogTask === undefined) {
      noMatch++;
      continue;
    }

    if (!catalogTask) {
      // Catalog has this item but Task is empty — nothing to backfill
      continue;
    }

    if (existingTask) {
      alreadyFilled++;
      continue;
    }

    matched++;
    updates.push({
      range: `Sectors!${taskColLetter}${i + 1}`,
      values: [[catalogTask]],
    });
  }

  console.log(`  ${matched} rows to update with Task content.`);
  console.log(`  ${alreadyFilled} rows already had Task (skipped).`);
  console.log(`  ${noMatch} rows not found in catalog (skipped).`);

  // =========================================================================
  // Step 4: Write updates
  // =========================================================================
  if (updates.length === 0) {
    console.log("\nNo updates needed.");
    return;
  }

  console.log(`\nWriting ${updates.length} Task values to Sectors sheet...`);
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      valueInputOption: "RAW",
      data: updates,
    },
  });

  console.log("\n========================================");
  console.log("Backfill complete!");
  console.log("========================================");
  console.log(`Updated: ${matched} rows`);
  console.log(`Skipped (already filled): ${alreadyFilled}`);
  console.log(`Skipped (no catalog match): ${noMatch}`);
  console.log("========================================");
}

main().catch((err) => {
  console.error("Backfill failed:", err.message);
  process.exit(1);
});

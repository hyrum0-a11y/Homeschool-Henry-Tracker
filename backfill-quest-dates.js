// ---------------------------------------------------------------------------
// backfill-quest-dates.js — One-time script to:
//   1. Add "Date Added" and "Date Resolved" headers to the Quests sheet if missing
//   2. Backfill Date Added from the Sectors sheet's "Date Quest Added" column
//   3. Backfill Date Resolved from quest status + "Date Completed" column
//
// Prerequisites:
//   - Service account must have Editor access on the spreadsheet
//
// Usage: node backfill-quest-dates.js
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

  // =========================================================================
  // Step 1: Read Quests sheet and check/add headers
  // =========================================================================
  console.log("Reading Quests worksheet...");
  const questsRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Quests",
  });
  const questValues = questsRes.data.values || [];
  if (questValues.length < 1) {
    console.log("Quests sheet is empty. Nothing to backfill.");
    return;
  }

  const headers = questValues[0];
  let dateAddedCol = headers.indexOf("Date Added");
  let dateResolvedCol = headers.indexOf("Date Resolved");

  // Add missing headers
  if (dateAddedCol === -1 || dateResolvedCol === -1) {
    const headerUpdates = [];
    if (dateAddedCol === -1) {
      dateAddedCol = headers.length;
      headers.push("Date Added");
      console.log(`  Adding "Date Added" header at column ${String.fromCharCode(65 + dateAddedCol)}`);
    }
    if (dateResolvedCol === -1) {
      dateResolvedCol = headers.length;
      headers.push("Date Resolved");
      console.log(`  Adding "Date Resolved" header at column ${String.fromCharCode(65 + dateResolvedCol)}`);
    }
    // Write updated header row
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Quests!A1:${String.fromCharCode(64 + headers.length)}1`,
      valueInputOption: "RAW",
      requestBody: { values: [headers] },
    });
    console.log("  Headers updated.");
  } else {
    console.log("  Headers already have Date Added and Date Resolved.");
  }

  const idCol = headers.indexOf("Quest ID");
  const statusCol = headers.indexOf("Status");
  const bossCol = headers.indexOf("Boss");
  const minionCol = headers.indexOf("Minion");
  const sectorCol = headers.indexOf("Sector");
  const dateCompletedCol = headers.indexOf("Date Completed");

  console.log(`  ${questValues.length - 1} quest rows found.`);

  // =========================================================================
  // Step 2: Read Sectors sheet for Date Quest Added cross-reference
  // =========================================================================
  console.log("\nReading Sectors worksheet for Date Quest Added...");
  const sectorsRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sectors",
  });
  const secValues = sectorsRes.data.values || [];
  const secHeaders = secValues[0] || [];
  const secSectorIdx = secHeaders.indexOf("Sector");
  const secBossIdx = secHeaders.indexOf("Boss");
  const secMinionIdx = secHeaders.indexOf("Minion");
  const secDateAddedIdx = secHeaders.indexOf("Date Quest Added");

  // Build lookup: Sector|Boss|Minion -> Date Quest Added
  const dateAddedLookup = {};
  if (secDateAddedIdx >= 0) {
    for (let i = 1; i < secValues.length; i++) {
      const row = secValues[i];
      const sector = row[secSectorIdx] || "";
      const boss = row[secBossIdx] || "";
      const minion = row[secMinionIdx] || "";
      const dateAdded = (row[secDateAddedIdx] || "").slice(0, 10);
      if (sector && boss && minion && dateAdded) {
        dateAddedLookup[`${sector}|${boss}|${minion}`] = dateAdded;
      }
    }
    console.log(`  ${Object.keys(dateAddedLookup).length} Sectors rows with Date Quest Added.`);
  } else {
    console.log("  WARNING: Sectors sheet has no 'Date Quest Added' column.");
  }

  // =========================================================================
  // Step 3: Compute backfill values
  // =========================================================================
  console.log("\nComputing backfill values...");
  const updates = [];
  let filledAdded = 0;
  let filledResolved = 0;
  let skippedAdded = 0;
  let skippedResolved = 0;
  const dateAddedLetter = String.fromCharCode(65 + dateAddedCol);
  const dateResolvedLetter = String.fromCharCode(65 + dateResolvedCol);

  for (let i = 1; i < questValues.length; i++) {
    const row = questValues[i];
    const rowNum = i + 1; // 1-based for Sheets API
    const status = (row[statusCol] || "").trim();
    const sector = row[sectorCol] || "";
    const boss = row[bossCol] || "";
    const minion = row[minionCol] || "";
    const dateCompleted = row[dateCompletedCol] || "";
    const existingDateAdded = (row[dateAddedCol] || "").trim();
    const existingDateResolved = (row[dateResolvedCol] || "").trim();

    // --- Date Added ---
    if (!existingDateAdded) {
      const key = `${sector}|${boss}|${minion}`;
      const fromSectors = dateAddedLookup[key];
      if (fromSectors) {
        updates.push({
          range: `Quests!${dateAddedLetter}${rowNum}`,
          values: [[fromSectors]],
        });
        filledAdded++;
      } else {
        skippedAdded++;
      }
    }

    // --- Date Resolved ---
    if (!existingDateResolved) {
      let resolvedDate = "";

      if (status === "Approved" && dateCompleted) {
        // Use Date Completed (should be YYYY-MM-DD)
        resolvedDate = dateCompleted.slice(0, 10);
      } else if (status === "Abandoned" && dateCompleted) {
        // Format is "2025-01-15 | Abandoned by: Name" — take first 10 chars
        resolvedDate = dateCompleted.slice(0, 10);
      }
      // Active, Submitted, Rejected without dates — leave blank

      if (resolvedDate && /^\d{4}-\d{2}-\d{2}$/.test(resolvedDate)) {
        updates.push({
          range: `Quests!${dateResolvedLetter}${rowNum}`,
          values: [[resolvedDate]],
        });
        filledResolved++;
      } else if (status === "Approved" || status === "Abandoned") {
        skippedResolved++;
      }
    }
  }

  console.log(`  Date Added:    ${filledAdded} to fill, ${skippedAdded} no source data`);
  console.log(`  Date Resolved: ${filledResolved} to fill, ${skippedResolved} no valid date`);

  // =========================================================================
  // Step 4: Write updates
  // =========================================================================
  if (updates.length === 0) {
    console.log("\nNo updates needed.");
    return;
  }

  console.log(`\nWriting ${updates.length} cell updates...`);
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
  console.log(`Date Added filled:    ${filledAdded}`);
  console.log(`Date Resolved filled: ${filledResolved}`);
  console.log("========================================");
}

main().catch((err) => {
  console.error("Backfill failed:", err.message);
  process.exit(1);
});

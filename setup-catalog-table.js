// ---------------------------------------------------------------------------
// setup-catalog-table.js — One-time script to:
//   1. Add "Subject" column to Master Catalog (boss-level mapping)
//   2. Format as a Google Sheets table (filters, banding, frozen header)
// ---------------------------------------------------------------------------

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");

// Load .env
try {
  const envFile = fs.readFileSync(path.join(__dirname, ".env"), "utf8");
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

const CATALOG_ID = process.env.CATALOG_SPREADSHEET_ID;
const CREDENTIALS_PATH = process.env.GOOGLE_CREDENTIALS_PATH || "./credentials.json";

// Boss → Subject mapping
const BOSS_SUBJECT_MAP = {
  "Digital Artist": "Visual Arts",
  "Pianist1": "Music",
  "Spaniard1": "Spanish",
  "Storyteller": "Creative Writing",
  "The COMPOSER": "English",
  "The Grammarian": "English",
  "The Texter": "Communication Skills",
  "The Cartographer": "Geography",
  "The Chemist": "Chemistry",
  "The Physicist": "Physics",
  "The Behavioralist": "Psychology",
  "The Cognitive Scientist": "Psychology",
  "The Governor": "Government",
  "The Socialite": "Sociology",
  "Miser Confused": "Critical Thinking",
  "Mister Calculus": "Calculus",
  "The Algebraist": "Algebra",
  "The Arithmetist": "Arithmetic",
  "The Triggonom": "Trigonometry",
  "The Vector": "Linear Algebra",
  "The Architect": "Architecture",
  "The Electrical Engineer": "Electrical Engineering",
  "The Engineer": "Engineering",
  "The Financier": "Financial Literacy",
  "Hygiene": "Personal Care",
  "The Cardiologist": "Health & Fitness",
  "The Dentist": "Personal Care",
};

async function main() {
  const creds = JSON.parse(fs.readFileSync(path.resolve(__dirname, CREDENTIALS_PATH), "utf8"));
  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  const sheets = google.sheets({ version: "v4", auth });

  // =========================================================================
  // Step 1: Read current catalog data
  // =========================================================================
  console.log("Reading catalog data...");
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: CATALOG_ID,
    range: "Catalog",
  });
  const allRows = res.data.values || [];
  const headers = allRows[0];
  const dataRows = allRows.slice(1);
  console.log(`  ${dataRows.length} rows, ${headers.length} columns`);

  const bossIdx = headers.indexOf("Boss");
  if (bossIdx === -1) throw new Error("Boss column not found");

  // Check if Subject column already exists
  if (headers.includes("Subject")) {
    console.log("  Subject column already exists — updating values only.");
  }

  // =========================================================================
  // Step 2: Build new data with Subject column inserted after Sector
  // =========================================================================
  const sectorIdx = headers.indexOf("Sector");
  const insertIdx = sectorIdx + 1; // Insert Subject right after Sector

  let newHeaders;
  let newDataRows;

  if (headers.includes("Subject")) {
    // Update existing Subject column
    const subjectIdx = headers.indexOf("Subject");
    newHeaders = headers;
    newDataRows = dataRows.map((row) => {
      const boss = row[bossIdx] || "";
      const newRow = [...row];
      // Pad row if needed
      while (newRow.length < headers.length) newRow.push("");
      newRow[subjectIdx] = BOSS_SUBJECT_MAP[boss] || "General";
      return newRow;
    });
  } else {
    // Insert new Subject column
    newHeaders = [
      ...headers.slice(0, insertIdx),
      "Subject",
      ...headers.slice(insertIdx),
    ];
    newDataRows = dataRows.map((row) => {
      const boss = row[bossIdx] || "";
      const subject = BOSS_SUBJECT_MAP[boss] || "General";
      return [
        ...row.slice(0, insertIdx),
        subject,
        ...row.slice(insertIdx),
      ];
    });
  }

  // =========================================================================
  // Step 3: Write updated data back
  // =========================================================================
  console.log("Writing updated data with Subject column...");

  // Clear existing data first
  await sheets.spreadsheets.values.clear({
    spreadsheetId: CATALOG_ID,
    range: "Catalog",
  });

  // Write new data
  await sheets.spreadsheets.values.update({
    spreadsheetId: CATALOG_ID,
    range: "Catalog!A1",
    valueInputOption: "RAW",
    requestBody: {
      values: [newHeaders, ...newDataRows],
    },
  });
  console.log(`  Wrote ${newDataRows.length} rows with ${newHeaders.length} columns.`);

  // =========================================================================
  // Step 4: Get sheet metadata for formatting
  // =========================================================================
  console.log("Formatting as table...");
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: CATALOG_ID,
    fields: "sheets.properties",
  });
  const sheetId = meta.data.sheets[0].properties.sheetId;
  const totalRows = newDataRows.length + 1; // +1 for header
  const totalCols = newHeaders.length;

  // =========================================================================
  // Step 5: Apply table formatting
  // =========================================================================
  const requests = [
    // Bold header with dark background
    {
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: 0,
          endRowIndex: 1,
          startColumnIndex: 0,
          endColumnIndex: totalCols,
        },
        cell: {
          userEnteredFormat: {
            textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
            backgroundColor: { red: 0.2, green: 0.2, blue: 0.25 },
          },
        },
        fields: "userEnteredFormat(textFormat,backgroundColor)",
      },
    },
    // Freeze header row
    {
      updateSheetProperties: {
        properties: {
          sheetId,
          gridProperties: { frozenRowCount: 1 },
        },
        fields: "gridProperties.frozenRowCount",
      },
    },
    // Add filter
    {
      setBasicFilter: {
        filter: {
          range: {
            sheetId,
            startRowIndex: 0,
            endRowIndex: totalRows,
            startColumnIndex: 0,
            endColumnIndex: totalCols,
          },
        },
      },
    },
    // Banded rows
    {
      addBanding: {
        bandedRange: {
          range: {
            sheetId,
            startRowIndex: 0,
            endRowIndex: totalRows,
            startColumnIndex: 0,
            endColumnIndex: totalCols,
          },
          rowProperties: {
            headerColor: { red: 0.2, green: 0.2, blue: 0.25 },
            firstBandColor: { red: 0.95, green: 0.95, blue: 0.97 },
            secondBandColor: { red: 1, green: 1, blue: 1 },
          },
        },
      },
    },
    // Auto-resize columns
    {
      autoResizeDimensions: {
        dimensions: {
          sheetId,
          dimension: "COLUMNS",
          startIndex: 0,
          endIndex: totalCols,
        },
      },
    },
  ];

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: CATALOG_ID,
    requestBody: { requests },
  });

  console.log("\n========================================");
  console.log("Setup complete!");
  console.log("========================================");
  console.log(`Catalog columns: ${newHeaders.join(", ")}`);
  console.log(`Total rows: ${totalRows} (including header)`);
  console.log("Features: Subject column, filters, banded rows, frozen header");
  console.log("========================================");
}

main().catch((err) => {
  console.error("Setup failed:", err.message);
  process.exit(1);
});

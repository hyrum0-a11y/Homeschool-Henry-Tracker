// ---------------------------------------------------------------------------
// setup-drive.js — One-time script to:
//   1. Create "Homeschool Tracker" folder in Google Drive
//   2. Share the folder with the spreadsheet owner
//   3. Move Henry's existing spreadsheet into the folder
//   4. Create a "Master Catalog" spreadsheet in the folder
//   5. Copy Sectors data into the catalog with an "In Henry's Sheet" flag
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

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const CREDENTIALS_PATH = process.env.GOOGLE_CREDENTIALS_PATH || "./credentials.json";
const OWNER_EMAIL = "hyrum.0@gmail.com";

async function main() {
  const creds = JSON.parse(fs.readFileSync(path.resolve(__dirname, CREDENTIALS_PATH), "utf8"));
  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
    ],
  });

  const drive = google.drive({ version: "v3", auth });
  const sheets = google.sheets({ version: "v4", auth });

  // =========================================================================
  // Step 1: Create "Homeschool Tracker" folder
  // =========================================================================
  console.log("Creating 'Homeschool Tracker' folder...");
  const folder = await drive.files.create({
    requestBody: {
      name: "Homeschool Tracker",
      mimeType: "application/vnd.google-apps.folder",
    },
    fields: "id, name",
  });
  const folderId = folder.data.id;
  console.log(`  Folder created: ${folder.data.name} (${folderId})`);

  // =========================================================================
  // Step 2: Share folder with owner
  // =========================================================================
  console.log(`Sharing folder with ${OWNER_EMAIL}...`);
  await drive.permissions.create({
    fileId: folderId,
    requestBody: {
      type: "user",
      role: "writer",
      emailAddress: OWNER_EMAIL,
    },
    sendNotificationEmail: false,
  });
  console.log("  Shared successfully.");

  // =========================================================================
  // Step 3: Move Henry's spreadsheet into the folder
  // =========================================================================
  console.log("Moving Henry's spreadsheet into folder...");
  try {
    const existing = await drive.files.get({
      fileId: SPREADSHEET_ID,
      fields: "parents",
    });
    const previousParents = (existing.data.parents || []).join(",");
    await drive.files.update({
      fileId: SPREADSHEET_ID,
      addParents: folderId,
      removeParents: previousParents,
      fields: "id, parents",
    });
    console.log("  Moved successfully.");
  } catch (err) {
    console.log(`  Could not move (service account may not have permission): ${err.message}`);
    console.log("  You can manually move it into the 'Homeschool Tracker' folder in Google Drive.");
  }

  // =========================================================================
  // Step 4: Read all data from Henry's Sectors sheet
  // =========================================================================
  console.log("Reading Henry's Sectors data...");
  const sectorsRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sectors",
  });
  const allRows = sectorsRes.data.values || [];
  const headers = allRows[0] || [];
  const dataRows = allRows.slice(1);
  console.log(`  Found ${dataRows.length} rows.`);

  // Column indices in source
  const colIdx = {};
  headers.forEach((h, i) => { colIdx[h] = i; });

  // =========================================================================
  // Step 5: Create Master Catalog spreadsheet in folder
  // =========================================================================
  console.log("Creating Master Catalog spreadsheet...");
  const catalogFile = await drive.files.create({
    requestBody: {
      name: "Master Catalog",
      mimeType: "application/vnd.google-apps.spreadsheet",
      parents: [folderId],
    },
    fields: "id, name",
  });
  const catalogId = catalogFile.data.id;
  console.log(`  Created: ${catalogFile.data.name} (${catalogId})`);

  // Share catalog with owner
  await drive.permissions.create({
    fileId: catalogId,
    requestBody: {
      type: "user",
      role: "writer",
      emailAddress: OWNER_EMAIL,
    },
    sendNotificationEmail: false,
  });
  console.log(`  Shared with ${OWNER_EMAIL}.`);

  // =========================================================================
  // Step 6: Build catalog data and write to Master Catalog
  // =========================================================================
  console.log("Populating Master Catalog...");

  // Catalog columns
  const catalogHeaders = [
    "Sector",
    "Subject",
    "Boss",
    "Minion",
    "Task",
    "Status",
    "Impact(1-3)",
    "Locked for what?",
    "Suggested Proof Method",
    "Grade Level",
    "In Henry's Sheet",
    "Source",
  ];

  const catalogRows = dataRows.map((row) => {
    const get = (col) => (colIdx[col] !== undefined ? row[colIdx[col]] || "" : "");
    return [
      get("Sector"),
      get("Subject"),
      get("Boss"),
      get("Minion"),
      get("Task"),
      get("Status"),
      get("Impact(1-3)"),
      get("Locked for what?"),
      "",  // Suggested Proof Method — empty for now
      "High School",  // Grade Level — default per user's plan
      "Yes",  // In Henry's Sheet — all copied rows are from Henry's sheet
      "",  // Source
    ];
  });

  // Rename the default "Sheet1" tab to "Catalog"
  const catalogMeta = await sheets.spreadsheets.get({
    spreadsheetId: catalogId,
    fields: "sheets.properties",
  });
  const defaultSheetId = catalogMeta.data.sheets[0].properties.sheetId;

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: catalogId,
    requestBody: {
      requests: [
        {
          updateSheetProperties: {
            properties: { sheetId: defaultSheetId, title: "Catalog" },
            fields: "title",
          },
        },
      ],
    },
  });

  // Write headers + data
  await sheets.spreadsheets.values.update({
    spreadsheetId: catalogId,
    range: "Catalog!A1",
    valueInputOption: "RAW",
    requestBody: {
      values: [catalogHeaders, ...catalogRows],
    },
  });
  console.log(`  Wrote ${catalogRows.length} rows to Master Catalog.`);

  // =========================================================================
  // Step 7: Format the header row (bold, freeze)
  // =========================================================================
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: catalogId,
    requestBody: {
      requests: [
        {
          repeatCell: {
            range: { sheetId: defaultSheetId, startRowIndex: 0, endRowIndex: 1 },
            cell: {
              userEnteredFormat: {
                textFormat: { bold: true },
                backgroundColor: { red: 0.2, green: 0.2, blue: 0.2 },
              },
            },
            fields: "userEnteredFormat(textFormat,backgroundColor)",
          },
        },
        {
          updateSheetProperties: {
            properties: {
              sheetId: defaultSheetId,
              gridProperties: { frozenRowCount: 1 },
            },
            fields: "gridProperties.frozenRowCount",
          },
        },
      ],
    },
  });

  // =========================================================================
  // Done — print summary
  // =========================================================================
  console.log("\n========================================");
  console.log("Setup complete!");
  console.log("========================================");
  console.log(`Folder:         Homeschool Tracker (${folderId})`);
  console.log(`Master Catalog: ${catalogId}`);
  console.log(`Henry's Sheet:  ${SPREADSHEET_ID}`);
  console.log(`\nAdd this to your .env file:`);
  console.log(`CATALOG_SPREADSHEET_ID=${catalogId}`);
  console.log(`DRIVE_FOLDER_ID=${folderId}`);
  console.log("========================================");
}

main().catch((err) => {
  console.error("Setup failed:", err.message);
  process.exit(1);
});

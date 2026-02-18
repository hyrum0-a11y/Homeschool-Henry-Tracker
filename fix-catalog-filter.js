// One-time fix: expand catalog table filter and banding to cover all rows
const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");

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

async function main() {
  const creds = JSON.parse(fs.readFileSync(path.resolve(__dirname, CREDENTIALS_PATH), "utf8"));
  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  const sheets = google.sheets({ version: "v4", auth });

  const dataRes = await sheets.spreadsheets.values.get({ spreadsheetId: CATALOG_ID, range: "Catalog!A:Z" });
  const totalRows = dataRes.data.values.length;
  const totalCols = dataRes.data.values[0].length;
  console.log("Data size:", totalRows, "rows x", totalCols, "cols");

  const meta = await sheets.spreadsheets.get({
    spreadsheetId: CATALOG_ID,
    fields: "sheets.properties,sheets.basicFilter,sheets.bandedRanges",
  });
  const sheet = meta.data.sheets[0];
  const sheetId = sheet.properties.sheetId;

  console.log("Current filter range:", sheet.basicFilter ? JSON.stringify(sheet.basicFilter.range) : "none");

  const requests = [];
  if (sheet.basicFilter) {
    requests.push({ clearBasicFilter: { sheetId } });
  }
  requests.push({
    setBasicFilter: {
      filter: {
        range: { sheetId, startRowIndex: 0, endRowIndex: totalRows, startColumnIndex: 0, endColumnIndex: totalCols },
      },
    },
  });

  if (sheet.bandedRanges && sheet.bandedRanges.length > 0) {
    const bandId = sheet.bandedRanges[0].bandedRangeId;
    requests.push({
      updateBanding: {
        bandedRange: {
          bandedRangeId: bandId,
          range: { sheetId, startRowIndex: 0, endRowIndex: totalRows, startColumnIndex: 0, endColumnIndex: totalCols },
          rowProperties: sheet.bandedRanges[0].rowProperties,
        },
        fields: "range",
      },
    });
  }

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: CATALOG_ID,
    requestBody: { requests },
  });

  console.log("Done! Filter and banding now cover all", totalRows, "rows.");
}

main().catch((err) => {
  console.error("Error:", err.message);
  process.exit(1);
});

/**
 * AI-Powered Lesson Photo Import
 *
 * Analyzes photos of completed lessons using Claude vision,
 * maps them to the Sectors sheet structure, and batch-imports.
 *
 * Usage:
 *   1. Drop photos into imports/
 *   2. Run: node import-lessons.js  (or npm run import)
 *   3. Review the AI classifications
 *   4. Approve to write to Google Sheets
 */

const Anthropic = require("@anthropic-ai/sdk");
const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");
const readline = require("readline");

// ---------------------------------------------------------------------------
// Load .env manually (same pattern as server.js)
// ---------------------------------------------------------------------------
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
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;

const IMPORT_DIR = path.join(__dirname, "imports");
const DONE_DIR = path.join(IMPORT_DIR, "done");
const EXTENSIONS = new Set([".jpg", ".jpeg", ".png", ".webp"]);

// ---------------------------------------------------------------------------
// Preflight checks
// ---------------------------------------------------------------------------
function preflight() {
  const errors = [];
  if (!SPREADSHEET_ID) errors.push("SPREADSHEET_ID not set in .env");
  if (!ANTHROPIC_API_KEY) errors.push("ANTHROPIC_API_KEY not set in .env — get one at https://console.anthropic.com/");
  if (errors.length) {
    console.error("\n  SETUP REQUIRED:\n");
    errors.forEach((e) => console.error("  - " + e));
    console.error("");
    process.exit(1);
  }
}

// ---------------------------------------------------------------------------
// Google Sheets auth (same pattern as server.js / fix-formulas.js)
// ---------------------------------------------------------------------------
async function getSheets() {
  const creds = JSON.parse(fs.readFileSync(path.resolve(__dirname, CREDENTIALS_PATH), "utf8"));
  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

// ---------------------------------------------------------------------------
// Convert raw sheet values to array of objects keyed by header
// ---------------------------------------------------------------------------
function parseTable(values) {
  if (!values || values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).map((row) => {
    const obj = {};
    for (let i = 0; i < headers.length; i++) {
      if (headers[i]) obj[headers[i]] = row[i] ?? "";
    }
    return obj;
  });
}

// ---------------------------------------------------------------------------
// Fetch existing context from Google Sheets
// ---------------------------------------------------------------------------
async function fetchContext(sheets) {
  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["Sectors", "Definitions"],
  });
  const [sectorsRaw, defsRaw] = res.data.valueRanges;
  const sectors = parseTable(sectorsRaw.values);
  const definitions = parseTable(defsRaw.values);

  // Read actual header order from Sectors sheet
  const headers = sectorsRaw.values ? sectorsRaw.values[0] : [];

  // Group existing bosses by sector
  const bossMap = {};
  const minionSet = new Set();
  for (const row of sectors) {
    const s = row["Sector"], b = row["Boss"], m = row["Minion"];
    if (!s || !b) continue;
    if (!bossMap[s]) bossMap[s] = new Set();
    bossMap[s].add(b);
    if (m) minionSet.add(m.toLowerCase());
  }

  // Valid sector names from Definitions
  const validSectors = definitions.filter((d) => d["Sector"]).map((d) => d["Sector"]);

  return { sectors, definitions, bossMap, minionSet, validSectors, headers };
}

// ---------------------------------------------------------------------------
// Build AI prompt with existing context
// ---------------------------------------------------------------------------
function buildPrompt(context) {
  const { bossMap, validSectors } = context;

  let sectorList = "";
  for (const sector of validSectors) {
    const bosses = bossMap[sector] ? Array.from(bossMap[sector]).join(", ") : "(no bosses yet)";
    sectorList += `  - ${sector}: Bosses = [${bosses}]\n`;
  }

  const system = `You are a homeschool lesson classifier for a gamified learning tracker.
The student is "Henry". His learning is tracked in a spreadsheet with "Sectors" (subject areas),
"Bosses" (topics within a sector), and "Minions" (individual lessons/skills).

When a lesson is completed, its status is "Enslaved".

AVAILABLE SECTORS AND EXISTING BOSSES:
${sectorList}
RULES:
1. You MUST use one of the existing sectors listed above. Never invent a new sector.
2. Prefer matching to an existing Boss when the lesson fits. Only suggest a new Boss if nothing existing is close.
3. New Boss names should follow the pattern of existing ones (creative/thematic names, often "The [Noun]" format).
4. Name the Minion clearly and concisely (2-6 words). It should describe the specific skill or lesson topic.
5. Impact (1-3): 1 = simple/quick lesson, 2 = moderate lesson, 3 = complex/deep lesson.
6. Return ONLY valid JSON. No markdown fencing, no explanation outside the JSON.`;

  const userPrefix = `Analyze this photo of a completed homeschool lesson.
Identify what subject/topic it covers and return a JSON object:
{
  "minion": "Lesson Name",
  "boss": "Topic Name",
  "sector": "Sector Name",
  "impact": 2,
  "confidence": "high",
  "reasoning": "Brief explanation of classification"
}`;

  return { system, userPrefix };
}

// ---------------------------------------------------------------------------
// Analyze a single image with Claude vision
// ---------------------------------------------------------------------------
async function analyzeImage(client, imagePath, prompt) {
  const imageData = fs.readFileSync(imagePath);
  const base64 = imageData.toString("base64");
  const ext = path.extname(imagePath).toLowerCase();
  const mediaType = { ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp" }[ext] || "image/jpeg";

  const response = await client.messages.create({
    model: "claude-sonnet-4-20250514",
    max_tokens: 1024,
    system: prompt.system,
    messages: [
      {
        role: "user",
        content: [
          { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } },
          { type: "text", text: prompt.userPrefix },
        ],
      },
    ],
  });

  const text = response.content[0].text;

  // Extract JSON even if wrapped in text
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) throw new Error("No JSON found in response");
  const parsed = JSON.parse(jsonMatch[0]);

  return {
    result: parsed,
    usage: { input: response.usage.input_tokens, output: response.usage.output_tokens },
    file: path.basename(imagePath),
  };
}

// ---------------------------------------------------------------------------
// Display preview table
// ---------------------------------------------------------------------------
function displayPreview(results, minionSet) {
  console.log("\n" + "=".repeat(100));
  console.log("  IMPORT PREVIEW");
  console.log("=".repeat(100));
  console.log(
    "  #  | File                 | Sector          | Boss               | Minion                     | Imp | Conf"
  );
  console.log("-".repeat(100));

  for (let i = 0; i < results.length; i++) {
    const r = results[i];
    if (r.error) {
      console.log(`  ${String(i + 1).padStart(2)} | ${r.file.substring(0, 20).padEnd(20)} | ** ERROR: ${r.error} **`);
      continue;
    }
    const idx = String(i + 1).padStart(2);
    const file = r.file.substring(0, 20).padEnd(20);
    const sector = (r.result.sector || "?").substring(0, 15).padEnd(15);
    const boss = (r.result.boss || "?").substring(0, 18).padEnd(18);
    const minion = (r.result.minion || "?").substring(0, 26).padEnd(26);
    const impact = String(r.result.impact || "?").padStart(3);
    const conf = (r.result.confidence || "?").padEnd(6);
    const dupe = minionSet.has((r.result.minion || "").toLowerCase()) ? " [DUP]" : "";
    console.log(`  ${idx} | ${file} | ${sector} | ${boss} | ${minion} | ${impact} | ${conf}${dupe}`);
  }

  console.log("=".repeat(100));

  // Token usage and cost
  const totalInput = results.reduce((s, r) => s + (r.usage?.input || 0), 0);
  const totalOutput = results.reduce((s, r) => s + (r.usage?.output || 0), 0);
  const cost = (totalInput * 3 + totalOutput * 15) / 1_000_000;
  console.log(`  Tokens: ${totalInput} in / ${totalOutput} out | Est. cost: $${cost.toFixed(4)}`);

  const errorCount = results.filter((r) => r.error).length;
  if (errorCount) console.log(`  Errors: ${errorCount} image(s) failed analysis (will be skipped)`);
  console.log("");
}

// ---------------------------------------------------------------------------
// Interactive review loop
// ---------------------------------------------------------------------------
async function interactiveReview(results) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  const ask = (q) => new Promise((resolve) => rl.question(q, resolve));

  // Filter out errors
  const valid = results.map((r, i) => ({ ...r, idx: i, approved: !r.error }));

  console.log("  Commands: [a]pprove all  |  s <#> skip  |  e <#> edit  |  q quit\n");

  while (true) {
    const input = (await ask("  > ")).trim().toLowerCase();

    if (input === "a") {
      console.log("  Approved all.");
      break;
    }
    if (input === "q") {
      console.log("  Aborted.");
      rl.close();
      process.exit(0);
    }

    const parts = input.split(/\s+/);
    const cmd = parts[0];
    const num = parseInt(parts[1]) - 1;

    if (cmd === "s" && num >= 0 && num < valid.length) {
      valid[num].approved = false;
      console.log(`  Skipped #${num + 1}: ${valid[num].file}`);
    } else if (cmd === "e" && num >= 0 && num < valid.length && valid[num].result) {
      const r = valid[num].result;
      console.log(`  Editing #${num + 1}: ${valid[num].file}`);
      const newMinion = await ask(`    Minion [${r.minion}]: `);
      const newBoss = await ask(`    Boss [${r.boss}]: `);
      const newSector = await ask(`    Sector [${r.sector}]: `);
      const newImpact = await ask(`    Impact [${r.impact}]: `);
      if (newMinion.trim()) r.minion = newMinion.trim();
      if (newBoss.trim()) r.boss = newBoss.trim();
      if (newSector.trim()) r.sector = newSector.trim();
      if (newImpact.trim()) r.impact = parseInt(newImpact) || r.impact;
      console.log(`    Updated #${num + 1}.`);
    } else {
      console.log("  Unknown command. Use: a, s <#>, e <#>, q");
    }
  }

  rl.close();
  return valid.filter((r) => r.approved && r.result);
}

// ---------------------------------------------------------------------------
// Build a row array matching the Sectors sheet header order
// ---------------------------------------------------------------------------
function buildRow(headers, data) {
  const statNames = ["INTELLIGENCE", "STAMINA", "TEMPO", "REPUTATION"];
  const statFormula = (stat) =>
    `=INDEX($A:$Z,ROW(),MATCH("Impact(1-3)",$1:$1,0))*INDEX(Definitions!$A:$Z,MATCH(INDEX($A:$Z,ROW(),MATCH("Sector",$1:$1,0)),INDEX(Definitions!$A:$Z,,MATCH("Sector",Definitions!$1:$1,0)),0),MATCH("${stat}",Definitions!$1:$1,0))`;

  const valueMap = {
    Boss: data.boss,
    Minion: data.minion,
    Sector: data.sector,
    Subject: data.subject || "",
    Task: data.task || "",
    Status: "Enslaved",
    "Impact(1-3)": data.impact,
    "Survival Mode Required": "",
    "Quest Status": "",
  };
  for (const stat of statNames) {
    valueMap[stat] = statFormula(stat);
  }

  return headers.map((h) => valueMap[h] ?? "");
}

// ---------------------------------------------------------------------------
// Batch-append approved rows to Sectors sheet
// ---------------------------------------------------------------------------
async function appendRows(sheets, headers, approved) {
  const rows = approved.map((r) => buildRow(headers, r.result));

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sectors!A:Z",
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: rows },
  });

  console.log(`\n  Appended ${rows.length} row(s) to Sectors sheet.`);
}

// ---------------------------------------------------------------------------
// Move processed images to imports/done/
// ---------------------------------------------------------------------------
function moveProcessed(approved) {
  for (const r of approved) {
    const src = path.join(IMPORT_DIR, r.file);
    const dest = path.join(DONE_DIR, r.file);
    if (fs.existsSync(dest)) {
      const ext = path.extname(r.file);
      const base = path.basename(r.file, ext);
      fs.renameSync(src, path.join(DONE_DIR, `${base}_${Date.now()}${ext}`));
    } else {
      fs.renameSync(src, dest);
    }
  }
  console.log(`  Moved ${approved.length} image(s) to imports/done/`);
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
async function main() {
  preflight();

  // Ensure directories exist
  fs.mkdirSync(IMPORT_DIR, { recursive: true });
  fs.mkdirSync(DONE_DIR, { recursive: true });

  // Scan for images
  const imageFiles = fs
    .readdirSync(IMPORT_DIR)
    .filter((f) => EXTENSIONS.has(path.extname(f).toLowerCase()) && fs.statSync(path.join(IMPORT_DIR, f)).isFile());

  if (imageFiles.length === 0) {
    console.log("\n  No images found in imports/");
    console.log("  Drop .jpg, .png, or .webp files there and run again.\n");
    return;
  }

  console.log(`\n  Found ${imageFiles.length} image(s) in imports/\n`);

  // Initialize clients
  const sheets = await getSheets();
  const anthropic = new Anthropic({ apiKey: ANTHROPIC_API_KEY });

  // Fetch context from sheets
  console.log("  Fetching existing data from Google Sheets...");
  const context = await fetchContext(sheets);
  console.log(`  Found ${context.validSectors.length} sectors, ${Object.values(context.bossMap).reduce((s, b) => s + b.size, 0)} bosses\n`);

  // Build AI prompt
  const prompt = buildPrompt(context);

  // Analyze each image
  const results = [];
  for (let i = 0; i < imageFiles.length; i++) {
    const file = imageFiles[i];
    const filePath = path.join(IMPORT_DIR, file);
    const sizeKB = (fs.statSync(filePath).size / 1024).toFixed(0);
    process.stdout.write(`  [${i + 1}/${imageFiles.length}] Analyzing ${file} (${sizeKB}KB)...`);

    try {
      const result = await analyzeImage(anthropic, filePath, prompt);
      results.push(result);
      console.log(` ${result.result.confidence} - ${result.result.minion}`);
    } catch (err) {
      results.push({ file, error: err.message, usage: {} });
      console.log(` ERROR: ${err.message}`);
    }

    // Small delay between API calls to avoid rate limits
    if (i < imageFiles.length - 1) {
      await new Promise((r) => setTimeout(r, 500));
    }
  }

  // Display preview
  displayPreview(results, context.minionSet);

  // Interactive review
  const approved = await interactiveReview(results);

  if (approved.length === 0) {
    console.log("\n  No items to import.\n");
    return;
  }

  // Append to sheet
  console.log(`\n  Writing ${approved.length} row(s) to Sectors sheet...`);
  await appendRows(sheets, context.headers, approved);

  // Move processed images
  moveProcessed(approved);

  console.log(`\n  Done! ${approved.length} minion(s) enslaved.\n`);
}

main().catch((err) => {
  console.error("\n  Fatal error:", err.message);
  process.exit(1);
});

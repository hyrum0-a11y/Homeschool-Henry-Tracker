const express = require("express");
const cookieParser = require("cookie-parser");
const { google } = require("googleapis");
const Anthropic = require("@anthropic-ai/sdk");
const multer = require("multer");
const fs = require("fs");
const path = require("path");

// ---------------------------------------------------------------------------
// Find project root by walking up from __dirname until we find .env
// ---------------------------------------------------------------------------
function findProjectRoot(startDir) {
  let dir = startDir;
  while (dir !== path.dirname(dir)) {
    if (fs.existsSync(path.join(dir, ".env"))) return dir;
    dir = path.dirname(dir);
  }
  return startDir; // fallback to script dir
}
const PROJECT_ROOT = findProjectRoot(__dirname);

// ---------------------------------------------------------------------------
// Load .env manually (no dotenv dependency needed)
// ---------------------------------------------------------------------------
try {
  const envPath = path.join(PROJECT_ROOT, ".env");
  const envFile = fs.readFileSync(envPath, "utf8");
  for (const line of envFile.split("\n")) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eqIdx = trimmed.indexOf("=");
    if (eqIdx === -1) continue;
    const key = trimmed.slice(0, eqIdx).trim();
    const val = trimmed.slice(eqIdx + 1).trim();
    if (!process.env[key]) process.env[key] = val;
  }
} catch {
  // .env is optional if env vars are set externally
}

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const CREDENTIALS_PATH = process.env.GOOGLE_CREDENTIALS_PATH || "./credentials.json";
const PORT = parseInt(process.env.PORT, 10) || 3000;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const IMPORT_DIR = path.join(PROJECT_ROOT, "imports");
const DONE_DIR = path.join(IMPORT_DIR, "done");

// ---------------------------------------------------------------------------
// Google Sheets auth (readwrite for fix-formulas support)
// ---------------------------------------------------------------------------
async function getSheets() {
  const creds = JSON.parse(fs.readFileSync(path.resolve(PROJECT_ROOT, CREDENTIALS_PATH), "utf8"));
  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

// ---------------------------------------------------------------------------
// Convert raw sheet values (2D array) into array of objects keyed by header
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
// Badge Definitions (conditions checked against live game state)
// ---------------------------------------------------------------------------
const STAT_TIERS = ["bronze", "copper", "silver", "gold", "platinum"];

function getStatTierIndex(levelName) {
  if (!levelName) return 0;
  const lower = levelName.toLowerCase();
  for (let i = STAT_TIERS.length - 1; i >= 0; i--) {
    if (lower.startsWith(STAT_TIERS[i])) return i;
  }
  return 0;
}

function getBossBadgeDef(sector, bossName) {
  return {
    category: "boss",
    name: bossName.toUpperCase(),
    icon: "\u{1F480}",
    description: `Defeated ${bossName} in ${sector} (all minions enslaved)`,
    color: "#ff0044",
  };
}

const BADGE_DEFINITIONS = {
  // Sector badges — one per sector, all bosses must be complete
  "sector:COMMUNICATION": { category: "sector", name: "COMM OVERLORD",   icon: "\u{1F4E1}", description: "Enslaved every minion in COMMUNICATION", color: "#ff00ff" },
  "sector:CORPOREAL":     { category: "sector", name: "MATTER SOVEREIGN", icon: "\u{269B}",  description: "Enslaved every minion in CORPOREAL",     color: "#00f2ff" },
  "sector:HUMANITY":      { category: "sector", name: "MIND ARCHITECT",   icon: "\u{1F9E0}", description: "Enslaved every minion in HUMANITY",      color: "#ffea00" },
  "sector:LOGIC":         { category: "sector", name: "LOGIC EMPEROR",    icon: "\u{1F9EE}", description: "Enslaved every minion in LOGIC",         color: "#00ff9d" },
  "sector:SYSTEMS":       { category: "sector", name: "SYS ADMIN",        icon: "\u{2699}",  description: "Enslaved every minion in SYSTEMS",       color: "#ff8800" },
  "sector:VITALITY":      { category: "sector", name: "VITALITY PRIME",   icon: "\u{1F4AA}", description: "Enslaved every minion in VITALITY",      color: "#00ff9d" },

  // Stat badges — 4 stats × 3 tiers (Silver, Gold, Platinum)
  "stat:intel:silver":     { category: "stat", name: "INTEL SILVER",     icon: "\u{1F4A1}", description: "Reached Silver tier in Intel",     color: "#c0c0c0" },
  "stat:intel:gold":       { category: "stat", name: "INTEL GOLD",       icon: "\u{1F4A1}", description: "Reached Gold tier in Intel",       color: "#ffd700" },
  "stat:intel:platinum":   { category: "stat", name: "INTEL PLATINUM",   icon: "\u{1F4A1}", description: "Reached Platinum tier in Intel",   color: "#e5e4e2" },
  "stat:stamina:silver":   { category: "stat", name: "STAMINA SILVER",   icon: "\u{1F6E1}", description: "Reached Silver tier in Stamina",   color: "#c0c0c0" },
  "stat:stamina:gold":     { category: "stat", name: "STAMINA GOLD",     icon: "\u{1F6E1}", description: "Reached Gold tier in Stamina",     color: "#ffd700" },
  "stat:stamina:platinum": { category: "stat", name: "STAMINA PLATINUM", icon: "\u{1F6E1}", description: "Reached Platinum tier in Stamina", color: "#e5e4e2" },
  "stat:tempo:silver":     { category: "stat", name: "TEMPO SILVER",     icon: "\u{26A1}",  description: "Reached Silver tier in Tempo",     color: "#c0c0c0" },
  "stat:tempo:gold":       { category: "stat", name: "TEMPO GOLD",       icon: "\u{26A1}",  description: "Reached Gold tier in Tempo",       color: "#ffd700" },
  "stat:tempo:platinum":   { category: "stat", name: "TEMPO PLATINUM",   icon: "\u{26A1}",  description: "Reached Platinum tier in Tempo",   color: "#e5e4e2" },
  "stat:rep:silver":       { category: "stat", name: "REP SILVER",       icon: "\u{1F451}", description: "Reached Silver tier in Reputation",   color: "#c0c0c0" },
  "stat:rep:gold":         { category: "stat", name: "REP GOLD",         icon: "\u{1F451}", description: "Reached Gold tier in Reputation",     color: "#ffd700" },
  "stat:rep:platinum":     { category: "stat", name: "REP PLATINUM",     icon: "\u{1F451}", description: "Reached Platinum tier in Reputation", color: "#e5e4e2" },

  // Meta badges — milestone achievements
  "meta:first-quest":    { category: "meta", name: "FIRST BLOOD",       icon: "\u{2694}",  description: "Completed your first quest",                       color: "#ff6600" },
  "meta:10-quests":      { category: "meta", name: "DECIMATOR",         icon: "\u{1F525}", description: "Completed 10 quests",                              color: "#ff6600" },
  "meta:25-quests":      { category: "meta", name: "QUARTER CENTURION", icon: "\u{1F31F}", description: "Completed 25 quests",                              color: "#ffea00" },
  "meta:50-quests":      { category: "meta", name: "HALF-CENTURY",      icon: "\u{1F4AB}", description: "Completed 50 quests",                              color: "#ff00ff" },
  "meta:100-quests":     { category: "meta", name: "CENTURION",         icon: "\u{1F3C6}", description: "Completed 100 quests",                             color: "#ffd700" },
  "meta:all-sectors":    { category: "meta", name: "TOTAL DOMINION",    icon: "\u{1F30D}", description: "Completed all 6 sectors",                          color: "#e5e4e2" },
  "meta:first-boss":     { category: "meta", name: "BOSS SLAYER",       icon: "\u{1F480}", description: "Defeated your first boss (all minions enslaved)",  color: "#ff0044" },
  "meta:survival-clear": { category: "meta", name: "SURVIVOR",          icon: "\u{1F6E1}", description: "Cleared all Survival Mode guardians",              color: "#ffd700" },
};

// ---------------------------------------------------------------------------
// Fetch all sheet data in one call (whole tables, headers included)
// ---------------------------------------------------------------------------
async function fetchSheetData(sheets) {
  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["Command_Center", "Definitions", "Sectors"],
  });
  const [cc, defs, sectors] = res.data.valueRanges;
  return {
    commandCenter: parseTable(cc.values),
    definitions: parseTable(defs.values),
    sectors: parseTable(sectors.values),
  };
}

// ---------------------------------------------------------------------------
// Verification code store (in-memory, codes expire after 10 minutes)
// ---------------------------------------------------------------------------
const verifyCodes = {}; // { "email@example.com": { code: "123456", expires: Date.now() + ms } }
const VERIFY_CODE_TTL = 10 * 60 * 1000; // 10 minutes

function generateVerifyCode(email) {
  const code = String(Math.floor(100000 + Math.random() * 900000));
  verifyCodes[email.toLowerCase()] = { code, expires: Date.now() + VERIFY_CODE_TTL };
  return code;
}

function checkVerifyCode(email, code) {
  const entry = verifyCodes[email.toLowerCase()];
  if (!entry) return false;
  if (Date.now() > entry.expires) { delete verifyCodes[email.toLowerCase()]; return false; }
  if (entry.code !== code) return false;
  delete verifyCodes[email.toLowerCase()]; // one-time use
  return true;
}

// Send verification code via Apps Script web app
async function sendVerifyEmail(email, code, name) {
  const scriptUrl = process.env.APPS_SCRIPT_URL;
  if (!scriptUrl) {
    console.warn("APPS_SCRIPT_URL not set — code will only be shown on verify page for testing");
    return false;
  }
  try {
    const url = `${scriptUrl}?email=${encodeURIComponent(email)}&code=${encodeURIComponent(code)}&name=${encodeURIComponent(name)}`;
    console.log("Calling Apps Script:", url.replace(/code=[^&]+/, "code=HIDDEN"));
    const resp = await fetch(url, { redirect: "follow" });
    const body = await resp.text();
    console.log("Apps Script response:", resp.status, body.substring(0, 200));
    if (body.trim() !== "OK") {
      console.error("Apps Script did not return OK:", body.substring(0, 500));
      return false;
    }
    return true;
  } catch (err) {
    console.error("Apps Script email error:", err.message);
    return false;
  }
}

// ---------------------------------------------------------------------------
// Ensure the Quests sheet tab exists (auto-create with headers if missing)
// ---------------------------------------------------------------------------
async function ensureQuestsSheet(sheets) {
  const expectedHeaders = ["Quest ID", "Boss", "Minion", "Sector", "Status", "Proof Type", "Proof Link", "Suggested By AI", "Date Completed", "Date Added", "Date Resolved", "Feedback"];
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: "sheets.properties.title",
  });
  const titles = meta.data.sheets.map((s) => s.properties.title);
  if (!titles.includes("Quests")) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{ addSheet: { properties: { title: "Quests" } } }],
      },
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Quests!A1:${String.fromCharCode(64 + expectedHeaders.length)}1`,
      valueInputOption: "RAW",
      requestBody: { values: [expectedHeaders] },
    });
    return;
  }
  // Check for missing columns and add them
  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Quests!1:1",
  });
  const currentHeaders = (headerRes.data.values && headerRes.data.values[0]) || [];
  const missing = expectedHeaders.filter((h) => !currentHeaders.includes(h));
  if (missing.length > 0) {
    const updated = [...currentHeaders, ...missing];
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Quests!A1:${String.fromCharCode(64 + updated.length)}1`,
      valueInputOption: "RAW",
      requestBody: { values: [updated] },
    });
  }
}

// ---------------------------------------------------------------------------
// Ensure Users sheet exists
// ---------------------------------------------------------------------------
async function ensureUsersSheet(sheets) {
  const expectedHeaders = ["Email", "Name", "Role"];
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: "sheets.properties.title",
  });
  const titles = meta.data.sheets.map((s) => s.properties.title);
  if (!titles.includes("Users")) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{ addSheet: { properties: { title: "Users" } } }],
      },
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Users!A1:${String.fromCharCode(64 + expectedHeaders.length)}1`,
      valueInputOption: "RAW",
      requestBody: { values: [expectedHeaders] },
    });
    return;
  }
  // Check for missing columns (e.g. PIN added later)
  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Users!1:1",
  });
  const currentHeaders = (headerRes.data.values && headerRes.data.values[0]) || [];
  const missing = expectedHeaders.filter((h) => !currentHeaders.includes(h));
  if (missing.length > 0) {
    const updated = [...currentHeaders, ...missing];
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Users!A1:${String.fromCharCode(64 + updated.length)}1`,
      valueInputOption: "RAW",
      requestBody: { values: [updated] },
    });
  }
}

// Look up user role from Users sheet by email
async function getUserRole(email) {
  if (!email) return null;
  try {
    const sheets = await getSheets();
    await ensureUsersSheet(sheets);
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Users",
    });
    const users = parseTable(res.data.values || []);
    const user = users.find((u) => (u["Email"] || "").toLowerCase() === email.toLowerCase());
    return user ? { name: user["Name"], role: (user["Role"] || "").toLowerCase(), email: user["Email"] } : null;
  } catch (err) {
    console.error("getUserRole error:", err.message);
    return null;
  }
}

// ---------------------------------------------------------------------------
// Ensure Teacher_Notes sheet exists
// ---------------------------------------------------------------------------
async function ensureTeacherNotesSheet(sheets) {
  const expectedHeaders = ["Date", "Author", "Subject", "Note"];
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: "sheets.properties.title",
  });
  const titles = meta.data.sheets.map((s) => s.properties.title);
  if (!titles.includes("Teacher_Notes")) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{ addSheet: { properties: { title: "Teacher_Notes" } } }],
      },
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: "Teacher_Notes!A1:D1",
      valueInputOption: "RAW",
      requestBody: { values: [expectedHeaders] },
    });
  }
}

async function ensureBadgesSheet(sheets) {
  const expectedHeaders = ["Badge ID", "Category", "Name", "Date Earned"];
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: "sheets.properties.title",
  });
  const titles = meta.data.sheets.map((s) => s.properties.title);
  if (!titles.includes("Badges")) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{ addSheet: { properties: { title: "Badges" } } }],
      },
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: "Badges!A1:D1",
      valueInputOption: "RAW",
      requestBody: { values: [expectedHeaders] },
    });
  }
}

// ---------------------------------------------------------------------------
// Fetch Quests sheet data
// ---------------------------------------------------------------------------
async function fetchQuestsData(sheets) {
  await ensureQuestsSheet(sheets);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Quests",
  });
  return parseTable(res.data.values);
}

// ---------------------------------------------------------------------------
// Collect artifact options from Definitions sheet
// ---------------------------------------------------------------------------
function getArtifactOptions(definitions) {
  const options = [];
  for (const row of definitions) {
    const artifact = (row["Artifact Options:"] || row["Artifact Options"] || "").trim();
    if (artifact && !options.includes(artifact)) {
      options.push(artifact);
    }
  }
  return options.length > 0 ? options : ["Document", "Spreadsheet", "Presentation", "Video", "Project"];
}

// ---------------------------------------------------------------------------
// Rule-based proof suggestion from sector stat weights
// ---------------------------------------------------------------------------
function generateProofSuggestion(sectorName, definitions) {
  const artifactOptions = getArtifactOptions(definitions);

  let weights = null;
  for (const row of definitions) {
    if (row["Sector"] && row["Sector"].toUpperCase() === sectorName.toUpperCase()) {
      weights = {
        intel: parseFloat(row["INTELLIGENCE"]) || 0,
        stamina: parseFloat(row["STAMINA"]) || 0,
        tempo: parseFloat(row["TEMPO"]) || 0,
        rep: parseFloat(row["REPUTATION"]) || 0,
      };
      break;
    }
  }

  if (!weights) {
    return { proofType: artifactOptions[0], suggestion: "Write a research summary about this minion.", artifactOptions };
  }

  const stats = [
    { name: "intel", value: weights.intel },
    { name: "stamina", value: weights.stamina },
    { name: "tempo", value: weights.tempo },
    { name: "rep", value: weights.rep },
  ];
  stats.sort((a, b) => b.value - a.value);

  const suggestions = {
    intel: { proofType: artifactOptions[0] || "Document", suggestion: "Write a research report or explanation demonstrating deep understanding." },
    stamina: { proofType: artifactOptions[1] || "Spreadsheet", suggestion: "Create a practice log or data tracker showing consistent effort over time." },
    tempo: { proofType: artifactOptions[2] || "Presentation", suggestion: "Deliver a timed presentation or speed drill demonstrating quick mastery." },
    rep: { proofType: artifactOptions[3] || "Video", suggestion: "Record a video teaching this concept to someone else or presenting it publicly." },
  };

  return { ...(suggestions[stats[0].name] || suggestions.intel), artifactOptions };
}

// ---------------------------------------------------------------------------
// Generate a short quest ID
// ---------------------------------------------------------------------------
function generateQuestId() {
  const ts = Date.now().toString(36).toUpperCase();
  const rand = Math.random().toString(36).substring(2, 5).toUpperCase();
  return "Q-" + ts + "-" + rand;
}

// ---------------------------------------------------------------------------
// Helper: sync Quest Status (and optionally main Status) to Sectors sheet
// ---------------------------------------------------------------------------
async function updateSectorsQuestStatus(sheets, sector, boss, minion, questStatus) {
  const secRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sectors",
  });
  const rows = secRes.data.values;
  if (!rows || rows.length < 2) return;

  const headers = rows[0];
  const questStatusCol = headers.indexOf("Quest Status");
  const statusCol = headers.indexOf("Status");
  const sectorCol = headers.indexOf("Sector");
  const bossCol = headers.indexOf("Boss");
  const minionCol = headers.indexOf("Minion");
  const dateAddedCol = headers.indexOf("Date Quest Added");
  const dateCompletedCol = headers.indexOf("Date Quest Completed");

  if (questStatusCol < 0 || sectorCol < 0 || bossCol < 0 || minionCol < 0) return;

  const colLetter = (idx) => {
    if (idx < 26) return String.fromCharCode(65 + idx);
    return String.fromCharCode(64 + Math.floor(idx / 26)) + String.fromCharCode(65 + (idx % 26));
  };

  const now = new Date().toISOString().slice(0, 10);

  const updates = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][sectorCol] === sector && rows[i][bossCol] === boss && rows[i][minionCol] === minion) {
      const rowNum = i + 1; // 1-based for Sheets API
      updates.push({ range: `Sectors!${colLetter(questStatusCol)}${rowNum}`, values: [[questStatus]] });

      // Auto-enslave when approved
      if (questStatus === "Approved" && statusCol >= 0) {
        updates.push({ range: `Sectors!${colLetter(statusCol)}${rowNum}`, values: [["Enslaved"]] });
        // Set Date Quest Completed
        if (dateCompletedCol >= 0) {
          updates.push({ range: `Sectors!${colLetter(dateCompletedCol)}${rowNum}`, values: [[now]] });
        }
      }

      // Quest added to board — set Date Quest Added
      if (questStatus === "Active" && dateAddedCol >= 0) {
        updates.push({ range: `Sectors!${colLetter(dateAddedCol)}${rowNum}`, values: [[now]] });
      }

      // Un-approve: revert to Engaged, clear completion date
      if ((questStatus === "Rejected" || questStatus === "Submitted") && statusCol >= 0) {
        const currentStatus = rows[i][statusCol] || "";
        if (currentStatus === "Enslaved") {
          updates.push({ range: `Sectors!${colLetter(statusCol)}${rowNum}`, values: [["Engaged"]] });
          if (dateCompletedCol >= 0) {
            updates.push({ range: `Sectors!${colLetter(dateCompletedCol)}${rowNum}`, values: [[""]] });
          }
        }
      }

      // Quest removed/abandoned — clear both dates
      if (questStatus === "" || questStatus === "Abandoned") {
        if (dateAddedCol >= 0) {
          updates.push({ range: `Sectors!${colLetter(dateAddedCol)}${rowNum}`, values: [[""]] });
        }
        if (dateCompletedCol >= 0) {
          updates.push({ range: `Sectors!${colLetter(dateCompletedCol)}${rowNum}`, values: [[""]] });
        }
      }

      break;
    }
  }

  if (updates.length > 0) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "RAW", data: updates },
    });

    // Auto-unlock: when a minion becomes Enslaved, check if any Locked minions' prerequisites are now met
    if (questStatus === "Approved") {
      try { await checkAndUnlockPrerequisites(sheets); } catch (e) { console.error("Auto-unlock check failed:", e.message); }
    }
  }
}

// ---------------------------------------------------------------------------
// Helper: look up a Command_Center row by stat name
// ---------------------------------------------------------------------------
function getStat(commandCenter, name) {
  const row = commandCenter.find(
    (r) => r["CORE STATS"].toLowerCase() === name.toLowerCase()
  );
  if (!row) return { value: 0, level: "", remaining: 0, totalPossible: 1 };
  return {
    value: parseFloat(row["VALUE (0-100)"]) || 0,
    level: row["Current Level"] || "",
    remaining: parseFloat(row["PTS Needed"]) || 0,
    totalPossible: parseFloat(row["Total Possible"]) || 1,
  };
}

// ---------------------------------------------------------------------------
// Build boss map from Sectors data (reusable utility)
// ---------------------------------------------------------------------------
function buildBossMap(sectors) {
  const bossMap = {};
  for (const row of sectors) {
    const sector = row["Sector"];
    const bossName = row["Boss"];
    const status = row["Status"];
    if (!sector || !bossName) continue;
    if (!bossMap[sector]) bossMap[sector] = {};
    if (!bossMap[sector][bossName]) {
      bossMap[sector][bossName] = { enslaved: 0, engaged: 0, locked: 0, total: 0 };
    }
    bossMap[sector][bossName].total++;
    if (status === "Enslaved") bossMap[sector][bossName].enslaved++;
    else if (status === "Engaged") bossMap[sector][bossName].engaged++;
    else if (status === "Locked") bossMap[sector][bossName].locked++;
  }
  return bossMap;
}

// ---------------------------------------------------------------------------
// Badge evaluation — returns Set of badge IDs earned based on current state
// ---------------------------------------------------------------------------
function evaluateBadges(sectors, commandCenter, quests, bossMap) {
  const earned = new Set();

  // Meta: quest count milestones
  const approvedCount = quests.filter((q) => q["Status"] === "Approved").length;
  if (approvedCount >= 1) earned.add("meta:first-quest");
  if (approvedCount >= 10) earned.add("meta:10-quests");
  if (approvedCount >= 25) earned.add("meta:25-quests");
  if (approvedCount >= 50) earned.add("meta:50-quests");
  if (approvedCount >= 100) earned.add("meta:100-quests");

  // Boss and sector badges
  let completedBossCount = 0;
  let completedSectors = 0;
  for (const sector in bossMap) {
    let allBossesComplete = true;
    for (const bossName in bossMap[sector]) {
      const b = bossMap[sector][bossName];
      if (b.total > 0 && b.enslaved === b.total) {
        earned.add(`boss:${sector}:${bossName}`);
        completedBossCount++;
      } else {
        allBossesComplete = false;
      }
    }
    if (allBossesComplete && Object.keys(bossMap[sector]).length > 0) {
      earned.add(`sector:${sector}`);
      completedSectors++;
    }
  }
  if (completedBossCount >= 1) earned.add("meta:first-boss");
  if (completedSectors >= 6) earned.add("meta:all-sectors");

  // Stat tier badges
  const statMap = { intel: "Intel", stamina: "Stamina", tempo: "Tempo", rep: "Reputation" };
  for (const [key, name] of Object.entries(statMap)) {
    const stat = getStat(commandCenter, name);
    const tierIdx = getStatTierIndex(stat.level);
    if (tierIdx >= 2) earned.add(`stat:${key}:silver`);
    if (tierIdx >= 3) earned.add(`stat:${key}:gold`);
    if (tierIdx >= 4) earned.add(`stat:${key}:platinum`);
  }

  // Survival clear badge
  let survivalCol = null;
  if (sectors.length > 0) {
    survivalCol = Object.keys(sectors[0]).find((k) => k.toLowerCase().includes("survival"));
  }
  if (survivalCol) {
    const survivalBosses = new Set();
    for (const r of sectors) {
      if ((r[survivalCol] || "").toUpperCase() === "X") {
        survivalBosses.add(`${r["Sector"]}|${r["Boss"]}`);
      }
    }
    let allSurvivalComplete = survivalBosses.size > 0;
    for (const key of survivalBosses) {
      const [sec, boss] = key.split("|");
      const b = bossMap[sec] && bossMap[sec][boss];
      if (!b || b.enslaved < b.total) { allSurvivalComplete = false; break; }
    }
    if (allSurvivalComplete && survivalBosses.size > 0) earned.add("meta:survival-clear");
  }

  return earned;
}

// ---------------------------------------------------------------------------
// Sync badges — compare earned vs sheet, write new ones
// ---------------------------------------------------------------------------
async function syncBadges(sheets, sectors, commandCenter, quests, bossMap) {
  await ensureBadgesSheet(sheets);

  const badgesRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Badges",
  });
  const currentBadges = parseTable(badgesRes.data.values || []);
  const alreadyEarned = new Set(currentBadges.map((b) => b["Badge ID"]));

  const shouldBeEarned = evaluateBadges(sectors, commandCenter, quests, bossMap);

  const newBadges = [];
  const now = new Date().toISOString().slice(0, 10);
  for (const badgeId of shouldBeEarned) {
    if (!alreadyEarned.has(badgeId)) {
      const parts = badgeId.split(":");
      const def = BADGE_DEFINITIONS[badgeId] || getBossBadgeDef(parts[1], parts.slice(2).join(":"));
      newBadges.push([badgeId, def.category, def.name, now]);
    }
  }

  if (newBadges.length > 0) {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Badges!A:D",
      valueInputOption: "RAW",
      requestBody: { values: newBadges },
    });
  }

  return {
    allEarned: currentBadges.concat(
      newBadges.map((row) => ({
        "Badge ID": row[0], "Category": row[1], "Name": row[2], "Date Earned": row[3],
      }))
    ),
    newlyEarned: newBadges.map((row) => row[0]),
  };
}

// ---------------------------------------------------------------------------
// HUD badge showcase HTML builder
// ---------------------------------------------------------------------------
function buildHudBadgeHtml(badgeResult) {
  const { allEarned, newlyEarned } = badgeResult;
  if (allEarned.length === 0) return "";

  const sorted = [...allEarned]
    .sort((a, b) => (b["Date Earned"] || "").localeCompare(a["Date Earned"] || ""))
    .slice(0, 8);

  const items = sorted.map((b) => {
    const parts = b["Badge ID"].split(":");
    const def = BADGE_DEFINITIONS[b["Badge ID"]] || getBossBadgeDef(parts[1], parts.slice(2).join(":"));
    const isNew = newlyEarned.includes(b["Badge ID"]);
    return `<div class="badge-item${isNew ? " badge-new" : ""}">
      <span class="badge-icon" style="color:${def.color}">${def.icon}</span>
      <span class="badge-name">${escHtml(def.name)}</span>
      <span class="badge-date">${(b["Date Earned"] || "").slice(0, 10)}</span>
    </div>`;
  }).join("");

  return `<div class="badge-showcase">
    <h2>\u{1F3C6} ACHIEVEMENTS (${allEarned.length})</h2>
    <div class="badge-row">${items}</div>
    <a href="/badges" class="badge-view-all">VIEW ALL BADGES &gt;&gt;</a>
  </div>`;
}

// ---------------------------------------------------------------------------
// Rejected quest alerts HTML builder
// ---------------------------------------------------------------------------
function buildRejectedAlertsHtml(quests) {
  const rejected = quests.filter((q) => q["Status"] === "Rejected");
  if (rejected.length === 0) return "";

  const rows = rejected.map((q) =>
    `<div class="rejected-entry">
      <span>${escHtml(q["Boss"])} &gt; ${escHtml(q["Minion"])}</span>
      ${q["Feedback"] ? `<div class="rejected-reason">${escHtml(q["Feedback"])}</div>` : ""}
    </div>`
  ).join("");

  return `<div class="rejected-alerts">
    <h2>\u{26A0} ${rejected.length} REJECTED QUEST${rejected.length > 1 ? "S" : ""}</h2>
    ${rows}
    <a href="/quests" style="display:block;text-align:center;color:#ff0044;font-size:0.7em;margin-top:8px;text-decoration:none;letter-spacing:2px;">GO TO QUEST BOARD &gt;&gt;</a>
  </div>`;
}

// ---------------------------------------------------------------------------
// Streak computation — aggregate activity dates, compute current & best streak
// ---------------------------------------------------------------------------
function computeStreak(quests, sectors) {
  const dates = new Set();

  // Quest activity dates
  for (const q of quests) {
    const added = (q["Date Added"] || "").slice(0, 10);
    const resolved = (q["Date Resolved"] || "").slice(0, 10);
    if (added && added.length === 10) dates.add(added);
    if (resolved && resolved.length === 10) dates.add(resolved);
  }

  // Minion enslaved dates
  for (const s of sectors) {
    const completed = (s["Date Quest Completed"] || "").slice(0, 10);
    if (completed && completed.length === 10) dates.add(completed);
  }

  if (dates.size === 0) return { currentStreak: 0, bestStreak: 0, totalActiveDays: 0 };

  const sorted = [...dates].sort();
  const today = new Date().toISOString().slice(0, 10);

  // Helper: get previous date string
  function prevDay(dateStr) {
    const d = new Date(dateStr + "T12:00:00Z");
    d.setUTCDate(d.getUTCDate() - 1);
    return d.toISOString().slice(0, 10);
  }

  // Current streak: walk backward from today
  let currentStreak = 0;
  let check = today;
  while (dates.has(check)) {
    currentStreak++;
    check = prevDay(check);
  }

  // Best streak: scan all sorted dates
  let bestStreak = 0;
  let run = 1;
  for (let i = 1; i < sorted.length; i++) {
    const expected = prevDay(sorted[i]);
    if (sorted[i - 1] === expected) {
      run++;
    } else {
      run = 1;
    }
    if (run > bestStreak) bestStreak = run;
  }
  if (sorted.length === 1) bestStreak = 1;
  if (run > bestStreak) bestStreak = run;
  if (currentStreak > bestStreak) bestStreak = currentStreak;

  return { currentStreak, bestStreak, totalActiveDays: dates.size };
}

function buildStreakHtml(streakData) {
  const { currentStreak, bestStreak } = streakData;
  const isRecord = currentStreak > 0 && currentStreak >= bestStreak;

  if (currentStreak === 0) {
    return `<div class="streak-display streak-inactive">
      <span class="streak-fire">\u{1F525}</span>
      <span class="streak-label">NO ACTIVE STREAK</span>
    </div>`;
  }

  return `<div class="streak-display${isRecord ? " streak-record" : ""}">
    <span class="streak-fire">\u{1F525}</span>
    <span class="streak-count">${currentStreak}</span>
    <span class="streak-label">DAY STREAK</span>
    <span class="streak-best">(BEST: ${bestStreak})</span>
    ${isRecord ? '<span class="streak-new-record">NEW RECORD!</span>' : ""}
  </div>`;
}

// ---------------------------------------------------------------------------
// SVG Radar Chart Generator
// ---------------------------------------------------------------------------
function buildRadarSVG(values, maxVal, ringCount, size, options = {}) {
  const {
    labels = ["INT", "STA", "TMP", "REP"],
    ringLabels = [],
    fillColor = "rgba(0,242,255,0.2)",
    strokeColor = "#00f2ff",
    showLabels = true,
    axisColors = null, // array of 4 colors, one per axis
  } = options;

  const cx = size / 2;
  const cy = size / 2;
  const pad = showLabels ? 28 : 10;
  const R = size / 2 - pad;

  // 4 axes: top, right, bottom, left
  const angles = [-Math.PI / 2, 0, Math.PI / 2, Math.PI];

  const pt = (frac, ai) => {
    const x = cx + frac * R * Math.cos(angles[ai]);
    const y = cy + frac * R * Math.sin(angles[ai]);
    return { x, y };
  };

  let svg = `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" xmlns="http://www.w3.org/2000/svg" style="overflow:visible">`;

  // Concentric ring polygons
  for (let r = 1; r <= ringCount; r++) {
    const f = r / ringCount;
    const pts = angles.map((_, i) => { const p = pt(f, i); return `${p.x},${p.y}`; }).join(" ");
    svg += `<polygon points="${pts}" fill="none" stroke="rgba(0,242,255,0.15)" stroke-width="1"/>`;
  }

  // Axis lines
  for (let i = 0; i < 4; i++) {
    const p = pt(1, i);
    const axColor = axisColors ? axisColors[i] : "rgba(0,242,255,0.3)";
    svg += `<line x1="${cx}" y1="${cy}" x2="${p.x}" y2="${p.y}" stroke="${axColor}" stroke-width="1" opacity="0.5"/>`;
  }

  // Data polygon
  const dataPts = values.map((v, i) => {
    const f = Math.min(v, maxVal) / maxVal;
    const p = pt(f, i);
    return `${p.x},${p.y}`;
  }).join(" ");
  svg += `<polygon points="${dataPts}" fill="${fillColor}" stroke="${strokeColor}" stroke-width="2"/>`;

  // Data dots
  values.forEach((v, i) => {
    const f = Math.min(v, maxVal) / maxVal;
    const p = pt(f, i);
    const dotColor = axisColors ? axisColors[i] : strokeColor;
    svg += `<circle cx="${p.x}" cy="${p.y}" r="${size > 120 ? 4 : 2.5}" fill="${dotColor}"/>`;
  });

  // Axis labels
  if (showLabels) {
    const lo = R + (size > 120 ? 14 : 10);
    const fontSize = size > 120 ? 14 : 10;
    labels.forEach((label, i) => {
      const x = cx + lo * Math.cos(angles[i]);
      const y = cy + lo * Math.sin(angles[i]);
      const anchor = i === 1 ? "start" : i === 3 ? "end" : "middle";
      const dy = i === 0 ? "-0.2em" : i === 2 ? "1em" : "0.35em";
      const labelColor = axisColors ? axisColors[i] : "#00f2ff";
      svg += `<text x="${x}" y="${y}" text-anchor="${anchor}" dy="${dy}" fill="${labelColor}" font-size="${fontSize}" font-family="'Courier New',monospace" style="text-transform:uppercase">${label}</text>`;
    });
  }

  // Ring labels (along the top axis)
  for (let r = 0; r < ringLabels.length && r < ringCount; r++) {
    const f = (r + 1) / ringCount;
    const y = cy - f * R;
    svg += `<text x="${cx + 4}" y="${y - 2}" fill="rgba(0,242,255,0.35)" font-size="7" font-family="'Courier New',monospace">${ringLabels[r]}</text>`;
  }

  svg += `</svg>`;
  return svg;
}

// ---------------------------------------------------------------------------
// Escape HTML special chars for safe embedding in attributes
// ---------------------------------------------------------------------------
function escHtml(str) {
  return String(str).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// ---------------------------------------------------------------------------
// Process data -> HTML
// ---------------------------------------------------------------------------
function processAllData(html, data, activeQuestKeys) {
  const { commandCenter, definitions, sectors } = data;

  // Core stats by name
  const intel = getStat(commandCenter, "Intel");
  const stamina = getStat(commandCenter, "Stamina");
  const tempo = getStat(commandCenter, "Tempo");
  const reputation = getStat(commandCenter, "Reputation");
  const confidence = getStat(commandCenter, "Confidence");
  const currentConfRank = confidence.level || "BRONZE: INITIATE";

  // Next tier sub-rank helper — extracts just the numeral (e.g. "III") for the next tier
  function nextSubRank(currentLevel) {
    for (let i = 0; i < definitions.length; i++) {
      const tierName = definitions[i]["Name"];
      if (tierName && tierName === currentLevel) {
        if (i + 1 < definitions.length && definitions[i + 1]["Name"]) {
          return definitions[i + 1]["Name"]; // e.g. "Silver I"
        }
        return "MAX";
      }
    }
    return "MAX";
  }

  const nextConfSub = nextSubRank(currentConfRank);
  const nextIntelSub = nextSubRank(intel.level);
  const nextStaminaSub = nextSubRank(stamina.level);
  const nextTempoSub = nextSubRank(tempo.level);
  const nextRepSub = nextSubRank(reputation.level);

  // -- MAIN RADAR CHART (4 stats, 0-100 scale, 5 level rings) --
  // Colors match the stat bars: Intel=cyan, Stamina=green, Tempo=magenta, Rep=orange
  const statColors = ["#00f2ff", "#00ff9d", "#ff00ff", "#ff8800"];
  const mainRadar = buildRadarSVG(
    [intel.value, stamina.value, tempo.value, reputation.value],
    100, 5, 200,
    {
      labels: ["INTEL", "STA", "TEMPO", "REP"],
      ringLabels: ["BRZ", "CPR", "SLV", "GLD", "PLT"],
      fillColor: "rgba(255,255,255,0.08)",
      strokeColor: "rgba(255,255,255,0.5)",
      axisColors: statColors,
    }
  );

  // -- SECTOR WEIGHT MAP from Definitions --
  const sectorWeights = {};
  for (const row of definitions) {
    if (row["Sector"]) {
      sectorWeights[row["Sector"].toUpperCase()] = {
        intel: parseFloat(row["INTELLIGENCE"]) || 0,
        stamina: parseFloat(row["STAMINA"]) || 0,
        tempo: parseFloat(row["TEMPO"]) || 0,
        rep: parseFloat(row["REPUTATION"]) || 0,
      };
    }
  }

  // -- BUILD BOSS MAP --
  const bossMap = buildBossMap(sectors);

  // -- BUILD SECTOR HTML (with radar + clickable bosses) --
  let bossHtml = "";
  for (const sector in bossMap) {
    // Sector weight radar chart
    const w = sectorWeights[sector.toUpperCase()];
    let sectorRadarHtml = "";
    if (w) {
      sectorRadarHtml = `<div class="sector-radar">${buildRadarSVG(
        [w.intel, w.stamina, w.tempo, w.rep],
        5, 5, 100,
        {
          labels: ["INT", "STA", "TMP", "REP"],
          fillColor: "rgba(255,255,255,0.08)",
          strokeColor: "rgba(255,255,255,0.5)",
          showLabels: true,
          axisColors: ["#00f2ff", "#00ff9d", "#ff00ff", "#ff8800"],
        }
      )}</div>`;
    }

    let sectorBossHtml = "";
    const sortedBossNames = Object.keys(bossMap[sector]).sort(
      (a, b) => bossMap[sector][b].total - bossMap[sector][a].total
    );

    for (const bossName of sortedBossNames) {
      const stats = bossMap[sector][bossName];
      const fraction = `(${stats.enslaved}/${stats.total})`;
      const diameter = 35 + stats.total * 6;
      const safeTotal = stats.total || 1;

      const pEnslaved = (stats.enslaved / safeTotal) * 100;
      const pEngaged = (stats.engaged / safeTotal) * 100;

      const gradient = `conic-gradient(
        #00ff9d 0% ${pEnslaved}%,
        #0a0b10 ${pEnslaved}% ${pEnslaved + 1}%,
        #ff6600 ${pEnslaved + 1}% ${pEnslaved + 1 + pEngaged}%,
        #0a0b10 ${pEnslaved + 1 + pEngaged}% ${pEnslaved + 2 + pEngaged}%,
        #2a2d36 ${pEnslaved + 2 + pEngaged}% 100%
      )`;

      const encodedBoss = encodeURIComponent(bossName);
      const encodedSector = encodeURIComponent(sector);

      sectorBossHtml += `
        <a class="boss-link" href="/boss/${encodedBoss}?sector=${encodedSector}" >
          <div class="boss-orb-container">
            <div class="boss-pie" style="width:${diameter}px; height:${diameter}px; background:${gradient};"></div>
            <div class="boss-label">${bossName} ${fraction}</div>
          </div>
        </a>`;
    }

    bossHtml += `<div class="sector-zone" data-sector="${sector.toUpperCase()}"><a class="sector-link" href="/sector/${encodeURIComponent(sector)}">${sector.toUpperCase()}</a>${sectorRadarHtml}<div class="sector-bosses">${sectorBossHtml}</div></div>`;
  }

  // -- SURVIVAL RING (Ring of Guardians) --
  // Find the survival column name dynamically (handles "Survival Mode Required", "Survival Mode Requirements", etc.)
  let survivalCol = null;
  if (sectors.length > 0) {
    const firstRow = sectors[0];
    survivalCol = Object.keys(firstRow).find((k) => k.toLowerCase().includes("survival"));
  }

  let survivalRingHtml = "";
  let survivalModeHeaderHtml = "";
  const survivalBossKeys = new Set();
  const survivalBosses = [];
  for (const row of sectors) {
    if (survivalCol && (row[survivalCol] || "").toUpperCase() === "X") {
      const key = `${row["Sector"]}|${row["Boss"]}`;
      if (!survivalBossKeys.has(key)) {
        survivalBossKeys.add(key);
        survivalBosses.push({ sector: row["Sector"], bossName: row["Boss"] });
      }
    }
  }

  if (survivalBosses.length > 0) {
    // Sort bosses: highest conquered fraction first, then alphabetical by name
    survivalBosses.sort((a, b) => {
      const aData = bossMap[a.sector]?.[a.bossName] || { enslaved: 0, total: 1 };
      const bDataB = bossMap[b.sector]?.[b.bossName] || { enslaved: 0, total: 1 };
      const aFrac = aData.total > 0 ? aData.enslaved / aData.total : 0;
      const bFrac = bDataB.total > 0 ? bDataB.enslaved / bDataB.total : 0;
      if (bFrac !== aFrac) return bFrac - aFrac; // descending by fraction
      return a.bossName.localeCompare(b.bossName); // ascending by name
    });

    let orbsHtml = "";
    let totalCompleted = 0;
    let totalMinions = 0;

    for (const sb of survivalBosses) {
      const bData = bossMap[sb.sector]?.[sb.bossName];
      if (!bData) continue;

      const pct = bData.total > 0 ? (bData.enslaved / bData.total) * 100 : 0;
      totalCompleted += bData.enslaved;
      totalMinions += bData.total;

      let stateClass;
      if (pct >= 100) {
        stateClass = "complete";
      } else if (bData.enslaved > 0) {
        stateClass = "in-progress";
      } else {
        stateClass = "not-started";
      }

      // Conic gradient for fill (red for progress, green when complete)
      const pEnslaved = (bData.enslaved / (bData.total || 1)) * 100;
      const gradient = pct >= 100
        ? "background: #00ff9d;"
        : `background: conic-gradient(#ff4444 0% ${pEnslaved.toFixed(1)}%, #1a1d26 ${pEnslaved.toFixed(1)}% 100%);`;

      const encodedBoss = encodeURIComponent(sb.bossName);
      const encodedSector = encodeURIComponent(sb.sector);

      orbsHtml += `
        <a class="guardian-orb ${stateClass}" href="/boss/${encodedBoss}?sector=${encodedSector}">
          <div class="guardian-orb-circle" style="${gradient}"></div>
          <div class="guardian-label">${escHtml(sb.bossName)}</div>
          <div class="guardian-fraction">${bData.enslaved}/${bData.total}</div>
        </a>`;
    }

    const overallPct = totalMinions > 0 ? Math.round((totalCompleted / totalMinions) * 100) : 0;
    const allGuardiansComplete = survivalBosses.length > 0 && overallPct >= 100;

    // Minecraft-style pixel heart SVG — glow proportional to boss completion
    const heartSvg = (fraction, gold) => {
      // fraction: 0.0 (empty) to 1.0 (full), gold: true when all guardians complete
      const f = Math.max(0, Math.min(1, fraction));
      let color, shadow, highlight;
      if (gold) {
        color = "#ffd700"; shadow = "#cc9900"; highlight = "#ffe680";
      } else if (f >= 1) {
        color = "#ff4444"; shadow = "#aa0000"; highlight = "#ff8888";
      } else if (f > 0) {
        // Blend from dim (#553333) to red (#ff4444) based on fraction
        const r = Math.round(0x55 + (0xff - 0x55) * f);
        const g = Math.round(0x33 * (1 - f));
        const b = Math.round(0x33 * (1 - f));
        color = `rgb(${r},${g},${b})`;
        shadow = "#1a1d26";
        highlight = `rgba(255,136,136,${(f * 0.7).toFixed(2)})`;
      } else {
        color = "#333"; shadow = "#1a1d26"; highlight = "#444";
      }
      const glowRadius = f > 0 ? Math.round(2 + f * 4) : 0;
      const glow = glowRadius > 0 ? `filter="drop-shadow(0 0 ${glowRadius}px ${color})"` : "";
      return `<svg width="16" height="14" viewBox="0 0 10 9" shape-rendering="crispEdges" ${glow}>
        <rect x="1" y="0" width="3" height="1" fill="${color}"/>
        <rect x="6" y="0" width="3" height="1" fill="${color}"/>
        <rect x="0" y="1" width="10" height="1" fill="${color}"/>
        <rect x="0" y="2" width="10" height="1" fill="${color}"/>
        <rect x="1" y="3" width="8" height="1" fill="${color}"/>
        <rect x="2" y="4" width="6" height="1" fill="${color}"/>
        <rect x="3" y="5" width="4" height="1" fill="${color}"/>
        <rect x="4" y="6" width="2" height="1" fill="${shadow}"/>
        <rect x="1" y="1" width="2" height="1" fill="${highlight}"/>
      </svg>`;
    };

    // Build hearts row — one per guardian boss, glow proportional to enslaved fraction
    let heartsHtml = "";
    for (const sb of survivalBosses) {
      const bData = bossMap[sb.sector]?.[sb.bossName];
      if (!bData) continue;
      const fraction = bData.total > 0 ? bData.enslaved / bData.total : 0;
      heartsHtml += `<span class="guardian-heart" title="${escHtml(sb.bossName)} (${bData.enslaved}/${bData.total})">${heartSvg(fraction, allGuardiansComplete)}</span>`;
    }

    // Mode banner + hearts — goes under confidence bar
    survivalModeHeaderHtml = allGuardiansComplete
      ? `<div class="mode-header mode-header-achieved">
           <div class="mode-banner-line"><span class="mode-banner mode-survival">GAME MODE: SURVIVAL</span><span class="hearts-row">${heartsHtml}</span></div>
           <div class="mode-subtitle mode-achieved">READY FOR THE REAL WORLD</div>
         </div>`
      : `<div class="mode-header">
           <div class="mode-banner-line"><span class="mode-banner mode-creative">GAME MODE: CREATIVE</span><span class="hearts-row">${heartsHtml}</span></div>
           <div class="mode-subtitle">ENSLAVE ALL GUARDIANS TO ENTER SURVIVAL MODE</div>
         </div>`;

    // Ring of Guardians orbs — stays in its own section
    const shieldIcon = `<svg width="18" height="20" viewBox="0 0 18 20" fill="none" style="vertical-align:middle;margin-right:4px;"><path d="M9 0L0 3.5V9C0 14 3.8 18.5 9 20C14.2 18.5 18 14 18 9V3.5L9 0Z" fill="rgba(255,68,68,0.3)" stroke="#ff4444" stroke-width="1.5"/><text x="9" y="13" text-anchor="middle" fill="#ff4444" font-size="8" font-family="monospace" font-weight="bold">S</text></svg>`;

    survivalRingHtml = `
      <div class="survival-ring${allGuardiansComplete ? ' survival-achieved' : ''}">
        <a class="survival-ring-title" href="/guardians">${shieldIcon} RING OF GUARDIANS</a>
        <div class="boss-key">
            <span class="key-item"><span class="key-swatch" style="background:#00ff9d;"></span> Enslaved</span>
            <span class="key-item"><span class="key-swatch" style="background:#ff6600;"></span> Engaged</span>
            <span class="key-item"><span class="key-swatch" style="background:#2a2d36; border: 1px solid #555;"></span> Locked</span>
        </div>
        <div class="guardian-ring">${orbsHtml}</div>
        <div class="survival-summary">${overallPct}% GUARDIAN PROTOCOL COMPLETE</div>
      </div>`;
  }

  // -- BOSS CONQUEST RANKINGS (bosses with remaining minions) --
  const incompleteBosses = [];
  for (const sector in bossMap) {
    for (const bossName in bossMap[sector]) {
      const stats = bossMap[sector][bossName];
      const remaining = stats.total - stats.enslaved;
      if (remaining > 0) {
        incompleteBosses.push({
          sector, bossName, ...stats,
          remaining,
          pct: stats.total > 0 ? (stats.enslaved / stats.total) * 100 : 0,
        });
      }
    }
  }
  // Sort: highest completion % first, then fewest remaining, then alphabetical
  incompleteBosses.sort((a, b) => {
    if (b.pct !== a.pct) return b.pct - a.pct;
    if (a.remaining !== b.remaining) return a.remaining - b.remaining;
    return a.bossName.localeCompare(b.bossName);
  });

  let rankHtml = "";
  incompleteBosses.slice(0, 10).forEach((boss, i) => {
    // Color gradient: gold for nearly complete → cyan for just started
    const t = incompleteBosses.length > 1 ? i / (incompleteBosses.length - 1) : 0;
    const r = Math.round(255 - t * 255 + t * 0);
    const g = Math.round(215 - t * 215 + t * 242);
    const b = Math.round(0 + t * 255);
    const color = `rgb(${r},${g},${b})`;

    const statusLabel = boss.pct >= 75 ? "NEARLY CONQUERED"
      : boss.pct >= 50 ? "HALFWAY THERE"
      : boss.pct > 0 ? "IN PROGRESS"
      : "NOT STARTED";

    const encodedBoss = encodeURIComponent(boss.bossName);
    const encodedSector = encodeURIComponent(boss.sector);

    rankHtml += `
      <a class="rank-entry" href="/boss/${encodedBoss}?sector=${encodedSector}" style="text-decoration:none;color:inherit;">
        <span class="rank-badge" style="color:${color}">#${i + 1}</span>
        <span class="rank-boss" style="color:${color}">${escHtml(boss.bossName)}</span>
        <div class="rank-bar-wrap">
          <div class="rank-bar" style="width:${boss.pct.toFixed(1)}%; background:${color};"></div>
          <span class="rank-info">${boss.enslaved}/${boss.total} ENSLAVED &mdash; ${boss.remaining} REMAINING</span>
        </div>
        <span class="rank-sector">${escHtml(boss.sector)}</span>
        <span class="rank-status-label" style="color:${color}">${statusLabel}</span>
      </a>`;
  });

  // -- TIER LIST (5 metals, sub-rank shown only on active) --
  const metals = [
    { rank: "platinum", label: "PLATINUM", subs: ["Platinum I", "Platinum II", "Platinum III"] },
    { rank: "gold",     label: "GOLD",     subs: ["Gold I", "Gold II", "Gold III"] },
    { rank: "silver",   label: "SILVER",   subs: ["Silver I", "Silver II", "Silver III"] },
    { rank: "copper",   label: "COPPER",   subs: ["Copper I", "Copper II", "Copper III"] },
    { rank: "bronze",   label: "BRONZE",   subs: ["Bronze I", "Bronze II", "Bronze III"] },
  ];
  const subNumerals = { 1: "I", 2: "II", 3: "III" };
  let activeMetalIdx = metals.length - 1;
  let activeSub = "";
  for (let i = 0; i < metals.length; i++) {
    const subIdx = metals[i].subs.indexOf(currentConfRank);
    if (subIdx !== -1) {
      activeMetalIdx = i;
      activeSub = " (" + subNumerals[subIdx + 1] + ")";
      break;
    }
  }
  let tierListHtml = '<ul class="tier-list">';
  for (let i = 0; i < metals.length; i++) {
    let cls = "tier-item rank-" + metals[i].rank;
    if (i === activeMetalIdx) cls += " active";
    else if (i === activeMetalIdx - 1) cls += " next";
    const suffix = i === activeMetalIdx ? activeSub : "";
    tierListHtml += '<li class="' + cls + '">' + metals[i].label + suffix + "</li>";
  }
  tierListHtml += "</ul>";

  // -- DYNAMIC BAR PROGRESS (based on tier thresholds) --
  // For status types, use "Status pts" column; for confidence, use "Confidence pts"
  function tierBar(value, levelName, ptsCol) {
    const idx = definitions.findIndex((d) => d["Name"] === levelName);
    if (idx === -1) return "0";
    const entry = parseFloat(definitions[idx][ptsCol]) || 0;
    const next = idx + 1 < definitions.length
      ? parseFloat(definitions[idx + 1][ptsCol]) || entry + 1
      : entry + 1;
    const span = next - entry || 1;
    return Math.min(100, ((value - entry) / span) * 100).toFixed(1);
  }

  // -- TEMPLATE REPLACEMENTS --
  return html
    .split("[[INTEL_RANK]]").join(intel.level)
    .split("[[INTEL_REM]]").join(intel.remaining.toFixed(1))
    .split("[[INTEL_NEXT]]").join(nextIntelSub)
    .split("[[INTEL_BAR]]").join(tierBar(intel.value, intel.level, "Status pts"))
    .split("[[STAMINA_RANK]]").join(stamina.level)
    .split("[[STAMINA_REM]]").join(stamina.remaining.toFixed(1))
    .split("[[STAMINA_NEXT]]").join(nextStaminaSub)
    .split("[[STAMINA_BAR]]").join(tierBar(stamina.value, stamina.level, "Status pts"))
    .split("[[TEMPO_RANK]]").join(tempo.level)
    .split("[[TEMPO_REM]]").join(tempo.remaining.toFixed(1))
    .split("[[TEMPO_NEXT]]").join(nextTempoSub)
    .split("[[TEMPO_BAR]]").join(tierBar(tempo.value, tempo.level, "Status pts"))
    .split("[[REP_RANK]]").join(reputation.level)
    .split("[[REPUTATION_REM]]").join(reputation.remaining.toFixed(1))
    .split("[[REP_NEXT]]").join(nextRepSub)
    .split("[[REPUTATION_BAR]]").join(tierBar(reputation.value, reputation.level, "Status pts"))
    .split("[[CONF_RANK]]").join(currentConfRank)
    .split("[[CONF_NEXT]]").join(nextConfSub)
    .split("[[CONF_REM]]").join(confidence.remaining.toFixed(1))
    .split("[[CONF_BAR]]").join(tierBar(confidence.value, currentConfRank, "Confidence pts"))
    .split("[[MAIN_RADAR]]").join(mainRadar)
    .split("[[TIER_LIST]]").join(tierListHtml)
    .split("[[BOSS_LIST]]").join(bossHtml)
    .split("[[POWER_RANKINGS]]").join(rankHtml)
    .split("[[SURVIVAL_RING]]").join(survivalRingHtml)
    .split("[[SURVIVAL_MODE_HEADER]]").join(survivalModeHeaderHtml);
}

// ---------------------------------------------------------------------------
// HTML Template
// ---------------------------------------------------------------------------
const HTML_TEMPLATE = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Sovereign HUD</title>
    <style>
    body {
        background: #0a0b10;
        color: #00f2ff;
        font-family: 'Courier New', monospace;
        padding: 20px;
        text-transform: uppercase;
        overflow-x: hidden;
    }
    .hud-container {
        border: 2px solid #00f2ff;
        padding: 20px;
        box-shadow: 0 0 15px #00f2ff;
        max-width: 960px;
        margin: auto;
    }
    h1 {
        text-shadow: 2px 2px #ff00ff;
        letter-spacing: 2px;
        text-align: center;
        margin-bottom: 15px;
    }
    /* Title block */
    .hud-title-block {
        text-align: center;
        padding: 15px 0 20px 0;
        margin-bottom: 10px;
    }
    .title-tagline {
        font-family: 'Courier New', monospace;
        font-size: 2em;
        letter-spacing: 6.6px;
        color: rgba(0, 242, 255, 0.6);
        text-transform: uppercase;
        margin-bottom: 4px;
        text-shadow: 0 0 8px rgba(0, 242, 255, 0.2);
    }
    .title-main {
        border-bottom: none;
        margin-bottom: 0;
        padding: 0;
        text-shadow: none;
    }
    .title-main span {
        background: linear-gradient(90deg, #00f2ff, #ff00ff, #ffea00, #00ff9d, #00f2ff);
        background-size: 300% 100%;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        animation: titleShimmer 6s linear infinite;
        font-size: 1.6em;
        letter-spacing: 5px;
        filter: drop-shadow(0 0 12px rgba(0, 242, 255, 0.4)) drop-shadow(0 0 25px rgba(255, 0, 255, 0.2));
    }
    @keyframes titleShimmer {
        0% { background-position: 0% 50%; }
        100% { background-position: 300% 50%; }
    }

    /* Confidence row: bar + radar side by side */
    .confidence-row {
        display: flex;
        align-items: center;
        gap: 25px;
        margin-bottom: 25px;
        padding-bottom: 15px;
    }
    .confidence-bar-section { flex: 3; min-width: 0; display: flex; flex-direction: column; justify-content: center; }
    .radar-main { flex-shrink: 0; }

    .stat-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 20px;
        margin-bottom: 40px;
    }
    .stat-label { font-size: 0.85em; font-weight: bold; margin-bottom: 5px; }
    .stat-next { font-size: 0.8em; font-weight: normal; opacity: 0.6; letter-spacing: 0.5px; }
    .stat-bar { background: #1a1d26; border: 1px solid #333; height: 25px; position: relative; overflow: hidden; }
    .fill { height: 100%; transition: width 0.5s ease-in-out; }
    .intel-fill      { background: linear-gradient(90deg, #00f2ff, #0077ff); box-shadow: 0 0 15px #00f2ff; }
    .stamina-fill    { background: linear-gradient(90deg, #00ff9d, #008844); box-shadow: 0 0 10px #00ff9d; }
    .tempo-fill      { background: linear-gradient(90deg, #ff00ff, #880088); box-shadow: 0 0 10px #ff00ff; }
    .rep-fill        { background: linear-gradient(90deg, #ff8800, #ff4400); box-shadow: 0 0 10px #ff8800; }
    .confidence-fill { background: linear-gradient(90deg, #ffea00, #ffaa00); box-shadow: 0 0 10px #ffea00; }

    .levels-right {
        flex: 1;
        padding-left: 20px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .tier-list { list-style: none; padding: 0; margin: 0; text-align: center; }
    .tier-item {
        font-size: 1.1em;
        margin: 10px 0;
        transition: all 0.3s ease;
        letter-spacing: 1px;
    }
    .rank-platinum { color: #e5e4e2; opacity: 0.4; }
    .rank-gold     { color: #ffd700; opacity: 0.4; }
    .rank-silver   { color: #c0c0c0; opacity: 0.4; }
    .rank-copper   { color: #b87333; opacity: 0.4; }
    .rank-bronze   { color: #cd7f32; opacity: 0.4; }
    .tier-item.active {
        opacity: 1;
        font-weight: bold;
        text-shadow: 0 0 12px currentColor;
    }
    .tier-item.next { opacity: 0.7; }

    /* Recently enslaved */
    .recent-enslaved {
        margin-top: 15px;
        border: 1px solid rgba(0, 255, 157, 0.2);
        background: rgba(0, 255, 157, 0.02);
        border-radius: 6px;
        padding: 12px 15px;
    }
    .recent-enslaved h2 {
        color: #00ff9d;
        font-size: 0.85em;
        letter-spacing: 3px;
        margin: 0 0 10px 0;
        text-align: center;
        text-shadow: 0 0 8px rgba(0, 255, 157, 0.4);
    }
    .re-entry {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 4px 0;
        font-size: 0.75em;
        border-bottom: 1px solid rgba(0, 255, 157, 0.05);
    }
    .re-entry:last-child { border-bottom: none; }
    .re-date { color: #666; width: 75px; flex-shrink: 0; }
    .re-minion { color: #00f2ff; font-weight: bold; }
    .re-arrow { color: #ff00ff; }
    .re-boss { color: #ff6600; }
    .re-sector { color: #555; margin-left: auto; font-size: 0.9em; }

    /* Achievement badges on HUD */
    .badge-showcase {
        margin-top: 15px;
        border: 1px solid rgba(255, 234, 0, 0.3);
        background: rgba(255, 234, 0, 0.02);
        border-radius: 6px;
        padding: 12px 15px;
    }
    .badge-showcase h2 {
        color: #ffea00;
        font-size: 0.85em;
        letter-spacing: 3px;
        margin: 0 0 10px 0;
        text-align: center;
        text-shadow: 0 0 8px rgba(255, 234, 0, 0.4);
    }
    .badge-row {
        display: flex;
        flex-wrap: wrap;
        justify-content: center;
        gap: 12px;
    }
    .badge-item {
        display: flex;
        flex-direction: column;
        align-items: center;
        width: 72px;
    }
    .badge-icon { font-size: 1.8em; filter: drop-shadow(0 0 6px currentColor); }
    .badge-name { font-size: 0.55em; color: #ffea00; text-align: center; margin-top: 4px; line-height: 1.2; }
    .badge-date { font-size: 0.45em; color: #666; }
    .badge-new { animation: badgePulse 1.5s ease-in-out infinite; }
    @keyframes badgePulse { 0%, 100% { transform: scale(1); } 50% { transform: scale(1.15); } }
    .badge-view-all {
        display: block;
        text-align: center;
        margin-top: 10px;
        color: #ffea00;
        font-size: 0.7em;
        text-decoration: none;
        letter-spacing: 2px;
    }
    .badge-view-all:hover { text-decoration: underline; }

    /* Rejected quest alerts */
    .rejected-alerts {
        margin-top: 15px;
        border: 1px solid rgba(255, 0, 68, 0.4);
        background: rgba(255, 0, 68, 0.03);
        border-radius: 6px;
        padding: 10px 15px;
    }
    .rejected-alerts h2 {
        color: #ff0044;
        font-size: 0.8em;
        letter-spacing: 2px;
        margin: 0 0 8px 0;
        text-align: center;
    }
    .rejected-entry {
        font-size: 0.75em;
        padding: 4px 0;
        border-bottom: 1px solid rgba(255, 0, 68, 0.1);
        color: #ff0044;
    }
    .rejected-entry:last-child { border-bottom: none; }
    .rejected-reason { color: #ff6666; font-size: 0.9em; text-transform: none; }

    /* Nav badges button */
    .nav-badges { color: #ffea00; border-color: #ffea00; box-shadow: 0 0 8px rgba(255,234,0,0.2); }
    .nav-badges:hover { background: #ffea00; color: #0a0b10; box-shadow: 0 0 15px rgba(255,234,0,0.5); }

    /* Streak display */
    .streak-display {
        margin-top: 15px;
        border: 1px solid rgba(255, 136, 0, 0.3);
        background: rgba(255, 136, 0, 0.02);
        border-radius: 6px;
        padding: 10px 15px;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 10px;
        flex-wrap: wrap;
    }
    .streak-fire { font-size: 1.6em; }
    .streak-count { font-size: 1.8em; color: #ff8800; font-weight: bold; text-shadow: 0 0 10px rgba(255,136,0,0.5); }
    .streak-label { font-size: 0.85em; color: #ff8800; letter-spacing: 2px; }
    .streak-best { font-size: 0.65em; color: #666; letter-spacing: 1px; }
    .streak-inactive { opacity: 0.4; }
    .streak-inactive .streak-label { color: #555; }
    .streak-record { border-color: rgba(255, 136, 0, 0.6); box-shadow: 0 0 12px rgba(255,136,0,0.15); }
    .streak-new-record {
        font-size: 0.6em;
        color: #ffea00;
        letter-spacing: 2px;
        animation: badgePulse 1.5s ease-in-out infinite;
    }

    /* Sector map */
    .sector-map {
        margin-top: 10px;
        padding: 15px 10px 40px 10px;
        text-align: center;
    }
    .sector-map h1 {
        font-size: 1.8em;
        margin-bottom: 20px;
        border-bottom: none !important;
        padding-bottom: 0;
        text-shadow: 0 0 10px rgba(0, 242, 255, 0.5);
    }
    #bosses {
        columns: 2;
        column-gap: 20px;
    }
    .sector-zone {
        border: 1px solid rgba(0, 242, 255, 0.3);
        background: rgba(0, 242, 255, 0.03);
        border-radius: 8px;
        padding: 20px 12px 12px 12px;
        position: relative;
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 8px;
        break-inside: avoid;
        margin-bottom: 20px;
    }
    .sector-zone::before {
        content: attr(data-sector);
        position: absolute;
        top: -12px;
        left: 50%;
        transform: translateX(-50%);
        background: #0a0b10;
        padding: 2px 15px;
        font-size: 0.85em;
        color: #ff00ff;
        border: 1px solid #ff00ff;
        letter-spacing: 2px;
        white-space: nowrap;
        z-index: 10;
        box-shadow: 0 0 8px rgba(255, 0, 255, 0.4);
        pointer-events: none;
        visibility: hidden;
    }
    .sector-link {
        position: absolute;
        top: -12px;
        left: 50%;
        transform: translateX(-50%);
        background: #0a0b10;
        padding: 2px 15px;
        font-size: 0.85em;
        color: #ff00ff;
        border: 1px solid #ff00ff;
        letter-spacing: 2px;
        white-space: nowrap;
        z-index: 11;
        box-shadow: 0 0 8px rgba(255, 0, 255, 0.4);
        text-decoration: none;
        cursor: pointer;
        transition: all 0.2s;
    }
    .sector-link:hover {
        background: #ff00ff;
        color: #0a0b10;
        box-shadow: 0 0 16px rgba(255, 0, 255, 0.8);
    }
    .sector-radar {
        margin-bottom: 5px;
    }
    .sector-bosses {
        display: flex;
        flex-wrap: wrap;
        justify-content: center;
        gap: 10px 14px;
        width: 100%;
    }
    .boss-link {
        text-decoration: none;
        color: inherit;
    }
    .boss-orb-container {
        display: flex;
        align-items: center;
        gap: 6px;
        position: relative;
        cursor: pointer;
    }
    .boss-pie {
        border-radius: 50%;
        border: 2px solid #00f2ff;
        box-shadow: 0 0 10px rgba(0, 242, 255, 0.4);
        transition: transform 0.2s ease;
    }
    .boss-pie:hover {
        transform: scale(1.2);
        box-shadow: 0 0 15px #00f2ff;
    }
    .boss-label {
        font-size: 0.65em;
        color: #00f2ff;
        text-align: left;
        word-wrap: break-word;
        line-height: 1.2;
        max-width: 80px;
    }
    .boss-key {
        display: flex;
        justify-content: center;
        gap: 25px;
        margin-bottom: 20px;
        font-size: 0.75em;
    }
    .key-item { display: flex; align-items: center; gap: 6px; }
    .key-swatch { display: inline-block; width: 12px; height: 12px; border-radius: 50%; }

    /* Power Rankings */
    .power-rankings {
        margin-top: 10px;
        padding: 15px 10px 20px 10px;
    }
    .power-rankings h1 {
        font-size: 1.4em;
        color: #ffd700;
        text-shadow: 2px 2px #ff00ff;
        border-bottom: none !important;
        margin-bottom: 15px;
    }
    .rank-entry {
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 8px;
        font-size: 0.8em;
    }
    .rank-badge {
        font-weight: bold;
        width: 30px;
        text-align: right;
        flex-shrink: 0;
    }
    .rank-bar-wrap {
        flex: 1;
        background: #1a1d26;
        border: 1px solid #333;
        height: 22px;
        position: relative;
        overflow: hidden;
    }
    .rank-bar {
        height: 100%;
        opacity: 0.7;
        transition: width 0.5s ease-in-out;
    }
    .rank-info {
        position: absolute;
        top: 2px;
        left: 8px;
        font-size: 0.85em;
        color: #fff;
        text-shadow: 1px 1px 2px #000;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        max-width: calc(100% - 16px);
    }
    .rank-boss {
        color: #ccc;
        font-size: 0.85em;
        font-weight: bold;
        width: 110px;
        flex-shrink: 0;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .rank-sector {
        width: 80px;
        text-align: center;
        font-size: 0.7em;
        color: #666;
        flex-shrink: 0;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .rank-status-label {
        width: 120px;
        text-align: right;
        font-size: 0.7em;
        font-weight: bold;
        letter-spacing: 1px;
        flex-shrink: 0;
    }
    .rank-subtitle {
        text-align: center;
        font-size: 0.65em;
        color: #ff6644;
        letter-spacing: 2px;
        margin-bottom: 15px;
        text-shadow: 0 0 6px rgba(255, 102, 68, 0.3);
    }
    a.rank-entry {
        transition: background 0.2s;
        padding: 2px 4px;
        border-radius: 3px;
    }
    a.rank-entry:hover {
        background: rgba(255, 255, 255, 0.05);
    }

    /* Side navigation panel */
    .side-nav {
        position: fixed;
        top: 50%;
        left: 0;
        transform: translateY(-50%);
        display: flex;
        flex-direction: column;
        gap: 4px;
        z-index: 1000;
        padding: 6px;
    }
    .side-nav a {
        display: block;
        text-decoration: none;
        background: #0a0b10;
        padding: 8px 14px;
        font-family: 'Courier New', monospace;
        font-size: 0.7em;
        font-weight: bold;
        letter-spacing: 2px;
        text-transform: uppercase;
        transition: all 0.3s;
        white-space: nowrap;
        border: 1px solid;
    }
    .side-nav a:hover {
        padding-left: 20px;
    }
    .nav-army { color: #00ff9d; border-color: #00ff9d; box-shadow: 0 0 8px rgba(0,255,157,0.2); }
    .nav-army:hover { background: #00ff9d; color: #0a0b10; box-shadow: 0 0 15px rgba(0,255,157,0.5); }
    .nav-progress { color: #ffea00; border-color: #ffea00; box-shadow: 0 0 8px rgba(255,234,0,0.2); }
    .nav-progress:hover { background: #ffea00; color: #0a0b10; box-shadow: 0 0 15px rgba(255,234,0,0.5); }
    .nav-quests { color: #ff6600; border-color: #ff6600; box-shadow: 0 0 8px rgba(255,102,0,0.2); }
    .nav-quests:hover { background: #ff6600; color: #0a0b10; box-shadow: 0 0 15px rgba(255,102,0,0.5); }
    .nav-quests.pulse { animation: pulse 2s infinite; }
    .nav-quests.dim { border-color: #333; color: #666; box-shadow: none; animation: none; }
    .nav-admin { color: #ff00ff; border-color: #ff00ff; box-shadow: 0 0 8px rgba(255,0,255,0.2); }
    .nav-admin:hover { background: #ff00ff; color: #0a0b10; box-shadow: 0 0 15px rgba(255,0,255,0.5); }
    .nav-user { color: #555; border-color: #333; box-shadow: none; font-size: 0.6em !important; letter-spacing: 1px !important; padding: 5px 10px !important; margin-top: 4px; }
    .nav-user:hover { border-color: #888; color: #888; background: transparent; }

    /* Quest UI */
    .quest-badge-link { text-decoration: none; }
    .quest-badge {
        display: inline-block;
        background: #ff6600;
        color: #0a0b10;
        font-size: 0.45em;
        padding: 3px 10px;
        vertical-align: middle;
        margin-left: 10px;
        letter-spacing: 1px;
        font-weight: bold;
        animation: pulse 2s infinite;
    }
    .quest-badge.dim { background: #333; color: #666; animation: none; }

    @keyframes pulse { 0% { opacity: 1; } 50% { opacity: 0.7; } 100% { opacity: 1; } }
    .glitch { font-size: 0.7em; color: #555; margin-top: 20px; text-align: center; line-height: 1.6; }

    /* Survival Mode — Ring of Guardians */
    .survival-ring {
        text-align: center;
        padding: 20px 10px;
        margin-bottom: 20px;
        border: 1px solid rgba(255, 68, 68, 0.3);
        border-radius: 8px;
        background: rgba(255, 68, 68, 0.03);
    }
    .survival-ring.survival-achieved {
        border-color: rgba(255, 215, 0, 0.5);
        background: rgba(255, 215, 0, 0.03);
        box-shadow: 0 0 20px rgba(255, 215, 0, 0.15);
    }
    .mode-header {
        padding: 10px 0 0;
        margin-top: 10px;
        position: relative;
    }
    .mode-header::before {
        content: "";
        display: block;
        height: 1px;
        margin-bottom: 10px;
        background: linear-gradient(90deg, transparent, rgba(255,234,0,0.4) 30%, rgba(0,242,255,0.4) 70%, transparent);
    }
    .mode-header-achieved {
        border-bottom-color: rgba(255, 215, 0, 0.3);
    }
    .mode-banner-line {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 10px;
        flex-wrap: wrap;
        margin-bottom: 4px;
    }
    .mode-banner {
        font-size: 0.95em;
        font-weight: bold;
        letter-spacing: 3px;
    }
    .mode-creative {
        color: #00f2ff;
        text-shadow: 0 0 10px rgba(0, 242, 255, 0.5);
    }
    .mode-survival {
        color: #ffd700;
        text-shadow: 0 0 15px rgba(255, 215, 0, 0.6);
        animation: pulse 2s infinite;
    }
    .mode-subtitle {
        font-size: 0.7em;
        font-weight: bold;
        text-align: center;
        letter-spacing: 2px;
        color: #ff6644;
        text-shadow: 0 0 8px rgba(255, 102, 68, 0.4);
    }
    .mode-achieved {
        color: #ffd700;
        text-shadow: 0 0 8px rgba(255, 215, 0, 0.4);
    }
    .hearts-row {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        flex-wrap: wrap;
    }
    .guardian-heart {
        cursor: help;
        transition: transform 0.2s;
    }
    .guardian-heart:hover {
        transform: scale(1.3);
    }
    .survival-ring-title {
        font-size: 0.9em;
        color: #ff4444;
        letter-spacing: 3px;
        margin-bottom: 15px;
        text-shadow: 0 0 8px rgba(255, 68, 68, 0.5);
        text-decoration: none;
        display: block;
        cursor: pointer;
        transition: color 0.2s, text-shadow 0.2s;
    }
    .survival-ring-title:hover {
        color: #ffea00;
        text-shadow: 0 0 12px rgba(255, 234, 0, 0.6);
    }
    .survival-achieved .survival-ring-title {
        color: #ffd700;
        text-shadow: 0 0 8px rgba(255, 215, 0, 0.5);
    }
    .survival-achieved .survival-ring-title:hover {
        color: #ffea00;
        text-shadow: 0 0 12px rgba(255, 234, 0, 0.8);
    }
    .guardian-ring {
        display: flex;
        flex-wrap: wrap;
        justify-content: center;
        gap: 25px;
        padding: 10px;
    }
    .guardian-orb {
        display: flex;
        flex-direction: column;
        align-items: center;
        width: 80px;
        text-decoration: none;
        color: inherit;
        cursor: pointer;
    }
    .guardian-orb-circle {
        width: 50px;
        height: 50px;
        border-radius: 50%;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .guardian-orb:hover .guardian-orb-circle {
        transform: scale(1.2);
        box-shadow: 0 0 20px #ffea00, 0 0 40px rgba(255, 234, 0, 0.4) !important;
        border-color: #ffea00 !important;
    }
    .guardian-orb:hover .guardian-label {
        color: #ffea00;
    }
    .guardian-orb.complete .guardian-orb-circle {
        box-shadow: 0 0 15px #00ff9d, 0 0 30px rgba(0, 255, 157, 0.3);
        border: 2px solid #00ff9d;
    }
    .guardian-orb.in-progress .guardian-orb-circle {
        box-shadow: 0 0 12px #ff6600, 0 0 20px rgba(255, 102, 0, 0.3);
        border: 2px solid #ff6600;
    }
    .guardian-orb.not-started .guardian-orb-circle {
        box-shadow: 0 0 5px rgba(255, 68, 68, 0.2);
        border: 2px solid #333;
    }
    .guardian-label {
        font-size: 0.6em;
        color: #aaa;
        margin-top: 6px;
        text-align: center;
        word-wrap: break-word;
        line-height: 1.2;
        width: 100%;
    }
    .guardian-fraction {
        font-size: 0.55em;
        color: #666;
        margin-top: 2px;
    }
    .survival-summary {
        font-size: 0.7em;
        color: #666;
        letter-spacing: 2px;
        margin-top: 10px;
    }

    @media (max-width: 600px) {
        body { padding: 10px; }
        .hud-container { padding: 15px; }
        .confidence-row { flex-direction: column; }
        .levels-right {
            padding-left: 0;
            padding-top: 15px;
        }
        .title-main span { font-size: 1.2em; letter-spacing: 3px; }
        .title-tagline { font-size: 0.75em; letter-spacing: 3px; }
        .stat-grid { grid-template-columns: 1fr; }
        #bosses { columns: 1; }
        .guardian-ring { gap: 15px; }
        .guardian-orb { width: 65px; }
        .guardian-orb-circle { width: 40px; height: 40px; }
        .side-nav {
            position: fixed;
            top: auto;
            bottom: 0;
            left: 0;
            right: 0;
            transform: none;
            flex-direction: row;
            justify-content: center;
            background: rgba(10,11,16,0.95);
            border-top: 1px solid #333;
            padding: 6px 4px;
        }
        .side-nav a { font-size: 0.6em; padding: 6px 8px; letter-spacing: 1px; }
        .side-nav a:hover { padding-left: 8px; }
    }
    </style>
</head>
<body>
    <div class="hud-container">
        <div class="hud-title-block">
            <div class="title-tagline">Path to Independence</div>
            <h1 class="title-main"><span>The Henry HUD</span></h1>
        </div>

        <div class="confidence-row">
            <div class="radar-main">[[MAIN_RADAR]]</div>
            <div class="confidence-bar-section">
                <div class="stat-label" style="color: #ffea00; font-size: 1.1em; text-align: center;">
                    CONFIDENCE: [[CONF_RANK]]
                    <div class="stat-next">+[[CONF_REM]] for [[CONF_NEXT]]</div>
                </div>
                <div class="stat-bar">
                    <div class="fill confidence-fill" style="width: [[CONF_BAR]]%;"></div>
                </div>
                [[SURVIVAL_MODE_HEADER]]
            </div>
            <div class="levels-right">
                [[TIER_LIST]]
            </div>
        </div>

        <div class="stat-grid">
            <div class="stat-container">
                <div class="stat-label" style="color: #00f2ff;">INTEL: [[INTEL_RANK]] <span class="stat-next">+[[INTEL_REM]] for [[INTEL_NEXT]]</span></div>
                <div class="stat-bar"><div class="fill intel-fill" style="width: [[INTEL_BAR]]%;"></div></div>
            </div>
            <div class="stat-container">
                <div class="stat-label" style="color: #00ff9d;">STAMINA: [[STAMINA_RANK]] <span class="stat-next">+[[STAMINA_REM]] for [[STAMINA_NEXT]]</span></div>
                <div class="stat-bar"><div class="fill stamina-fill" style="width: [[STAMINA_BAR]]%;"></div></div>
            </div>
            <div class="stat-container">
                <div class="stat-label" style="color: #ff00ff;">TEMPO: [[TEMPO_RANK]] <span class="stat-next">+[[TEMPO_REM]] for [[TEMPO_NEXT]]</span></div>
                <div class="stat-bar"><div class="fill tempo-fill" style="width: [[TEMPO_BAR]]%;"></div></div>
            </div>
            <div class="stat-container">
                <div class="stat-label" style="color: #ff8800;">REPUTATION: [[REP_RANK]] <span class="stat-next">+[[REPUTATION_REM]] for [[REP_NEXT]]</span></div>
                <div class="stat-bar"><div class="fill rep-fill" style="width: [[REPUTATION_BAR]]%;"></div></div>
            </div>
        </div>

        [[SURVIVAL_RING]]

        [[RECENT_ENSLAVED]]

        [[STREAK_SECTION]]

        [[REJECTED_ALERTS]]

        <div class="sector-map">
            <h1>Bosses & Minions</h1>
            <div class="boss-key">
                <span class="key-item"><span class="key-swatch" style="background:#00ff9d;"></span> Enslaved</span>
                <span class="key-item"><span class="key-swatch" style="background:#ff6600;"></span> Engaged</span>
                <span class="key-item"><span class="key-swatch" style="background:#2a2d36; border: 1px solid #555;"></span> Locked</span>
            </div>
            <div id="bosses">[[BOSS_LIST]]</div>
        </div>

        <div class="power-rankings">
            <h1>Boss Conquest [[QUEST_BADGE]]</h1>
            <div class="rank-subtitle">BOSSES WITH MINIONS REMAINING &mdash; CONQUER THEM ALL</div>
            <div class="rank-entry rank-header">
                <span class="rank-badge"></span>
                <span class="rank-boss" style="color:#aaa;">BOSS</span>
                <div class="rank-bar-wrap" style="background:none;border:none;"></div>
                <span class="rank-sector" style="color:#aaa;">SECTOR</span>
                <span class="rank-status-label" style="color:#aaa;">STATUS</span>
            </div>
            [[POWER_RANKINGS]]
        </div>

        <div class="glitch">
            <span style="color: #ff00ff;">SYSTEM: ONLINE</span><br>
            > DATA_STREAM_SYNCED...<br>
            > NO_THREATS_DETECTED_IN_CORE...
        </div>
    </div>
    <nav class="side-nav">
        <a href="/army" class="nav-army">[[ARMY_LINK]]</a>
        <a href="/badges" class="nav-badges">[[BADGES_LINK]]</a>
        <a href="/progress" class="nav-progress">&#x1F4CA; PROGRESS</a>
        <a href="/quests" class="nav-quests">[[QUEST_LINK]]</a>
        [[ADMIN_NAV]]
        <a href="/login" class="nav-user">[[USER_NAV]]</a>
    </nav>
</body>
</html>`;

// ---------------------------------------------------------------------------
// Boss Detail Page Template
// ---------------------------------------------------------------------------
function buildBossPage(bossName, sector, minions, totals, activeQuestKeys, isSurvivalBoss) {
  const statusColor = { Enslaved: "#00ff9d", Engaged: "#ff6600", Locked: "#555" };
  const norm = (raw, max) => ((parseFloat(raw) || 0) / max * 100).toFixed(1);
  const survivalBadge = isSurvivalBoss
    ? `<div class="survival-badge"><svg width="14" height="16" viewBox="0 0 18 20" fill="none" style="vertical-align:middle;margin-right:3px;"><path d="M9 0L0 3.5V9C0 14 3.8 18.5 9 20C14.2 18.5 18 14 18 9V3.5L9 0Z" fill="rgba(255,68,68,0.3)" stroke="#ff4444" stroke-width="1.5"/><text x="9" y="13" text-anchor="middle" fill="#ff4444" font-size="8" font-family="monospace" font-weight="bold">S</text></svg> SURVIVAL MODE GUARDIAN</div>`
    : "";

  let rows = "";
  for (const m of minions) {
    const sc = statusColor[m["Status"]] || "#555";
    const nInt = norm(m["INTELLIGENCE"], totals.intel);
    const nSta = norm(m["STAMINA"], totals.stamina);
    const nTmp = norm(m["TEMPO"], totals.tempo);
    const nRep = norm(m["REPUTATION"], totals.rep);
    const nTotal = (parseFloat(nInt) + parseFloat(nSta) + parseFloat(nTmp) + parseFloat(nRep)).toFixed(1);
    const qKey = bossName + "|" + m["Minion"];
    const onQuest = activeQuestKeys && activeQuestKeys.has(qKey);
    let questBtn;
    if (onQuest) {
      questBtn = `<span style="color:#ffea00;text-shadow:0 0 6px #ffea00;" title="On quest board">&#x2605;</span>`;
    } else if (m["Status"] === "Engaged") {
      questBtn = `<input type="checkbox" class="quest-chk" data-boss="${escHtml(bossName)}" data-minion="${escHtml(m["Minion"])}" data-sector="${escHtml(sector)}" title="Select for quest board">`;
    } else {
      questBtn = `<span style="opacity:0.2;">-</span>`;
    }
    const prereqText = m["Status"] === "Locked" && m["Locked for what?"] ? m["Locked for what?"] : "";
    rows += `
      <tr>
        <td>${questBtn}</td>
        <td title="${escHtml(m["Task"] || "")}">
          ${escHtml(m["Minion"])}
          ${prereqText ? `<div style="font-size:0.75em;color:#ff0044;margin-top:2px;text-transform:none;">Requires: ${escHtml(prereqText)}</div>` : ""}
        </td>
        <td style="color:${sc}; font-weight:bold;">${m["Status"]}</td>
        <td>${nInt}</td>
        <td>${nSta}</td>
        <td>${nTmp}</td>
        <td>${nRep}</td>
        <td>${m["Impact(1-3)"] || ""}</td>
        <td style="color:#ffea00; font-weight:bold;">${nTotal}</td>
      </tr>`;
  }

  return `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>${bossName} - Minion Roster</title>
    <style>
    body {
        background: #0a0b10;
        color: #00f2ff;
        font-family: 'Courier New', monospace;
        padding: 20px;
        text-transform: uppercase;
        overflow-x: auto;
    }
    .hud-container {
        border: 2px solid #00f2ff;
        padding: 20px;
        box-shadow: 0 0 15px #00f2ff;
        max-width: 960px;
        margin: auto;
    }
    h1 {
        text-shadow: 2px 2px #ff00ff;
        border-bottom: 1px solid #00f2ff;
        letter-spacing: 2px;
        text-align: center;
        margin-bottom: 5px;
    }
    .sector-tag {
        text-align: center;
        color: #ff00ff;
        font-size: 0.85em;
        letter-spacing: 2px;
        margin-bottom: 20px;
    }
    .survival-badge {
        text-align: center;
        color: #ff4444;
        font-size: 0.7em;
        letter-spacing: 3px;
        margin-bottom: 5px;
        text-shadow: 0 0 8px rgba(255, 68, 68, 0.4);
    }
    .back-link {
        display: inline-block;
        color: #00f2ff;
        text-decoration: none;
        border: 1px solid #00f2ff;
        padding: 6px 15px;
        margin-bottom: 15px;
        font-size: 0.8em;
        transition: all 0.2s;
    }
    .back-link:hover { background: #00f2ff; color: #0a0b10; }
    table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.8em;
    }
    th {
        background: #1a1d26;
        color: #ffea00;
        padding: 10px 8px;
        text-align: left;
        border-bottom: 2px solid #00f2ff;
    }
    td {
        padding: 8px;
        border-bottom: 1px solid #1a1d26;
    }
    tr:hover td { background: rgba(0, 242, 255, 0.05); }
    .quest-add-btn {
        background: none;
        border: 1px solid #ff6600;
        color: #ff6600;
        width: 24px;
        height: 24px;
        font-size: 1.1em;
        font-weight: bold;
        cursor: pointer;
        font-family: 'Courier New', monospace;
        transition: all 0.2s;
        padding: 0;
        line-height: 22px;
    }
    .quest-add-btn:hover { background: #ff6600; color: #0a0b10; box-shadow: 0 0 8px rgba(255, 102, 0, 0.5); }
    td[title]:not([title=""]) { cursor: help; border-bottom: 1px dotted #555; }
    .quest-chk { width: 16px; height: 16px; accent-color: #ff6600; cursor: pointer; }
    .quest-batch-bar {
        display: flex; align-items: center; justify-content: center; gap: 15px;
        margin-top: 15px; padding: 10px; border: 1px solid #ff6600;
        background: rgba(255, 102, 0, 0.08);
    }
    .quest-batch-count { color: #ffea00; font-weight: bold; font-size: 0.85em; letter-spacing: 2px; }
    .quest-batch-btn {
        background: #ff6600; color: #0a0b10; border: none; padding: 8px 18px;
        font-family: 'Courier New', monospace; font-weight: bold; font-size: 0.85em;
        cursor: pointer; letter-spacing: 1px; transition: all 0.2s;
    }
    .quest-batch-btn:hover { background: #ffea00; box-shadow: 0 0 10px rgba(255, 234, 0, 0.5); }
    </style>
</head>
<body>
    <div class="hud-container">
        <a class="back-link" href="/">&lt; BACK TO HUD</a>
        <h1>${bossName}</h1>
        ${survivalBadge}
        <div class="sector-tag">SECTOR: ${sector}</div>
        <table>
            <thead>
                <tr>
                    <th></th>
                    <th>Minion</th>
                    <th>Status</th>
                    <th>INT</th>
                    <th>STA</th>
                    <th>TMP</th>
                    <th>REP</th>
                    <th>Impact</th>
                    <th>Total</th>
                </tr>
            </thead>
            <tbody>${rows}</tbody>
        </table>
        <div class="quest-batch-bar" style="display:none;">
            <span class="quest-batch-count">0 SELECTED</span>
            <button type="button" class="quest-batch-btn">ADD SELECTED TO QUEST BOARD</button>
        </div>
        <div style="text-align:center;margin-top:15px;font-size:0.75em;color:#ff6600;">SELECT ENGAGED MINIONS TO ADD TO YOUR QUEST BOARD</div>
    </div>
    <script>
    (function() {
        const bar = document.querySelector('.quest-batch-bar');
        const countEl = bar.querySelector('.quest-batch-count');
        const btn = bar.querySelector('.quest-batch-btn');
        const checkboxes = document.querySelectorAll('.quest-chk');
        function updateBar() {
            const checked = document.querySelectorAll('.quest-chk:checked');
            if (checked.length > 0) {
                bar.style.display = 'flex';
                countEl.textContent = checked.length + ' SELECTED';
            } else {
                bar.style.display = 'none';
            }
        }
        checkboxes.forEach(function(chk) { chk.addEventListener('change', updateBar); });
        btn.addEventListener('click', function() {
            const checked = document.querySelectorAll('.quest-chk:checked');
            if (checked.length === 0) return;
            var items = [];
            checked.forEach(function(chk) {
                items.push({ boss: chk.dataset.boss, minion: chk.dataset.minion, sector: chk.dataset.sector });
            });
            var form = document.createElement('form');
            form.method = 'POST';
            form.action = '/quest/start-batch';
            var input = document.createElement('input');
            input.type = 'hidden';
            input.name = 'items';
            input.value = JSON.stringify(items);
            var redir = document.createElement('input');
            redir.type = 'hidden';
            redir.name = 'redirect';
            redir.value = window.location.pathname + window.location.search;
            form.appendChild(input);
            form.appendChild(redir);
            document.body.appendChild(form);
            form.submit();
        });
    })();
    </script>
</body>
</html>`;
}

// ---------------------------------------------------------------------------
// AI Import: Fetch context from Sheets for the AI prompt
// ---------------------------------------------------------------------------
async function fetchImportContext(sheets) {
  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["Sectors", "Definitions"],
  });
  const [sectorsRaw, defsRaw] = res.data.valueRanges;
  const sectors = parseTable(sectorsRaw.values);
  const definitions = parseTable(defsRaw.values);
  const headers = sectorsRaw.values ? sectorsRaw.values[0] : [];

  const bossMap = {};
  const minionSet = new Set();
  for (const row of sectors) {
    const s = row["Sector"], b = row["Boss"], m = row["Minion"];
    if (!s || !b) continue;
    if (!bossMap[s]) bossMap[s] = new Set();
    bossMap[s].add(b);
    if (m) minionSet.add(m.toLowerCase());
  }

  const validSectors = definitions.filter((d) => d["Sector"]).map((d) => d["Sector"]);
  return { sectors, definitions, bossMap, minionSet, validSectors, headers };
}

// ---------------------------------------------------------------------------
// AI Import: Build the Claude vision prompt
// ---------------------------------------------------------------------------
function buildImportPrompt(context) {
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
// AI Import: Analyze a single image with Claude vision
// ---------------------------------------------------------------------------
async function analyzeImageAI(client, imagePath, prompt) {
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
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) throw new Error("No JSON found in response");
  return {
    result: JSON.parse(jsonMatch[0]),
    usage: { input: response.usage.input_tokens, output: response.usage.output_tokens },
  };
}

// ---------------------------------------------------------------------------
// AI Import: Build a row array matching the Sectors sheet header order
// ---------------------------------------------------------------------------
function buildSectorsRow(headers, data) {
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
    "Date Quest Added": "",
    "Date Quest Completed": "",
  };
  for (const stat of statNames) {
    valueMap[stat] = statFormula(stat);
  }
  return headers.map((h) => valueMap[h] ?? "");
}

// ---------------------------------------------------------------------------
// Express server
// ---------------------------------------------------------------------------
const app = express();
app.set("trust proxy", 1); // trust first proxy (nginx/Caddy on Linode)
app.use(express.urlencoded({ extended: false }));
app.use(express.json());
app.use(cookieParser());

// ---------------------------------------------------------------------------
// Rate limiting (in-memory, no extra dependency)
// ---------------------------------------------------------------------------
const rateLimits = {};
const RATE_WINDOW = 60 * 1000; // 1 minute
const RATE_MAX = 60; // requests per window per IP

app.use((req, res, next) => {
  const ip = req.ip;
  const now = Date.now();
  if (!rateLimits[ip] || now - rateLimits[ip].start > RATE_WINDOW) {
    rateLimits[ip] = { start: now, count: 1 };
  } else {
    rateLimits[ip].count++;
  }
  if (rateLimits[ip].count > RATE_MAX) {
    return res.status(429).send("Too many requests. Try again in a minute.");
  }
  next();
});

// Clean up stale rate limit entries every 5 minutes
setInterval(() => {
  const cutoff = Date.now() - RATE_WINDOW;
  for (const ip in rateLimits) {
    if (rateLimits[ip].start < cutoff) delete rateLimits[ip];
  }
}, 5 * 60 * 1000).unref();

// Ensure import directories exist
fs.mkdirSync(IMPORT_DIR, { recursive: true });
fs.mkdirSync(DONE_DIR, { recursive: true });

const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, IMPORT_DIR),
    filename: (req, file, cb) => {
      const unique = Date.now() + "-" + Math.round(Math.random() * 1e6);
      cb(null, unique + path.extname(file.originalname).toLowerCase());
    },
  }),
  fileFilter: (req, file, cb) => {
    const allowed = new Set([".jpg", ".jpeg", ".png", ".webp"]);
    cb(null, allowed.has(path.extname(file.originalname).toLowerCase()));
  },
  limits: { fileSize: 20 * 1024 * 1024 },
});

// ---------------------------------------------------------------------------
// Login / Logout / User identification
// ---------------------------------------------------------------------------
app.get("/login", async (req, res) => {
  try {
    const sheets = await getSheets();
    await ensureUsersSheet(sheets);
    const usersRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Users",
    });
    const users = parseTable(usersRes.data.values || []);
    const currentEmail = req.cookies.userEmail || "";

    const userOptions = users.map((u) => {
      const selected = u["Email"].toLowerCase() === currentEmail.toLowerCase() ? " selected" : "";
      return `<option value="${escHtml(u["Email"])}"${selected}>${escHtml(u["Name"] || u["Email"])} (${escHtml(u["Role"] || "??")})</option>`;
    }).join("");
    const hasUsers = users.length > 0;
    const successMsg = req.query.added === "1" ? `<div style="color:#00ff9d;font-size:0.75em;margin-bottom:15px;border:1px solid rgba(0,255,157,0.3);padding:8px;">&#x2714; USER ADDED</div>` : "";

    res.send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Login - Sovereign HUD</title>
    <style>
    body { background: #0a0b10; color: #00f2ff; font-family: 'Courier New', monospace; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; text-transform: uppercase; }
    .login-box { border: 2px solid #ffea00; padding: 40px; box-shadow: 0 0 30px rgba(255,234,0,0.3); text-align: center; max-width: 450px; width: 90%; }
    h1 { color: #ffea00; text-shadow: 2px 2px #ff00ff; letter-spacing: 4px; margin: 0 0 10px; }
    .subtitle { color: #888; font-size: 0.7em; letter-spacing: 3px; margin-bottom: 25px; }
    .form-group { margin-bottom: 15px; text-align: left; }
    .form-group label { display: block; font-size: 0.7em; color: #ffea00; letter-spacing: 2px; margin-bottom: 5px; }
    select, input[type="text"], input[type="email"] { width: 100%; padding: 12px; background: #1a1d26; border: 1px solid #333; color: #00f2ff; font-family: 'Courier New', monospace; font-size: 0.9em; box-sizing: border-box; }
    select:focus, input:focus { outline: none; border-color: #ffea00; box-shadow: 0 0 5px rgba(255,234,0,0.3); }
    .login-btn { width: 100%; padding: 14px; background: #ffea00; color: #0a0b10; border: none; font-family: 'Courier New', monospace; font-size: 1em; font-weight: bold; letter-spacing: 3px; cursor: pointer; transition: all 0.2s; }
    .login-btn:hover { background: #00ff9d; box-shadow: 0 0 15px rgba(0,255,157,0.5); }
    .add-btn { width: 100%; padding: 12px; background: transparent; color: #00f2ff; border: 1px solid #00f2ff; font-family: 'Courier New', monospace; font-size: 0.85em; font-weight: bold; letter-spacing: 2px; cursor: pointer; transition: all 0.2s; }
    .add-btn:hover { background: #00f2ff; color: #0a0b10; }
    .current-user { margin-top: 20px; font-size: 0.7em; color: #555; letter-spacing: 1px; }
    .current-user a { color: #ff4444; text-decoration: none; }
    .current-user a:hover { text-decoration: underline; }
    .skip-link { display: block; margin-top: 15px; font-size: 0.7em; color: #555; text-decoration: none; letter-spacing: 1px; }
    .skip-link:hover { color: #00f2ff; }
    .divider { border: none; border-top: 1px solid #333; margin: 25px 0; }
    .section-label { font-size: 0.65em; color: #888; letter-spacing: 2px; margin-bottom: 15px; }
    .setup-note { font-size: 0.65em; color: #ff8800; letter-spacing: 1px; margin-bottom: 15px; text-transform: none; }
    </style>
</head>
<body>
    <div class="login-box">
        <h1>&#x1F512; IDENTIFY</h1>
        <div class="subtitle">WHO IS ACCESSING THE HUD?</div>
        ${successMsg}
        ${hasUsers ? `
        <form method="POST" action="/login">
            <div class="form-group">
                <label>SELECT USER</label>
                <select name="email" required>
                    <option value="" disabled ${currentEmail ? "" : "selected"}>Choose your identity...</option>
                    ${userOptions}
                </select>
            </div>
            <button type="submit" class="login-btn">ENTER HUD</button>
        </form>
        ${currentEmail ? `<div class="current-user">LOGGED IN AS: ${escHtml(currentEmail)} &mdash; <a href="/logout">SWITCH USER</a></div>` : ""}
        ` : `
        <div class="setup-note">No users found. Create your first account to get started:</div>
        <form method="POST" action="/login/add-user">
            <div class="form-group">
                <label>NAME</label>
                <input type="text" name="name" required placeholder="e.g. Hyrum">
            </div>
            <div class="form-group">
                <label>EMAIL</label>
                <input type="email" name="email" required placeholder="e.g. hyrum.0@gmail.com" style="text-transform:none;">
            </div>
            <div class="form-group">
                <label>ROLE</label>
                <select name="role" required>
                    <option value="teacher" selected>Teacher (Admin Access)</option>
                    <option value="student">Student</option>
                </select>
            </div>
            <button type="submit" class="add-btn">CREATE ACCOUNT &amp; ENTER</button>
        </form>
        `}
    </div>
</body>
</html>`);
  } catch (err) {
    console.error("Login page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// Step 1: User selects identity → generate code → email it → show verify page
app.post("/login", async (req, res) => {
  const email = (req.body.email || "").trim();
  if (!email) return res.redirect("/login");

  try {
    const sheets = await getSheets();
    await ensureUsersSheet(sheets);

    // Find user in sheet
    const usersRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Users",
    });
    const users = parseTable(usersRes.data.values || []);
    const user = users.find((u) => (u["Email"] || "").toLowerCase() === email.toLowerCase());
    if (!user) return res.redirect("/login");

    const userName = user["Name"] || email;

    // Generate code and store in memory only (never on the sheet)
    const code = generateVerifyCode(email);

    // Try to email the code via Apps Script
    const emailSent = await sendVerifyEmail(email, code, userName);

    // Show verification page
    res.send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Verify - Sovereign HUD</title>
    <style>
    body { background: #0a0b10; color: #00f2ff; font-family: 'Courier New', monospace; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; text-transform: uppercase; }
    .login-box { border: 2px solid #ffea00; padding: 40px; box-shadow: 0 0 30px rgba(255,234,0,0.3); text-align: center; max-width: 450px; width: 90%; }
    h1 { color: #ffea00; text-shadow: 2px 2px #ff00ff; letter-spacing: 4px; margin: 0 0 10px; }
    .subtitle { color: #888; font-size: 0.7em; letter-spacing: 2px; margin-bottom: 20px; }
    .user-badge { color: #00ff9d; font-size: 1.1em; font-weight: bold; letter-spacing: 2px; margin-bottom: 20px; }
    .instructions { color: #ccc; font-size: 0.75em; text-transform: none; line-height: 1.8; margin-bottom: 20px; text-align: left; padding: 15px; border: 1px solid rgba(0,242,255,0.2); background: rgba(0,242,255,0.03); }
    .form-group { margin-bottom: 15px; }
    .form-group label { display: block; font-size: 0.7em; color: #ffea00; letter-spacing: 2px; margin-bottom: 5px; }
    .code-input { width: 200px; padding: 15px; background: #1a1d26; border: 2px solid #ffea00; color: #00ff9d; font-family: 'Courier New', monospace; font-size: 1.5em; text-align: center; letter-spacing: 8px; box-sizing: border-box; }
    .code-input:focus { outline: none; box-shadow: 0 0 15px rgba(255,234,0,0.4); }
    .login-btn { width: 100%; padding: 14px; background: #ffea00; color: #0a0b10; border: none; font-family: 'Courier New', monospace; font-size: 1em; font-weight: bold; letter-spacing: 3px; cursor: pointer; transition: all 0.2s; margin-top: 10px; }
    .login-btn:hover { background: #00ff9d; box-shadow: 0 0 15px rgba(0,255,157,0.5); }
    .back-link { display: block; margin-top: 15px; font-size: 0.7em; color: #555; text-decoration: none; letter-spacing: 1px; }
    .back-link:hover { color: #00f2ff; }
    .error-msg { color: #ff4444; font-size: 0.75em; margin-bottom: 10px; border: 1px solid rgba(255,68,68,0.3); padding: 8px; }
    .test-code { color: #ff8800; font-size: 1.3em; letter-spacing: 6px; margin: 10px 0; padding: 10px; border: 2px dashed #ff8800; }
    </style>
</head>
<body>
    <div class="login-box">
        <h1>&#x1F50D; VERIFY</h1>
        <div class="subtitle">PROVE YOUR IDENTITY</div>
        <div class="user-badge">${escHtml(userName)}</div>
        ${req.query.error === "1" ? '<div class="error-msg">&#x2717; INCORRECT CODE. TRY AGAIN.</div>' : ""}
        ${emailSent ? `
        <div class="instructions">
            <strong>A 6-digit verification code has been emailed to:</strong><br>
            <span style="color:#ffea00;">${escHtml(email)}</span><br><br>
            Check your inbox for the code, then enter it below.<br>
            <span style="color:#ff8800;">Tip: If you don't see it, check Gmail's "All Mail" folder &mdash; it may skip your inbox.</span><br>
            The code expires in <strong>10 minutes</strong>.
        </div>
        ` : `
        <div class="instructions" style="border-color: rgba(255,136,0,0.4); background: rgba(255,136,0,0.05);">
            <strong style="color:#ff8800;">Email not configured &mdash; showing code for testing:</strong>
        </div>
        <div class="test-code">${code}</div>
        `}
        <form method="POST" action="/login/verify">
            <input type="hidden" name="email" value="${escHtml(email)}">
            <div class="form-group">
                <label>ENTER YOUR 6-DIGIT CODE</label>
                <input type="text" name="code" class="code-input" maxlength="6" pattern="[0-9]{6}" required autofocus placeholder="------">
            </div>
            <button type="submit" class="login-btn">VERIFY &amp; ENTER</button>
        </form>
        <a class="back-link" href="/login">&larr; BACK TO USER SELECT</a>
    </div>
</body>
</html>`);
  } catch (err) {
    console.error("Login verify page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// Step 2: Verify code matches what's in memory
app.post("/login/verify", async (req, res) => {
  const email = (req.body.email || "").trim();
  const code = (req.body.code || "").trim();
  if (!email || !code) return res.redirect("/login");

  try {
    // Check code against in-memory store (not the sheet)
    if (!checkVerifyCode(email, code)) {
      // Wrong or expired code — show error with retry form
      const sheets = await getSheets();
      const usersRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Users",
      });
      const users = parseTable(usersRes.data.values || []);
      const user = users.find((u) => (u["Email"] || "").toLowerCase() === email.toLowerCase());
      const userName = user ? user["Name"] || email : email;

      return res.send(`<!DOCTYPE html>
<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Verify - Sovereign HUD</title>
<style>
body { background: #0a0b10; color: #00f2ff; font-family: 'Courier New', monospace; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; text-transform: uppercase; }
.login-box { border: 2px solid #ffea00; padding: 40px; box-shadow: 0 0 30px rgba(255,234,0,0.3); text-align: center; max-width: 450px; width: 90%; }
h1 { color: #ffea00; text-shadow: 2px 2px #ff00ff; letter-spacing: 4px; margin: 0 0 10px; }
.error-msg { color: #ff4444; font-size: 0.75em; margin-bottom: 15px; border: 1px solid rgba(255,68,68,0.3); padding: 8px; }
.form-group { margin-bottom: 15px; }
.form-group label { display: block; font-size: 0.7em; color: #ffea00; letter-spacing: 2px; margin-bottom: 5px; }
.code-input { width: 200px; padding: 15px; background: #1a1d26; border: 2px solid #ffea00; color: #00ff9d; font-family: 'Courier New', monospace; font-size: 1.5em; text-align: center; letter-spacing: 8px; box-sizing: border-box; }
.code-input:focus { outline: none; box-shadow: 0 0 15px rgba(255,234,0,0.4); }
.login-btn { width: 100%; padding: 14px; background: #ffea00; color: #0a0b10; border: none; font-family: 'Courier New', monospace; font-size: 1em; font-weight: bold; letter-spacing: 3px; cursor: pointer; }
.back-link { display: block; margin-top: 15px; font-size: 0.7em; color: #555; text-decoration: none; }
.resend-note { color: #888; font-size: 0.65em; margin-top: 12px; text-transform: none; }
</style></head><body><div class="login-box">
<h1>&#x1F50D; VERIFY</h1>
<div style="color:#00ff9d;font-weight:bold;margin-bottom:10px;">${escHtml(userName)}</div>
<div class="error-msg">&#x2717; INCORRECT OR EXPIRED CODE. TRY AGAIN.</div>
<form method="POST" action="/login/verify">
<input type="hidden" name="email" value="${escHtml(email)}">
<div class="form-group"><label>ENTER YOUR 6-DIGIT CODE</label>
<input type="text" name="code" class="code-input" maxlength="6" pattern="[0-9]{6}" required autofocus placeholder="------"></div>
<button type="submit" class="login-btn">VERIFY &amp; ENTER</button></form>
<div class="resend-note">Code expired? <a href="/login" style="color:#ffea00;">Go back</a> and select your account again to get a new code.</div>
<a class="back-link" href="/login">&larr; BACK TO USER SELECT</a>
</div></body></html>`);
    }

    // Code verified — look up user info for cookies
    const sheets = await getSheets();
    const usersRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Users",
    });
    const users = parseTable(usersRes.data.values || []);
    const user = users.find((u) => (u["Email"] || "").toLowerCase() === email.toLowerCase());
    if (!user) return res.redirect("/login");

    // Set cookies (30 day expiry)
    const role = (user["Role"] || "").toLowerCase();
    const opts = { maxAge: 30 * 24 * 60 * 60 * 1000, httpOnly: true, sameSite: "lax", secure: req.secure };
    res.cookie("userEmail", user["Email"], opts);
    res.cookie("userName", user["Name"], opts);
    res.cookie("role", role, opts);
    res.redirect("/");
  } catch (err) {
    console.error("Verify error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

app.post("/login/add-user", async (req, res) => {
  try {
    const name = (req.body.name || "").trim();
    const email = (req.body.email || "").trim();
    const role = (req.body.role || "").trim().toLowerCase();
    if (!name || !email || !role) return res.redirect("/login");

    const sheets = await getSheets();
    await ensureUsersSheet(sheets);

    // Check for duplicate email
    const existingRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Users",
    });
    const existing = parseTable(existingRes.data.values || []);
    if (existing.some((u) => (u["Email"] || "").toLowerCase() === email.toLowerCase())) {
      return res.redirect("/login"); // already exists
    }

    // Append new user (Email, Name, Role)
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Users!A:C",
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: { values: [[email, name, role.charAt(0).toUpperCase() + role.slice(1)]] },
    });

    // If no users existed (first setup), auto-login
    if (existing.length === 0) {
      const opts = { maxAge: 30 * 24 * 60 * 60 * 1000, httpOnly: true, sameSite: "lax", secure: req.secure };
      res.cookie("userEmail", email, opts);
      res.cookie("userName", name, opts);
      res.cookie("role", role, opts);
      return res.redirect("/");
    }

    res.redirect("/login?added=1");
  } catch (err) {
    console.error("Add user error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

app.get("/logout", (req, res) => {
  res.clearCookie("userEmail");
  res.clearCookie("userName");
  res.clearCookie("role");
  res.redirect("/login");
});

// Middleware: require teacher role for admin routes
function requireTeacher(req, res, next) {
  const role = req.cookies && req.cookies.role;
  if (role === "teacher") return next();
  res.status(403).send(`<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Access Denied</title>
<style>body{background:#0a0b10;color:#ff4444;font-family:'Courier New',monospace;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;text-transform:uppercase;}
.box{border:2px solid #ff4444;padding:40px;text-align:center;max-width:400px;box-shadow:0 0 20px rgba(255,68,68,0.3);}
h1{margin:0 0 10px;}a{color:#00f2ff;}</style></head>
<body><div class="box"><h1>&#x1F6AB; ACCESS DENIED</h1><p>TEACHER ROLE REQUIRED</p><a href="/login">LOG IN</a> &bull; <a href="/">BACK TO HUD</a></div></body></html>`);
}

// Apply teacher middleware to all admin routes
app.use("/admin", requireTeacher);

// Require login for all routes below (login/logout are defined above)
app.use((req, res, next) => {
  if (req.cookies && req.cookies.userEmail) return next();
  res.redirect("/login");
});

// ---------------------------------------------------------------------------
// Main HUD
// ---------------------------------------------------------------------------
app.get("/", async (req, res) => {
  if (!req.cookies.userEmail) return res.redirect("/login");
  try {
    const sheets = await getSheets();
    const data = await fetchSheetData(sheets);

    // Quest data for badge + active indicators + recently enslaved
    let activeQuestCount = 0;
    let activeQuestKeys = new Set();
    let recentEnslavedHtml = "";
    let quests = [];
    try {
      quests = await fetchQuestsData(sheets);
      const activeQuests = quests.filter((q) => q["Status"] === "Active");
      activeQuestCount = activeQuests.length;
      for (const q of activeQuests) {
        activeQuestKeys.add(q["Boss"] + "|" + q["Minion"]);
      }
      // Build recently enslaved list (5 most recent approved quests)
      const recentEnslaved = quests
        .filter((q) => q["Status"] === "Approved" && q["Date Resolved"])
        .sort((a, b) => (b["Date Resolved"] || "").localeCompare(a["Date Resolved"] || ""))
        .slice(0, 5);
      if (recentEnslaved.length > 0) {
        const rows = recentEnslaved.map((q) =>
          `<div class="re-entry">
            <span class="re-date">${escHtml((q["Date Resolved"] || "").slice(0, 10))}</span>
            <span class="re-minion">${escHtml(q["Minion"])}</span>
            <span class="re-arrow">&rarr;</span>
            <span class="re-boss">${escHtml(q["Boss"])}</span>
            <span class="re-sector">${escHtml(q["Sector"])}</span>
          </div>`
        ).join("");
        recentEnslavedHtml = `
          <div class="recent-enslaved">
            <h2>RECENTLY ENSLAVED</h2>
            ${rows}
          </div>`;
      }
    } catch { /* Quests sheet may not exist yet */ }

    // Badge evaluation and sync
    let badgeResult = { allEarned: [], newlyEarned: [] };
    try {
      const bossMap = buildBossMap(data.sectors);
      badgeResult = await syncBadges(sheets, data.sectors, data.commandCenter, quests, bossMap);
    } catch (e) { console.error("Badge sync error:", e.message); }

    // Streak computation
    const streakData = computeStreak(quests, data.sectors);

    let html = processAllData(HTML_TEMPLATE, data, activeQuestKeys);
    html = html.split("[[RECENT_ENSLAVED]]").join(recentEnslavedHtml);
    html = html.split("[[STREAK_SECTION]]").join(buildStreakHtml(streakData));
    html = html.split("[[REJECTED_ALERTS]]").join(buildRejectedAlertsHtml(quests));

    const questBadgeHtml = activeQuestCount > 0
      ? `<a href="/quests" class="quest-badge-link"><span class="quest-badge">${activeQuestCount} ACTIVE</span></a>`
      : `<a href="/quests" class="quest-badge-link"><span class="quest-badge dim">QUEST BOARD</span></a>`;
    html = html.split("[[QUEST_BADGE]]").join(questBadgeHtml);

    const questLinkText = activeQuestCount > 0
      ? `&#x2694; ${activeQuestCount} QUEST${activeQuestCount > 1 ? "S" : ""}`
      : `&#x2694; QUESTS`;
    const fabClass = activeQuestCount > 0 ? " pulse" : " dim";
    html = html.split("[[QUEST_LINK]]").join(questLinkText);
    html = html.split('class="nav-quests"').join('class="nav-quests' + fabClass + '"');

    // Army count (Enslaved minions)
    html = html.split("[[ARMY_LINK]]").join(`&#x2694; HENRY'S ARMY`);

    // Badges nav link
    const badgeCount = badgeResult.allEarned.length;
    html = html.split("[[BADGES_LINK]]").join(
      badgeCount > 0 ? `&#x1F3C6; ${badgeCount} BADGE${badgeCount > 1 ? "S" : ""}` : `&#x1F3C6; BADGES`
    );

    // Admin nav — show for teachers (will be wired to user role system)
    const userRole = req.cookies && req.cookies.role;
    const adminNavHtml = userRole === "teacher"
      ? `<a href="/admin" class="nav-admin">&#x2699; ADMIN</a>`
      : "";
    html = html.split("[[ADMIN_NAV]]").join(adminNavHtml);

    // User identity nav
    const userName = req.cookies && req.cookies.userName;
    const userNavText = userName ? `&#x1F464; ${escHtml(userName)}` : `&#x1F464; LOG IN`;
    html = html.split("[[USER_NAV]]").join(userNavText);

    res.send(html);
  } catch (err) {
    console.error("HUD error:", err);
    res.status(500).send(
      `<pre style="background:#0a0b10;color:#ff0000;padding:20px;font-family:monospace;">` +
      `SYSTEM ERROR\n\n${err.message}\n\n` +
      `Check:\n` +
      `  1. credentials.json exists in the project root\n` +
      `  2. SPREADSHEET_ID is set in .env\n` +
      `  3. The sheet is shared with the service account email</pre>`
    );
  }
});

app.get("/boss/:bossName", async (req, res) => {
  try {
    const bossName = decodeURIComponent(req.params.bossName);
    const sector = req.query.sector || "";
    const sheets = await getSheets();
    const batchRes = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SPREADSHEET_ID,
      ranges: ["Sectors", "Command_Center"],
    });
    const [sectorsRaw, ccRaw] = batchRes.data.valueRanges;
    const allMinions = parseTable(sectorsRaw.data || sectorsRaw.values ? sectorsRaw.values : []);
    const commandCenter = parseTable(ccRaw.values);
    const minions = allMinions.filter(
      (r) => r["Boss"] === bossName && (!sector || r["Sector"] === sector)
    );
    const totals = {
      intel: getStat(commandCenter, "Intel").totalPossible,
      stamina: getStat(commandCenter, "Stamina").totalPossible,
      tempo: getStat(commandCenter, "Tempo").totalPossible,
      rep: getStat(commandCenter, "Reputation").totalPossible,
    };

    // Detect survival boss
    let survivalCol = null;
    if (allMinions.length > 0) {
      survivalCol = Object.keys(allMinions[0]).find((k) => k.toLowerCase().includes("survival"));
    }
    let isSurvivalBoss = false;
    if (survivalCol) {
      isSurvivalBoss = allMinions.some(
        (r) => r["Boss"] === bossName && (!sector || r["Sector"] === sector) && (r[survivalCol] || "").toUpperCase() === "X"
      );
    }

    // Fetch active quests to mark already-queued minions
    let activeQuestKeys = new Set();
    try {
      const quests = await fetchQuestsData(sheets);
      for (const q of quests.filter((q) => q["Status"] === "Active")) {
        activeQuestKeys.add(q["Boss"] + "|" + q["Minion"]);
      }
    } catch {}

    res.send(buildBossPage(bossName, sector, minions, totals, activeQuestKeys, isSurvivalBoss));
  } catch (err) {
    console.error("Boss page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// Sector Detail Page — shows all bosses in a sector with their minion tables
// ---------------------------------------------------------------------------
function buildSectorPage(sectorName, bosses, totals, activeQuestKeys, survivalBossNames) {
  const statusColor = { Enslaved: "#00ff9d", Engaged: "#ff6600", Locked: "#555" };
  const norm = (raw, max) => ((parseFloat(raw) || 0) / max * 100).toFixed(1);
  const survivalSet = survivalBossNames || new Set();
  const shieldSvg = `<svg width="14" height="16" viewBox="0 0 18 20" fill="none" style="vertical-align:middle;margin-right:3px;"><path d="M9 0L0 3.5V9C0 14 3.8 18.5 9 20C14.2 18.5 18 14 18 9V3.5L9 0Z" fill="rgba(255,68,68,0.3)" stroke="#ff4444" stroke-width="1.5"/><text x="9" y="13" text-anchor="middle" fill="#ff4444" font-size="8" font-family="monospace" font-weight="bold">S</text></svg>`;

  let bossBlocks = "";
  for (const { bossName, minions } of bosses) {
    let rows = "";
    for (const m of minions) {
      const sc = statusColor[m["Status"]] || "#555";
      const nInt = norm(m["INTELLIGENCE"], totals.intel);
      const nSta = norm(m["STAMINA"], totals.stamina);
      const nTmp = norm(m["TEMPO"], totals.tempo);
      const nRep = norm(m["REPUTATION"], totals.rep);
      const nTotal = (parseFloat(nInt) + parseFloat(nSta) + parseFloat(nTmp) + parseFloat(nRep)).toFixed(1);
      const qKey = bossName + "|" + m["Minion"];
      const onQuest = activeQuestKeys && activeQuestKeys.has(qKey);
      let questBtn;
      if (onQuest) {
        questBtn = `<span style="color:#ffea00;text-shadow:0 0 6px #ffea00;" title="On quest board">&#x2605;</span>`;
      } else if (m["Status"] === "Engaged") {
        questBtn = `<input type="checkbox" class="quest-chk" data-boss="${escHtml(bossName)}" data-minion="${escHtml(m["Minion"])}" data-sector="${escHtml(sectorName)}" title="Select for quest board">`;
      } else {
        questBtn = `<span style="opacity:0.2;">-</span>`;
      }
      const spText = m["Status"] === "Locked" && m["Locked for what?"] ? m["Locked for what?"] : "";
      rows += `
        <tr>
          <td>${questBtn}</td>
          <td title="${escHtml(m["Task"] || "")}">
            ${escHtml(m["Minion"])}
            ${spText ? `<div style="font-size:0.75em;color:#ff0044;margin-top:2px;text-transform:none;">Requires: ${escHtml(spText)}</div>` : ""}
          </td>
          <td style="color:${sc}; font-weight:bold;">${m["Status"]}</td>
          <td>${nInt}</td>
          <td>${nSta}</td>
          <td>${nTmp}</td>
          <td>${nRep}</td>
          <td>${m["Impact(1-3)"] || ""}</td>
          <td style="color:#ffea00; font-weight:bold;">${nTotal}</td>
        </tr>`;
    }

    const isSurvival = survivalSet.has(bossName);
    bossBlocks += `
      <div class="boss-block${isSurvival ? ' survival-boss' : ''}">
        <h2><a href="/boss/${encodeURIComponent(bossName)}?sector=${encodeURIComponent(sectorName)}" class="boss-link">${isSurvival ? shieldSvg + ' ' : ''}${escHtml(bossName)}</a>${isSurvival ? '<span class="survival-tag">GUARDIAN</span>' : ''}</h2>
        <table>
          <thead>
            <tr><th></th><th>Minion</th><th>Status</th><th>INT</th><th>STA</th><th>TMP</th><th>REP</th><th>IMP</th><th>TOTAL</th></tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
  }

  return `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>${escHtml(sectorName)} - Sector Overview</title>
    <style>
    body {
        background: #0a0b10;
        color: #00f2ff;
        font-family: 'Courier New', monospace;
        margin: 0; padding: 20px;
        text-transform: uppercase;
    }
    .hud-container { max-width: 900px; margin: 0 auto; }
    .back-link { display: inline-block; color: #00f2ff; text-decoration: none; border: 1px solid #00f2ff; padding: 6px 15px; margin-bottom: 15px; font-size: 0.8em; transition: all 0.2s; }
    .back-link:hover { background: #00f2ff; color: #0a0b10; }
    h1 { color: #ff00ff; text-align: center; margin: 20px 0 5px 0; text-shadow: 0 0 15px rgba(255,0,255,0.4); letter-spacing: 4px; }
    .sector-tag { text-align: center; font-size: 0.7em; color: #888; letter-spacing: 3px; margin-bottom: 30px; }
    .boss-block { margin-bottom: 35px; }
    .boss-block h2 { color: #ff6600; font-size: 1em; letter-spacing: 3px; margin-bottom: 8px; text-shadow: 0 0 8px rgba(255,102,0,0.3); }
    .boss-block.survival-boss { border-left: 3px solid rgba(255, 68, 68, 0.4); padding-left: 12px; }
    .boss-link { color: #ff6600; text-decoration: none; }
    .boss-link:hover { text-decoration: underline; color: #ffea00; }
    .survival-tag {
        font-size: 0.6em; color: #ff4444; letter-spacing: 2px; margin-left: 10px;
        border: 1px solid rgba(255, 68, 68, 0.4); padding: 2px 6px; vertical-align: middle;
        text-shadow: 0 0 6px rgba(255, 68, 68, 0.3);
    }
    table { width: 100%; border-collapse: collapse; font-size: 0.8em; margin-bottom: 10px; }
    th { padding: 8px; text-align: left; border-bottom: 2px solid #00f2ff; }
    td { padding: 8px; border-bottom: 1px solid #1a1d26; }
    tr:hover td { background: rgba(0, 242, 255, 0.05); }
    td[title]:not([title=""]) { cursor: help; border-bottom: 1px dotted #555; }
    .quest-chk { width: 16px; height: 16px; accent-color: #ff6600; cursor: pointer; }
    .quest-batch-bar {
        display: flex; align-items: center; justify-content: center; gap: 15px;
        margin-top: 15px; padding: 10px; border: 1px solid #ff6600;
        background: rgba(255, 102, 0, 0.08);
    }
    .quest-batch-count { color: #ffea00; font-weight: bold; font-size: 0.85em; letter-spacing: 2px; }
    .quest-batch-btn {
        background: #ff6600; color: #0a0b10; border: none; padding: 8px 18px;
        font-family: 'Courier New', monospace; font-weight: bold; font-size: 0.85em;
        cursor: pointer; letter-spacing: 1px; transition: all 0.2s;
    }
    .quest-batch-btn:hover { background: #ffea00; box-shadow: 0 0 10px rgba(255, 234, 0, 0.5); }
    @media (max-width: 700px) {
        table { font-size: 0.7em; }
        td, th { padding: 5px; }
    }
    </style>
</head>
<body>
    <div class="hud-container">
        <a class="back-link" href="/">&lt; BACK TO HUD</a>
        <h1>${escHtml(sectorName)}</h1>
        <div class="sector-tag">SECTOR OVERVIEW &mdash; ${bosses.length} BOSS${bosses.length !== 1 ? "ES" : ""}</div>
        ${bossBlocks}
        <div class="quest-batch-bar" style="display:none;">
            <span class="quest-batch-count">0 SELECTED</span>
            <button type="button" class="quest-batch-btn">ADD SELECTED TO QUEST BOARD</button>
        </div>
    </div>
    <script>
    (function() {
        var bar = document.querySelector('.quest-batch-bar');
        var countEl = bar.querySelector('.quest-batch-count');
        var btn = bar.querySelector('.quest-batch-btn');
        var checkboxes = document.querySelectorAll('.quest-chk');
        function updateBar() {
            var checked = document.querySelectorAll('.quest-chk:checked');
            if (checked.length > 0) {
                bar.style.display = 'flex';
                countEl.textContent = checked.length + ' SELECTED';
            } else {
                bar.style.display = 'none';
            }
        }
        checkboxes.forEach(function(chk) { chk.addEventListener('change', updateBar); });
        btn.addEventListener('click', function() {
            var checked = document.querySelectorAll('.quest-chk:checked');
            if (checked.length === 0) return;
            var items = [];
            checked.forEach(function(chk) {
                items.push({ boss: chk.dataset.boss, minion: chk.dataset.minion, sector: chk.dataset.sector });
            });
            var form = document.createElement('form');
            form.method = 'POST';
            form.action = '/quest/start-batch';
            var input = document.createElement('input');
            input.type = 'hidden';
            input.name = 'items';
            input.value = JSON.stringify(items);
            var redir = document.createElement('input');
            redir.type = 'hidden';
            redir.name = 'redirect';
            redir.value = window.location.pathname;
            form.appendChild(input);
            form.appendChild(redir);
            document.body.appendChild(form);
            form.submit();
        });
    })();
    </script>
</body>
</html>`;
}

app.get("/sector/:sectorName", async (req, res) => {
  try {
    const sectorName = decodeURIComponent(req.params.sectorName);
    const sheets = await getSheets();
    const batchRes = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SPREADSHEET_ID,
      ranges: ["Sectors", "Command_Center"],
    });
    const [sectorsRaw, ccRaw] = batchRes.data.valueRanges;
    const allMinions = parseTable(sectorsRaw.values || []);
    const commandCenter = parseTable(ccRaw.values);

    const sectorMinions = allMinions.filter((r) => r["Sector"] === sectorName);
    const totals = {
      intel: getStat(commandCenter, "Intel").totalPossible,
      stamina: getStat(commandCenter, "Stamina").totalPossible,
      tempo: getStat(commandCenter, "Tempo").totalPossible,
      rep: getStat(commandCenter, "Reputation").totalPossible,
    };

    // Detect survival bosses in this sector
    let survivalCol = null;
    if (allMinions.length > 0) {
      survivalCol = Object.keys(allMinions[0]).find((k) => k.toLowerCase().includes("survival"));
    }
    const survivalBossNames = new Set();
    if (survivalCol) {
      for (const r of allMinions) {
        if (r["Sector"] === sectorName && (r[survivalCol] || "").toUpperCase() === "X") {
          survivalBossNames.add(r["Boss"]);
        }
      }
    }

    // Group by boss
    const bossMap = {};
    for (const m of sectorMinions) {
      const boss = m["Boss"] || "Unknown";
      if (!bossMap[boss]) bossMap[boss] = [];
      bossMap[boss].push(m);
    }
    const bosses = Object.keys(bossMap).sort().map((b) => ({ bossName: b, minions: bossMap[b] }));

    // Fetch active quests
    let activeQuestKeys = new Set();
    try {
      const quests = await fetchQuestsData(sheets);
      for (const q of quests.filter((q) => q["Status"] === "Active")) {
        activeQuestKeys.add(q["Boss"] + "|" + q["Minion"]);
      }
    } catch {}

    res.send(buildSectorPage(sectorName, bosses, totals, activeQuestKeys, survivalBossNames));
  } catch (err) {
    console.error("Sector page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// Guardians Page — shows all survival-mode bosses with their minions
// ---------------------------------------------------------------------------
function buildGuardiansPage(bosses, totals, activeQuestKeys, survivalBossKeys) {
  const statusColor = { Enslaved: "#00ff9d", Engaged: "#ff6600", Locked: "#555" };
  const norm = (raw, max) => ((parseFloat(raw) || 0) / max * 100).toFixed(1);

  const shieldSvg = `<svg width="14" height="16" viewBox="0 0 18 20" fill="none" style="vertical-align:middle;margin-right:3px;"><path d="M9 0L0 3.5V9C0 14 3.8 18.5 9 20C14.2 18.5 18 14 18 9V3.5L9 0Z" fill="rgba(255,68,68,0.3)" stroke="#ff4444" stroke-width="1.5"/><text x="9" y="13" text-anchor="middle" fill="#ff4444" font-size="8" font-family="monospace" font-weight="bold">S</text></svg>`;

  let bossBlocks = "";
  for (const { bossName, sector, minions } of bosses) {
    let rows = "";
    for (const m of minions) {
      const sc = statusColor[m["Status"]] || "#555";
      const nInt = norm(m["INTELLIGENCE"], totals.intel);
      const nSta = norm(m["STAMINA"], totals.stamina);
      const nTmp = norm(m["TEMPO"], totals.tempo);
      const nRep = norm(m["REPUTATION"], totals.rep);
      const nTotal = (parseFloat(nInt) + parseFloat(nSta) + parseFloat(nTmp) + parseFloat(nRep)).toFixed(1);
      const qKey = bossName + "|" + m["Minion"];
      const onQuest = activeQuestKeys && activeQuestKeys.has(qKey);
      let questBtn;
      if (onQuest) {
        questBtn = `<span style="color:#ffea00;text-shadow:0 0 6px #ffea00;" title="On quest board">&#x2605;</span>`;
      } else if (m["Status"] === "Engaged") {
        questBtn = `<input type="checkbox" class="quest-chk" data-boss="${escHtml(bossName)}" data-minion="${escHtml(m["Minion"])}" data-sector="${escHtml(sector)}" title="Select for quest board">`;
      } else {
        questBtn = `<span style="opacity:0.2;">-</span>`;
      }
      rows += `
        <tr>
          <td>${questBtn}</td>
          <td title="${escHtml(m["Task"] || "")}">${escHtml(m["Minion"])}</td>
          <td style="color:${sc}; font-weight:bold;">${m["Status"]}</td>
          <td>${nInt}</td>
          <td>${nSta}</td>
          <td>${nTmp}</td>
          <td>${nRep}</td>
          <td>${m["Impact(1-3)"] || ""}</td>
          <td style="color:#ffea00; font-weight:bold;">${nTotal}</td>
        </tr>`;
    }

    const enslaved = minions.filter((m) => m["Status"] === "Enslaved").length;
    const total = minions.length;
    const pct = total > 0 ? Math.round((enslaved / total) * 100) : 0;
    const pctColor = pct >= 100 ? "#00ff9d" : pct > 0 ? "#ff6600" : "#555";

    bossBlocks += `
      <div class="boss-block">
        <h2>
          <a href="/boss/${encodeURIComponent(bossName)}?sector=${encodeURIComponent(sector)}" class="boss-link">
            ${shieldSvg} ${escHtml(bossName)}
          </a>
          <span class="boss-sector-tag">${escHtml(sector)}</span>
          <span class="boss-pct" style="color:${pctColor};">${pct}%</span>
        </h2>
        <table>
          <thead>
            <tr><th></th><th>Minion</th><th>Status</th><th>INT</th><th>STA</th><th>TMP</th><th>REP</th><th>IMP</th><th>TOTAL</th></tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
  }

  return `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Ring of Guardians - Sovereign HUD</title>
    <style>
    body {
        background: #0a0b10;
        color: #00f2ff;
        font-family: 'Courier New', monospace;
        margin: 0; padding: 20px;
        text-transform: uppercase;
    }
    .hud-container { max-width: 900px; margin: 0 auto; }
    .back-link { display: inline-block; color: #00f2ff; text-decoration: none; border: 1px solid #00f2ff; padding: 6px 15px; margin-bottom: 15px; font-size: 0.8em; transition: all 0.2s; }
    .back-link:hover { background: #00f2ff; color: #0a0b10; }
    h1 { color: #ff4444; text-align: center; margin: 20px 0 5px 0; text-shadow: 0 0 15px rgba(255,68,68,0.4); letter-spacing: 4px; }
    .guardian-subtitle { text-align: center; font-size: 0.7em; color: #888; letter-spacing: 3px; margin-bottom: 30px; }
    .boss-block { margin-bottom: 35px; }
    .boss-block h2 {
        color: #ff6600; font-size: 1em; letter-spacing: 3px; margin-bottom: 8px;
        text-shadow: 0 0 8px rgba(255,102,0,0.3);
        display: flex; align-items: center; gap: 10px; flex-wrap: wrap;
    }
    .boss-link { color: #ff6600; text-decoration: none; }
    .boss-link:hover { text-decoration: underline; color: #ffea00; }
    .boss-sector-tag { font-size: 0.7em; color: #ff00ff; letter-spacing: 2px; }
    .boss-pct { font-size: 0.8em; margin-left: auto; letter-spacing: 1px; }
    table { width: 100%; border-collapse: collapse; font-size: 0.8em; margin-bottom: 10px; }
    th { padding: 8px; text-align: left; border-bottom: 2px solid #00f2ff; }
    td { padding: 8px; border-bottom: 1px solid #1a1d26; }
    tr:hover td { background: rgba(0, 242, 255, 0.05); }
    td[title]:not([title=""]) { cursor: help; border-bottom: 1px dotted #555; }
    .quest-chk { width: 16px; height: 16px; accent-color: #ff6600; cursor: pointer; }
    .quest-batch-bar {
        display: flex; align-items: center; justify-content: center; gap: 15px;
        margin-top: 15px; padding: 10px; border: 1px solid #ff6600;
        background: rgba(255, 102, 0, 0.08);
    }
    .quest-batch-count { color: #ffea00; font-weight: bold; font-size: 0.85em; letter-spacing: 2px; }
    .quest-batch-btn {
        background: #ff6600; color: #0a0b10; border: none; padding: 8px 18px;
        font-family: 'Courier New', monospace; font-weight: bold; font-size: 0.85em;
        cursor: pointer; letter-spacing: 1px; transition: all 0.2s;
    }
    .quest-batch-btn:hover { background: #ffea00; box-shadow: 0 0 10px rgba(255, 234, 0, 0.5); }
    @media (max-width: 700px) {
        table { font-size: 0.7em; }
        td, th { padding: 5px; }
    }
    </style>
</head>
<body>
    <div class="hud-container">
        <a class="back-link" href="/">&lt; BACK TO HUD</a>
        <h1>&#x1F6E1; Ring of Guardians</h1>
        <div class="guardian-subtitle">SURVIVAL MODE BOSSES &mdash; ENSLAVE ALL GUARDIANS TO UNLOCK SURVIVAL</div>
        ${bossBlocks}
        <div class="quest-batch-bar" style="display:none;">
            <span class="quest-batch-count">0 SELECTED</span>
            <button type="button" class="quest-batch-btn">ADD SELECTED TO QUEST BOARD</button>
        </div>
        <div style="text-align:center;margin-top:15px;font-size:0.75em;color:#ff6600;">SELECT ENGAGED MINIONS TO ADD TO YOUR QUEST BOARD</div>
    </div>
    <script>
    (function() {
        var bar = document.querySelector('.quest-batch-bar');
        var countEl = bar.querySelector('.quest-batch-count');
        var btn = bar.querySelector('.quest-batch-btn');
        var checkboxes = document.querySelectorAll('.quest-chk');
        function updateBar() {
            var checked = document.querySelectorAll('.quest-chk:checked');
            if (checked.length > 0) {
                bar.style.display = 'flex';
                countEl.textContent = checked.length + ' SELECTED';
            } else {
                bar.style.display = 'none';
            }
        }
        checkboxes.forEach(function(chk) { chk.addEventListener('change', updateBar); });
        btn.addEventListener('click', function() {
            var checked = document.querySelectorAll('.quest-chk:checked');
            if (checked.length === 0) return;
            var items = [];
            checked.forEach(function(chk) {
                items.push({ boss: chk.dataset.boss, minion: chk.dataset.minion, sector: chk.dataset.sector });
            });
            var form = document.createElement('form');
            form.method = 'POST';
            form.action = '/quest/start-batch';
            var input = document.createElement('input');
            input.type = 'hidden';
            input.name = 'items';
            input.value = JSON.stringify(items);
            var redir = document.createElement('input');
            redir.type = 'hidden';
            redir.name = 'redirect';
            redir.value = '/guardians';
            form.appendChild(input);
            form.appendChild(redir);
            document.body.appendChild(form);
            form.submit();
        });
    })();
    </script>
</body>
</html>`;
}

app.get("/guardians", async (req, res) => {
  try {
    const sheets = await getSheets();
    const batchRes = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SPREADSHEET_ID,
      ranges: ["Sectors", "Command_Center"],
    });
    const [sectorsRaw, ccRaw] = batchRes.data.valueRanges;
    const allMinions = parseTable(sectorsRaw.values || []);
    const commandCenter = parseTable(ccRaw.values);

    // Find survival column dynamically
    let survivalCol = null;
    if (allMinions.length > 0) {
      survivalCol = Object.keys(allMinions[0]).find((k) => k.toLowerCase().includes("survival"));
    }

    // Collect survival boss keys
    const survivalBossKeys = new Set();
    for (const row of allMinions) {
      if (survivalCol && (row[survivalCol] || "").toUpperCase() === "X") {
        survivalBossKeys.add(`${row["Sector"]}|${row["Boss"]}`);
      }
    }

    // Filter minions that belong to survival bosses
    const survivalMinions = allMinions.filter((r) =>
      survivalBossKeys.has(`${r["Sector"]}|${r["Boss"]}`)
    );

    const totals = {
      intel: getStat(commandCenter, "Intel").totalPossible,
      stamina: getStat(commandCenter, "Stamina").totalPossible,
      tempo: getStat(commandCenter, "Tempo").totalPossible,
      rep: getStat(commandCenter, "Reputation").totalPossible,
    };

    // Group by boss
    const bossOrder = [];
    const bossGroups = {};
    for (const m of survivalMinions) {
      const key = `${m["Sector"]}|${m["Boss"]}`;
      if (!bossGroups[key]) {
        bossGroups[key] = { bossName: m["Boss"], sector: m["Sector"], minions: [] };
        bossOrder.push(key);
      }
      bossGroups[key].minions.push(m);
    }

    // Sort by completion: highest first
    bossOrder.sort((a, b) => {
      const aG = bossGroups[a], bG = bossGroups[b];
      const aPct = aG.minions.filter((m) => m["Status"] === "Enslaved").length / (aG.minions.length || 1);
      const bPct = bG.minions.filter((m) => m["Status"] === "Enslaved").length / (bG.minions.length || 1);
      if (bPct !== aPct) return bPct - aPct;
      return aG.bossName.localeCompare(bG.bossName);
    });

    const bosses = bossOrder.map((k) => bossGroups[k]);

    // Fetch active quests
    let activeQuestKeys = new Set();
    try {
      const quests = await fetchQuestsData(sheets);
      for (const q of quests.filter((q) => q["Status"] === "Active")) {
        activeQuestKeys.add(q["Boss"] + "|" + q["Minion"]);
      }
    } catch {}

    res.send(buildGuardiansPage(bosses, totals, activeQuestKeys, survivalBossKeys));
  } catch (err) {
    console.error("Guardians page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// Quest Board Page Template
// ---------------------------------------------------------------------------
function buildQuestBoardPage(questRows, activeCount, totalCount) {
  return `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Quest Board - Sovereign HUD</title>
    <style>
    body {
        background: #0a0b10;
        color: #00f2ff;
        font-family: 'Courier New', monospace;
        padding: 20px;
        text-transform: uppercase;
        overflow-x: hidden;
    }
    .hud-container {
        border: 2px solid #00f2ff;
        padding: 20px;
        box-shadow: 0 0 15px #00f2ff;
        max-width: 960px;
        margin: auto;
    }
    h1 {
        text-shadow: 2px 2px #ff00ff;
        border-bottom: 1px solid #00f2ff;
        letter-spacing: 2px;
        text-align: center;
        margin-bottom: 15px;
    }
    .back-link {
        display: inline-block;
        color: #00f2ff;
        text-decoration: none;
        border: 1px solid #00f2ff;
        padding: 6px 15px;
        margin-bottom: 15px;
        font-size: 0.8em;
        transition: all 0.2s;
    }
    .back-link:hover { background: #00f2ff; color: #0a0b10; }
    .quest-stats {
        text-align: center;
        margin-bottom: 20px;
        font-size: 0.85em;
        color: #ff6600;
    }
    .quest-card {
        border: 1px solid rgba(0, 242, 255, 0.3);
        background: rgba(0, 242, 255, 0.03);
        border-radius: 6px;
        padding: 15px;
        margin-bottom: 15px;
        transition: border-color 0.2s;
    }
    .quest-card:hover { border-color: #00f2ff; }
    .quest-card[data-status="Approved"] { border-color: rgba(0, 255, 157, 0.3); opacity: 0.7; }
    .quest-card[data-status="Rejected"] { border-color: rgba(255, 0, 68, 0.5); }
    .quest-card[data-status="Abandoned"] { border-color: rgba(85, 85, 85, 0.3); opacity: 0.4; }
    .quest-header {
        display: flex;
        justify-content: space-between;
        margin-bottom: 10px;
        font-size: 0.75em;
    }
    .quest-id { color: #555; }
    .quest-status { font-weight: bold; letter-spacing: 2px; }
    .quest-body { margin-bottom: 12px; }
    .quest-target { font-size: 1.1em; margin-bottom: 5px; }
    .quest-boss { color: #888; }
    .quest-arrow { color: #ff00ff; margin: 0 8px; }
    .quest-minion { color: #00f2ff; font-weight: bold; }
    .quest-sector { font-size: 0.75em; color: #ff00ff; margin-bottom: 8px; }
    .quest-suggestion {
        font-size: 0.8em;
        color: #ccc;
        border-left: 2px solid #ff00ff;
        padding-left: 10px;
        margin-top: 8px;
        text-transform: none;
    }
    .quest-task-label {
        color: #ff00ff;
        font-weight: bold;
        letter-spacing: 1px;
    }
    .quest-proof-type { color: #ffea00; margin-right: 5px; }
    .quest-footer { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
    .quest-submit-form { display: flex; gap: 8px; flex: 1; align-items: center; }
    .proof-input {
        flex: 1;
        background: #1a1d26;
        border: 1px solid #333;
        color: #00f2ff;
        padding: 6px 10px;
        font-family: 'Courier New', monospace;
        font-size: 0.75em;
        text-transform: none;
    }
    .proof-input:focus { outline: none; border-color: #00f2ff; box-shadow: 0 0 5px rgba(0, 242, 255, 0.3); }
    .artifact-select {
        background: #1a1d26;
        border: 1px solid #ffea00;
        color: #ffea00;
        padding: 6px 8px;
        font-family: 'Courier New', monospace;
        font-size: 0.75em;
        text-transform: uppercase;
        cursor: pointer;
        min-width: 120px;
    }
    .artifact-select:focus { outline: none; border-color: #ffea00; box-shadow: 0 0 5px rgba(255, 234, 0, 0.3); }
    .quest-submit-btn {
        background: none;
        border: 1px solid #00ff9d;
        color: #00ff9d;
        padding: 6px 15px;
        font-family: 'Courier New', monospace;
        font-size: 0.75em;
        cursor: pointer;
        text-transform: uppercase;
        transition: all 0.2s;
    }
    .quest-submit-btn:hover:not(:disabled) { background: #00ff9d; color: #0a0b10; }
    .quest-submit-btn:disabled {
        border-color: #444;
        color: #555;
        cursor: not-allowed;
        opacity: 0.5;
    }
    .quest-abandon-btn {
        background: none;
        border: 1px solid #ff0044;
        color: #ff0044;
        width: 28px;
        height: 28px;
        font-size: 0.8em;
        font-weight: bold;
        cursor: pointer;
        font-family: 'Courier New', monospace;
        transition: all 0.2s;
        padding: 0;
        flex-shrink: 0;
    }
    .quest-abandon-btn:hover { background: #ff0044; color: #0a0b10; box-shadow: 0 0 8px rgba(255, 0, 68, 0.5); }
    .quest-date { font-size: 0.7em; color: #555; }
    .quest-status-col { text-align: right; }
    .quest-reject-reason {
        font-size: 0.75em; color: #ff0044; margin-top: 2px;
        text-transform: none; letter-spacing: 0.5px;
    }
    .quest-feedback { font-size: 0.8em; color: #ffea00; border-left: 2px solid #ffea00; padding: 6px 10px; margin-top: 6px; text-transform: none; background: rgba(255,234,0,0.05); }
    .quest-feedback-label { color: #ff8800; font-weight: bold; text-transform: uppercase; }
    .quest-proof-link { font-size: 0.75em; color: #555; text-transform: none; }
    .no-quests { text-align: center; color: #555; padding: 40px; font-size: 0.9em; }
    @media (max-width: 600px) {
        body { padding: 10px; }
        .hud-container { padding: 15px; }
        .quest-submit-form { flex-direction: column; }
    }
    </style>
</head>
<body>
    <div class="hud-container">
        <a class="back-link" href="/">&lt; BACK TO HUD</a>
        <h1>Quest Board</h1>
        <div class="quest-stats">${activeCount} ACTIVE / ${totalCount} TOTAL QUESTS</div>
        ${questRows || '<div class="no-quests">NO QUESTS YET. START ONE FROM POWER RANKINGS.</div>'}
    </div>
    <script>
    function confirmAbandon(form) {
        return confirm('Are you sure you want to abandon this quest?');
    }
    (function() {
        document.querySelectorAll('.proof-input').forEach(function(input) {
            const btn = input.closest('form').querySelector('.quest-submit-btn');
            if (!btn) return;
            input.addEventListener('input', function() {
                btn.disabled = input.value.trim().length === 0;
            });
        });
    })();
    </script>
</body>
</html>`;
}

// ---------------------------------------------------------------------------
// Quest routes
// ---------------------------------------------------------------------------
app.post("/quest/start", async (req, res) => {
  try {
    const { boss, minion, sector } = req.body;
    if (!boss || !minion || !sector) {
      return res.status(400).send("Missing required fields: boss, minion, sector");
    }

    const sheets = await getSheets();
    await ensureQuestsSheet(sheets);

    const [defRes, sectorsRes] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: "Definitions" }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: "Sectors" }),
    ]);
    const definitions = parseTable(defRes.data.values);
    const sectorsRows = parseTable(sectorsRes.data.values);

    const { proofType, suggestion: fallbackSuggestion } = generateProofSuggestion(sector, definitions);

    // Look up the Task from the Sectors sheet for this specific boss+minion
    const matchingRow = sectorsRows.find(
      (r) => r["Boss"] === boss && r["Minion"] === minion && r["Sector"] === sector
    );
    const taskDetail = (matchingRow && matchingRow["Task"]) ? matchingRow["Task"] : fallbackSuggestion;

    const questId = generateQuestId();
    const today = new Date().toISOString().slice(0, 10);

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Quests!A:L",
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[questId, boss, minion, sector, "Active", proofType, "", taskDetail, "", today, "", ""]],
      },
    });

    await updateSectorsQuestStatus(sheets, sector, boss, minion, "Active");

    res.redirect("/quests");
  } catch (err) {
    console.error("Quest start error:", err);
    res.status(500).send(`<pre style="color:red">Error starting quest: ${err.message}</pre>`);
  }
});

// Batch add multiple minions to quest board at once
app.post("/quest/start-batch", async (req, res) => {
  try {
    let { items, redirect } = req.body;
    // items is a JSON string: [{ boss, minion, sector }, ...]
    if (typeof items === "string") items = JSON.parse(items);
    if (!Array.isArray(items) || items.length === 0) {
      return res.status(400).send("No items selected");
    }

    const sheets = await getSheets();
    await ensureQuestsSheet(sheets);

    const [defRes, sectorsRes] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: "Definitions" }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: "Sectors" }),
    ]);
    const definitions = parseTable(defRes.data.values);
    const sectorsRows = parseTable(sectorsRes.data.values);

    const today = new Date().toISOString().slice(0, 10);
    const newRows = [];
    for (const item of items) {
      const { boss, minion, sector } = item;
      if (!boss || !minion || !sector) continue;

      const { proofType, suggestion: fallbackSuggestion } = generateProofSuggestion(sector, definitions);
      const matchingRow = sectorsRows.find(
        (r) => r["Boss"] === boss && r["Minion"] === minion && r["Sector"] === sector
      );
      const taskDetail = (matchingRow && matchingRow["Task"]) ? matchingRow["Task"] : fallbackSuggestion;
      const questId = generateQuestId();
      newRows.push([questId, boss, minion, sector, "Active", proofType, "", taskDetail, "", today, "", ""]);
    }

    if (newRows.length > 0) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Quests!A:L",
        valueInputOption: "RAW",
        insertDataOption: "INSERT_ROWS",
        requestBody: { values: newRows },
      });

      // Sync Quest Status to Sectors for each item
      for (const item of items) {
        if (item.boss && item.minion && item.sector) {
          await updateSectorsQuestStatus(sheets, item.sector, item.boss, item.minion, "Active");
        }
      }
    }

    res.redirect(redirect || "/quests");
  } catch (err) {
    console.error("Quest start-batch error:", err);
    res.status(500).send(`<pre style="color:red">Error starting quests: ${err.message}</pre>`);
  }
});

app.post("/quest/submit", async (req, res) => {
  try {
    const { questId, proofLink, artifactType } = req.body;
    if (!questId) return res.status(400).send("Missing questId");

    const sheets = await getSheets();
    const questsRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Quests",
    });
    const rows = questsRes.data.values;
    if (!rows || rows.length < 2) return res.status(404).send("No quests found");

    const headerRow = rows[0];
    const idCol = headerRow.indexOf("Quest ID");
    const statusCol = headerRow.indexOf("Status");
    const proofTypeCol = headerRow.indexOf("Proof Type");
    const proofLinkCol = headerRow.indexOf("Proof Link");
    const dateCol = headerRow.indexOf("Date Completed");
    const bossCol = headerRow.indexOf("Boss");
    const minionCol = headerRow.indexOf("Minion");
    const sectorCol = headerRow.indexOf("Sector");

    let targetRowIdx = -1;
    let questRow = null;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][idCol] === questId) {
        targetRowIdx = i + 1; // 1-based for Sheets API
        questRow = rows[i];
        break;
      }
    }

    if (targetRowIdx === -1) return res.status(404).send("Quest not found");

    const updates = [];
    updates.push({
      range: "Quests!" + String.fromCharCode(65 + statusCol) + targetRowIdx,
      values: [["Submitted"]],
    });
    if (artifactType && proofTypeCol >= 0) {
      updates.push({
        range: "Quests!" + String.fromCharCode(65 + proofTypeCol) + targetRowIdx,
        values: [[artifactType]],
      });
    }
    if (proofLink) {
      updates.push({
        range: "Quests!" + String.fromCharCode(65 + proofLinkCol) + targetRowIdx,
        values: [[proofLink]],
      });
    }
    updates.push({
      range: "Quests!" + String.fromCharCode(65 + dateCol) + targetRowIdx,
      values: [[new Date().toISOString().split("T")[0]]],
    });
    // Clear old feedback on resubmission
    const feedbackCol = headerRow.indexOf("Feedback");
    if (feedbackCol >= 0) {
      updates.push({
        range: "Quests!" + String.fromCharCode(65 + feedbackCol) + targetRowIdx,
        values: [[""]],
      });
    }
    // Clear Date Resolved (quest is active again)
    const dateResolvedCol = headerRow.indexOf("Date Resolved");
    if (dateResolvedCol >= 0) {
      updates.push({
        range: "Quests!" + String.fromCharCode(65 + dateResolvedCol) + targetRowIdx,
        values: [[""]],
      });
    }

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "RAW", data: updates },
    });

    // Sync Quest Status to Sectors
    const qBoss = questRow[bossCol];
    const qMinion = questRow[minionCol];
    const qSector = questRow[sectorCol];
    if (qBoss && qMinion && qSector) {
      await updateSectorsQuestStatus(sheets, qSector, qBoss, qMinion, "Submitted");
    }

    res.redirect("/quests");
  } catch (err) {
    console.error("Quest submit error:", err);
    res.status(500).send(`<pre style="color:red">Error submitting quest: ${err.message}</pre>`);
  }
});

app.post("/quest/remove", requireTeacher, async (req, res) => {
  try {
    const { questId } = req.body;
    if (!questId) return res.status(400).send("Missing questId");
    const loggedInUser = (req.cookies && req.cookies.userName) || (req.cookies && req.cookies.userEmail) || "Unknown";

    const sheets = await getSheets();
    const questsRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Quests",
    });
    const rows = questsRes.data.values;
    if (!rows || rows.length < 2) return res.status(404).send("No quests found");

    const headers = rows[0];
    const idCol = headers.indexOf("Quest ID");
    const statusCol = headers.indexOf("Status");
    const bossCol = headers.indexOf("Boss");
    const minionCol = headers.indexOf("Minion");
    const sectorCol = headers.indexOf("Sector");
    const dateCol = headers.indexOf("Date Completed");

    let targetRowIdx = -1;
    let questRow = null;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][idCol] === questId) {
        targetRowIdx = i; // 0-based data index
        questRow = rows[i];
        break;
      }
    }
    if (targetRowIdx === -1) return res.status(404).send("Quest not found");

    // Clear Quest Status on Sectors
    const qBoss = questRow[bossCol];
    const qMinion = questRow[minionCol];
    const qSector = questRow[sectorCol];
    if (qBoss && qMinion && qSector) {
      await updateSectorsQuestStatus(sheets, qSector, qBoss, qMinion, "");
    }

    // Instead of deleting, mark as Abandoned with logged-in user and date for audit trail
    const now = new Date().toISOString().slice(0, 10);
    const parentNote = `Abandoned by: ${loggedInUser}`;
    // Update Status to "Abandoned" and Date Completed to the audit note
    const updatedRow = [...questRow];
    while (updatedRow.length < headers.length) updatedRow.push("");
    updatedRow[statusCol] = "Abandoned";
    updatedRow[dateCol] = `${now} | ${parentNote}`;
    const dateResolvedCol = headers.indexOf("Date Resolved");
    if (dateResolvedCol >= 0) updatedRow[dateResolvedCol] = now;

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Quests!A${targetRowIdx + 1}:${String.fromCharCode(64 + headers.length)}${targetRowIdx + 1}`,
      valueInputOption: "RAW",
      requestBody: { values: [updatedRow] },
    });

    res.redirect("/quests");
  } catch (err) {
    console.error("Quest remove error:", err);
    res.status(500).send(`<pre style="color:red">Error removing quest: ${err.message}</pre>`);
  }
});

app.get("/quests", async (req, res) => {
  try {
    const sheets = await getSheets();
    const [quests, defRes] = await Promise.all([
      fetchQuestsData(sheets),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: "Definitions" }),
    ]);
    const definitions = parseTable(defRes.data.values);
    const artifactOptions = getArtifactOptions(definitions);

    const statusOrder = ["Active", "Submitted", "Approved", "Rejected"];
    const statusColors = { Active: "#ff6600", Submitted: "#ffea00", Approved: "#00ff9d", Rejected: "#ff0044" };

    // Only show actionable quests — Active, Submitted, Rejected (not Approved/Abandoned)
    const visible = quests.filter((q) => q["Status"] !== "Abandoned" && q["Status"] !== "Approved");

    const sorted = [...visible].sort((a, b) =>
      statusOrder.indexOf(a["Status"]) - statusOrder.indexOf(b["Status"])
    );

    const isTeacher = req.cookies && req.cookies.role === "teacher";
    let questRows = "";
    for (const q of sorted) {
      const sc = statusColors[q["Status"]] || "#555";
      const isActive = q["Status"] === "Active";
      const isRejected = q["Status"] === "Rejected";
      const canAbandon = isTeacher && (isActive || isRejected);
      const abandonBtn = canAbandon
        ? `<form method="POST" action="/quest/remove" style="margin:0;" onsubmit="return confirmAbandon(this)">
             <input type="hidden" name="questId" value="${q["Quest ID"]}">
             <button type="submit" class="quest-abandon-btn" title="Abandon quest">X</button>
           </form>`
        : "";

      // Build artifact type dropdown for submission
      const currentType = q["Proof Type"] || "";
      let submitForm = "";
      if (isActive || isRejected) {
        const opts = artifactOptions.map((opt) =>
          `<option value="${escHtml(opt)}"${opt === currentType ? " selected" : ""}>${escHtml(opt)}</option>`
        ).join("");
        submitForm = `<form class="quest-submit-form" method="POST" action="/quest/submit">
             <input type="hidden" name="questId" value="${q["Quest ID"]}">
             <select name="artifactType" class="artifact-select"><option value="" disabled${!currentType ? " selected" : ""}>ARTIFACT...</option>${opts}</select>
             <input type="text" name="proofLink" placeholder="PASTE LINK OR INCLUDE DETAILS..." class="proof-input" required>
             <button type="submit" class="quest-submit-btn" disabled>SUBMIT</button>
           </form>`;
      } else {
        const proofDisplay = currentType
          ? `<span class="quest-proof-type">[${escHtml(currentType)}]</span> ${q["Proof Link"] || "---"}`
          : `${q["Proof Link"] || "---"}`;
        submitForm = `<span class="quest-proof-link">${proofDisplay}</span>`;
      }

      questRows += `
        <div class="quest-card" data-status="${q["Status"]}">
          <div class="quest-header">
            <span class="quest-id">${q["Quest ID"]}</span>
            <div class="quest-status-col">
              <span class="quest-status" style="color:${sc}">${q["Status"]}</span>
              ${isRejected && q["Feedback"] ? `<div class="quest-reject-reason">REASON: ${escHtml(q["Feedback"])}</div>` : ""}
            </div>
          </div>
          <div class="quest-body">
            <div class="quest-target">
              <span class="quest-boss">${escHtml(q["Boss"])}</span>
              <span class="quest-arrow">&gt;</span>
              <span class="quest-minion">${escHtml(q["Minion"])}</span>
            </div>
            <div class="quest-sector">SECTOR: ${escHtml(q["Sector"])}</div>
            <div class="quest-suggestion">
              <span class="quest-task-label">TASK:</span> ${q["Suggested By AI"] || "No task details available."}
            </div>
          </div>
          <div class="quest-footer">
            ${submitForm}
            ${abandonBtn}
            ${q["Date Completed"] ? '<span class="quest-date">COMPLETED: ' + q["Date Completed"] + "</span>" : ""}
          </div>
        </div>`;
    }

    const activeCount = visible.filter((q) => q["Status"] === "Active").length;
    res.send(buildQuestBoardPage(questRows, activeCount, visible.length));
  } catch (err) {
    console.error("Quest board error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// Badges Page
// ---------------------------------------------------------------------------
app.get("/badges", async (req, res) => {
  try {
    const sheets = await getSheets();
    await ensureBadgesSheet(sheets);

    const batchRes = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SPREADSHEET_ID,
      ranges: ["Sectors", "Command_Center", "Quests", "Badges"],
    });
    const [sectorsRaw, ccRaw, questsRaw, badgesRaw] = batchRes.data.valueRanges;
    const sectors = parseTable(sectorsRaw.values || []);
    const commandCenter = parseTable(ccRaw.values || []);
    const quests = parseTable(questsRaw.values || []);
    const earnedBadges = parseTable(badgesRaw.values || []);
    const earnedSet = new Set(earnedBadges.map((b) => b["Badge ID"]));
    const earnedLookup = {};
    for (const b of earnedBadges) earnedLookup[b["Badge ID"]] = b["Date Earned"];

    const bossMap = buildBossMap(sectors);

    // Build full badge catalog: static definitions + dynamic boss badges
    const allBadges = { ...BADGE_DEFINITIONS };
    for (const sector in bossMap) {
      for (const bossName in bossMap[sector]) {
        const id = `boss:${sector}:${bossName}`;
        if (!allBadges[id]) allBadges[id] = getBossBadgeDef(sector, bossName);
      }
    }

    const categories = [
      { key: "meta",   label: "SPECIAL ACHIEVEMENTS", color: "#ffea00" },
      { key: "boss",   label: "BOSS DEFEATS",         color: "#ff0044" },
      { key: "sector", label: "SECTOR DOMINION",       color: "#ff00ff" },
      { key: "stat",   label: "STAT MILESTONES",       color: "#00f2ff" },
    ];

    let sectionsHtml = "";
    for (const cat of categories) {
      const badges = Object.entries(allBadges)
        .filter(([, def]) => def.category === cat.key)
        .sort((a, b) => {
          const aE = earnedSet.has(a[0]) ? 0 : 1;
          const bE = earnedSet.has(b[0]) ? 0 : 1;
          if (aE !== bE) return aE - bE;
          return a[1].name.localeCompare(b[1].name);
        });

      const earnedInCat = badges.filter(([id]) => earnedSet.has(id)).length;

      let badgeCards = "";
      for (const [id, def] of badges) {
        const isEarned = earnedSet.has(id);
        const date = isEarned ? (earnedLookup[id] || "").slice(0, 10) : "";
        badgeCards += `
          <div class="badge-card ${isEarned ? "earned" : "locked"}">
            <span class="bc-icon" style="color:${isEarned ? def.color : "#333"}">${def.icon}</span>
            <div class="bc-name">${escHtml(def.name)}</div>
            <div class="bc-desc">${escHtml(def.description)}</div>
            ${isEarned ? `<div class="bc-date">EARNED: ${date}</div>` : '<div class="bc-locked">LOCKED</div>'}
          </div>`;
      }
      sectionsHtml += `
        <div class="badge-category">
          <h2 style="color:${cat.color}">${cat.label} <span class="bc-count">(${earnedInCat}/${badges.length})</span></h2>
          <div class="badge-grid">${badgeCards}</div>
        </div>`;
    }

    const totalEarned = earnedBadges.length;
    const totalPossible = Object.keys(allBadges).length;

    res.send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Badges - Sovereign HUD</title>
    <style>
    body {
        background: #0a0b10;
        color: #00f2ff;
        font-family: 'Courier New', monospace;
        padding: 20px;
        text-transform: uppercase;
        overflow-x: hidden;
    }
    .hud-container {
        border: 2px solid #ffea00;
        padding: 20px;
        box-shadow: 0 0 15px rgba(255, 234, 0, 0.5);
        max-width: 960px;
        margin: auto;
    }
    h1 {
        text-shadow: 2px 2px #ff00ff;
        border-bottom: 1px solid #ffea00;
        letter-spacing: 2px;
        text-align: center;
        margin-bottom: 5px;
        color: #ffea00;
    }
    .back-link {
        display: inline-block;
        color: #ffea00;
        text-decoration: none;
        font-size: 0.85em;
        margin-bottom: 10px;
        letter-spacing: 1px;
    }
    .back-link:hover { text-decoration: underline; }
    .badge-summary {
        text-align: center;
        font-size: 1.1em;
        color: #ffea00;
        margin: 10px 0 20px 0;
        text-shadow: 0 0 8px rgba(255, 234, 0, 0.3);
    }
    .badge-summary .earned-count { font-size: 1.6em; color: #00ff9d; }
    .badge-category { margin-bottom: 30px; }
    .badge-category h2 {
        font-size: 1em;
        letter-spacing: 3px;
        margin-bottom: 15px;
        border-bottom: 1px solid rgba(255, 255, 255, 0.1);
        padding-bottom: 8px;
    }
    .bc-count { font-weight: normal; font-size: 0.7em; }
    .badge-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(140px, 1fr));
        gap: 15px;
    }
    .badge-card {
        border: 1px solid rgba(0, 242, 255, 0.3);
        background: rgba(0, 242, 255, 0.03);
        border-radius: 8px;
        padding: 15px;
        text-align: center;
        transition: all 0.3s;
    }
    .badge-card.earned {
        border-color: rgba(255, 234, 0, 0.5);
        box-shadow: 0 0 10px rgba(255, 234, 0, 0.1);
    }
    .badge-card.locked {
        opacity: 0.35;
        filter: grayscale(1);
    }
    .bc-icon { font-size: 2.5em; display: block; margin-bottom: 8px; }
    .bc-name { font-size: 0.8em; color: #ffea00; font-weight: bold; letter-spacing: 1px; margin-bottom: 4px; }
    .bc-desc { font-size: 0.65em; color: #888; text-transform: none; line-height: 1.3; }
    .bc-date { font-size: 0.6em; color: #00ff9d; margin-top: 6px; }
    .bc-locked { font-size: 0.6em; color: #555; margin-top: 6px; letter-spacing: 2px; }
    .nav-buttons {
        display: flex;
        gap: 10px;
        justify-content: center;
        flex-wrap: wrap;
        margin-top: 20px;
        padding-top: 15px;
        border-top: 1px dashed #333;
    }
    .nav-buttons a {
        color: #00f2ff;
        text-decoration: none;
        border: 1px solid #00f2ff;
        padding: 8px 16px;
        border-radius: 4px;
        font-size: 0.75em;
        letter-spacing: 1px;
        transition: all 0.3s;
    }
    .nav-buttons a:hover {
        background: #00f2ff;
        color: #0a0b10;
    }
    @media (max-width: 600px) {
        .badge-grid { grid-template-columns: repeat(auto-fill, minmax(120px, 1fr)); gap: 10px; }
        .badge-card { padding: 10px; }
        .bc-icon { font-size: 2em; }
    }
    </style>
</head>
<body>
    <div class="hud-container">
        <a href="/" class="back-link">&lt; BACK TO HUD</a>
        <h1>&#x1F3C6; ACHIEVEMENT BADGES</h1>
        <div class="badge-summary">
            <span class="earned-count">${totalEarned}</span> / ${totalPossible} BADGES EARNED
        </div>
        ${sectionsHtml}
        <div class="nav-buttons">
            <a href="/">&lt; HUD</a>
            <a href="/army">ARMY</a>
            <a href="/quests">QUESTS</a>
            <a href="/progress">PROGRESS</a>
        </div>
    </div>
</body>
</html>`);
  } catch (err) {
    console.error("Badges page error:", err);
    res.status(500).send(`<pre style="background:#0a0b10;color:#ff0000;padding:20px;font-family:monospace;">BADGES ERROR\n\n${err.message}</pre>`);
  }
});

app.get("/army", async (req, res) => {
  try {
    const sheets = await getSheets();
    const batchRes = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SPREADSHEET_ID,
      ranges: ["Sectors", "Command_Center", "Quests"],
    });
    const [sectorsRaw, ccRaw, questsRaw] = batchRes.data.valueRanges;
    const allMinions = parseTable(sectorsRaw.values);
    const commandCenter = parseTable(ccRaw.values);
    const quests = parseTable(questsRaw.values);
    const enslaved = allMinions.filter((r) => r["Status"] === "Enslaved");

    // Build enslaved date lookup from Quests: Sector|Boss|Minion -> Date Resolved
    const enslavedDateLookup = {};
    for (const q of quests) {
      if (q["Status"] === "Approved" && q["Date Resolved"]) {
        const key = `${q["Sector"]}|${q["Boss"]}|${q["Minion"]}`;
        enslavedDateLookup[key] = q["Date Resolved"].slice(0, 10);
      }
    }

    // Detect survival bosses
    let survivalCol = null;
    if (allMinions.length > 0) {
      survivalCol = Object.keys(allMinions[0]).find((k) => k.toLowerCase().includes("survival"));
    }
    const survivalBossKeys = new Set();
    if (survivalCol) {
      for (const r of allMinions) {
        if ((r[survivalCol] || "").toUpperCase() === "X") {
          survivalBossKeys.add(`${r["Sector"]}|${r["Boss"]}`);
        }
      }
    }

    const shieldSvg = `<svg width="12" height="14" viewBox="0 0 18 20" fill="none" style="vertical-align:middle;margin-right:2px;"><path d="M9 0L0 3.5V9C0 14 3.8 18.5 9 20C14.2 18.5 18 14 18 9V3.5L9 0Z" fill="rgba(255,68,68,0.3)" stroke="#ff4444" stroke-width="1.5"/><text x="9" y="13" text-anchor="middle" fill="#ff4444" font-size="8" font-family="monospace" font-weight="bold">S</text></svg>`;

    const totals = {
      intel: getStat(commandCenter, "Intel").totalPossible,
      stamina: getStat(commandCenter, "Stamina").totalPossible,
      tempo: getStat(commandCenter, "Tempo").totalPossible,
      rep: getStat(commandCenter, "Reputation").totalPossible,
    };
    const norm = (raw, max) => ((parseFloat(raw) || 0) / max * 100).toFixed(1);

    // Determine "recent" threshold (7 days ago)
    const recentCutoff = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);

    // Group by sector, attach enslaved date for sorting
    const bySector = {};
    for (const m of enslaved) {
      const s = m["Sector"] || "Unknown";
      if (!bySector[s]) bySector[s] = [];
      const key = `${s}|${m["Boss"]}|${m["Minion"]}`;
      m._enslavedDate = enslavedDateLookup[key] || "";
      bySector[s].push(m);
    }

    let sectionsHtml = "";
    for (const sector of Object.keys(bySector).sort()) {
      const minions = bySector[sector];
      // Sort by enslaved date descending (newest first), undated last
      minions.sort((a, b) => (b._enslavedDate || "0000").localeCompare(a._enslavedDate || "0000"));
      let rows = "";
      for (const m of minions) {
        const nInt = norm(m["INTELLIGENCE"], totals.intel);
        const nSta = norm(m["STAMINA"], totals.stamina);
        const nTmp = norm(m["TEMPO"], totals.tempo);
        const nRep = norm(m["REPUTATION"], totals.rep);
        const nTotal = (parseFloat(nInt) + parseFloat(nSta) + parseFloat(nTmp) + parseFloat(nRep)).toFixed(1);
        const isSurvival = survivalBossKeys.has(`${sector}|${m["Boss"]}`);
        const isRecent = m._enslavedDate >= recentCutoff;
        const bossCell = isSurvival
          ? `<td class="survival-boss-cell">${shieldSvg} ${escHtml(m["Boss"])}</td>`
          : `<td>${escHtml(m["Boss"])}</td>`;
        const trClass = [isSurvival && "survival-row", isRecent && "recent-row"].filter(Boolean).join(" ");
        rows += `<tr${trClass ? ` class="${trClass}"` : ''}>
          ${bossCell}
          <td>${escHtml(m["Minion"])}</td>
          <td class="date-col">${m._enslavedDate || "—"}</td>
          <td style="color:#00f2ff">${nInt}</td>
          <td style="color:#00ff9d">${nSta}</td>
          <td style="color:#ff00ff">${nTmp}</td>
          <td style="color:#ff8800">${nRep}</td>
          <td style="color:#ffea00;font-weight:bold;">${nTotal}</td>
        </tr>`;
      }
      sectionsHtml += `
        <div class="army-sector">
          <h2>${escHtml(sector)} <span class="army-count">(${minions.length})</span></h2>
          <table>
            <thead><tr>
              <th>Boss</th><th>Minion</th><th>Enslaved</th>
              <th>INT</th><th>STA</th><th>TMP</th><th>REP</th><th>Total</th>
            </tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>`;
    }

    res.send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Henry's Army - Sovereign HUD</title>
    <style>
    body {
        background: #0a0b10;
        color: #00f2ff;
        font-family: 'Courier New', monospace;
        padding: 20px;
        text-transform: uppercase;
        overflow-x: hidden;
    }
    .hud-container {
        border: 2px solid #00ff9d;
        padding: 20px;
        box-shadow: 0 0 15px rgba(0, 255, 157, 0.5);
        max-width: 960px;
        margin: auto;
    }
    h1 {
        text-shadow: 2px 2px #ff00ff;
        letter-spacing: 2px;
        text-align: center;
        margin-bottom: 5px;
        color: #00ff9d;
    }
    .army-subtitle {
        text-align: center;
        color: #00ff9d;
        font-size: 0.85em;
        margin-bottom: 25px;
        letter-spacing: 3px;
    }
    .back-link {
        display: inline-block;
        color: #00f2ff;
        text-decoration: none;
        border: 1px solid #00f2ff;
        padding: 6px 15px;
        margin-bottom: 15px;
        font-size: 0.8em;
        transition: all 0.2s;
    }
    .back-link:hover { background: #00f2ff; color: #0a0b10; }
    .army-sector {
        margin-bottom: 25px;
        border: 1px solid rgba(0, 255, 157, 0.2);
        background: rgba(0, 255, 157, 0.02);
        border-radius: 6px;
        padding: 15px;
    }
    .army-sector h2 {
        color: #ff00ff;
        font-size: 1em;
        letter-spacing: 2px;
        margin: 0 0 12px 0;
        border-bottom: 1px solid rgba(255, 0, 255, 0.3);
        padding-bottom: 8px;
    }
    .army-count { color: #00ff9d; font-size: 0.85em; }
    table { width: 100%; border-collapse: collapse; font-size: 0.8em; }
    th {
        background: #1a1d26;
        color: #ffea00;
        padding: 8px 6px;
        text-align: left;
        border-bottom: 2px solid #00ff9d;
    }
    td { padding: 6px; border-bottom: 1px solid #1a1d26; }
    tr:hover td { background: rgba(0, 255, 157, 0.05); }
    tr.survival-row { background: rgba(255, 68, 68, 0.04); }
    tr.survival-row:hover td { background: rgba(255, 68, 68, 0.08); }
    .survival-boss-cell { color: #ff4444; text-shadow: 0 0 6px rgba(255, 68, 68, 0.3); }
    .date-col { color: #888; font-size: 0.85em; }
    tr.recent-row { background: rgba(0, 255, 157, 0.06); }
    tr.recent-row td { border-left: 0; }
    tr.recent-row td:first-child { border-left: 2px solid #00ff9d; }
    .empty-army {
        text-align: center;
        color: #555;
        padding: 40px;
        font-size: 0.9em;
    }
    .army-emblem {
        text-align: center;
        margin: 8px 0 12px 0;
        padding-bottom: 15px;
        border-bottom: 1px solid #00ff9d;
        animation: emblemPulse 4s ease-in-out infinite;
    }
    @keyframes emblemPulse {
        0%, 100% { filter: drop-shadow(0 0 6px rgba(0,242,255,0.3)); }
        50% { filter: drop-shadow(0 0 14px rgba(0,242,255,0.6)) drop-shadow(0 0 20px rgba(255,0,255,0.2)); }
    }
    @media (max-width: 600px) {
        body { padding: 10px; }
        .hud-container { padding: 15px; }
    }
    </style>
</head>
<body>
    <div class="hud-container">
        <a class="back-link" href="/">&lt; BACK TO HUD</a>
        <h1>Henry's Army</h1>
        <div class="army-emblem">
            <svg width="60" height="60" viewBox="0 0 60 60" xmlns="http://www.w3.org/2000/svg">
                <defs>
                    <linearGradient id="emblemGrad" x1="0%" y1="0%" x2="100%" y2="100%">
                        <stop offset="0%" stop-color="#00f2ff" stop-opacity="0.8"/>
                        <stop offset="50%" stop-color="#ff00ff" stop-opacity="0.6"/>
                        <stop offset="100%" stop-color="#ffea00" stop-opacity="0.8"/>
                    </linearGradient>
                    <filter id="emblemGlow">
                        <feGaussianBlur stdDeviation="2" result="blur"/>
                        <feMerge><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge>
                    </filter>
                </defs>
                <polygon points="30,2 54,16 54,44 30,58 6,44 6,16" fill="none" stroke="url(#emblemGrad)" stroke-width="1.5" filter="url(#emblemGlow)"/>
                <polygon points="30,10 46,20 46,40 30,50 14,40 14,20" fill="rgba(0,242,255,0.05)" stroke="rgba(0,242,255,0.3)" stroke-width="1"/>
                <line x1="30" y1="2" x2="30" y2="10" stroke="rgba(0,242,255,0.4)" stroke-width="1"/>
                <line x1="54" y1="16" x2="46" y2="20" stroke="rgba(255,0,255,0.4)" stroke-width="1"/>
                <line x1="54" y1="44" x2="46" y2="40" stroke="rgba(255,0,255,0.4)" stroke-width="1"/>
                <line x1="30" y1="58" x2="30" y2="50" stroke="rgba(255,234,0,0.4)" stroke-width="1"/>
                <line x1="6" y1="44" x2="14" y2="40" stroke="rgba(0,242,255,0.4)" stroke-width="1"/>
                <line x1="6" y1="16" x2="14" y2="20" stroke="rgba(0,242,255,0.4)" stroke-width="1"/>
                <circle cx="30" cy="30" r="3" fill="#00f2ff" opacity="0.6"/>
                <circle cx="30" cy="30" r="1.5" fill="#fff"/>
                <path d="M30 38 L30 22 M25 27 L30 22 L35 27" stroke="#ffea00" stroke-width="1.5" fill="none" stroke-linecap="round" stroke-linejoin="round" filter="url(#emblemGlow)"/>
            </svg>
        </div>
        <div class="army-subtitle">${enslaved.length} MINIONS ENSLAVED</div>
        ${sectionsHtml || '<div class="empty-army">NO MINIONS ENSLAVED YET. CONQUER THEM THROUGH THE QUEST BOARD.</div>'}
    </div>
</body>
</html>`);
  } catch (err) {
    console.error("Army page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// Progress & Accomplishments Page
// ---------------------------------------------------------------------------
app.get("/progress", async (req, res) => {
  try {
    // Period filter
    const period = req.query.period || "";
    const now = new Date();
    let periodStart = null;
    let periodLabel = "ALL TIME";
    if (period === "30d") {
      periodStart = new Date(now);
      periodStart.setDate(periodStart.getDate() - 30);
      periodLabel = "LAST 30 DAYS";
    } else if (period === "semester") {
      const month = now.getMonth(); // 0-indexed
      if (month >= 7) { // Aug-Dec: semester started Aug 1
        periodStart = new Date(now.getFullYear(), 7, 1);
      } else { // Jan-Jul: semester started Jan 1
        periodStart = new Date(now.getFullYear(), 0, 1);
      }
      periodLabel = "THIS SEMESTER";
    } else if (period === "year") {
      periodStart = new Date(now.getFullYear(), 0, 1);
      periodLabel = "THIS YEAR";
    } else if (period === "custom" && req.query.from) {
      periodStart = new Date(req.query.from + "T00:00:00");
      if (isNaN(periodStart.getTime())) periodStart = null;
      const customTo = req.query.to ? new Date(req.query.to + "T23:59:59") : now;
      periodLabel = `${req.query.from} TO ${(req.query.to || now.toISOString().slice(0, 10))}`;
    }
    const periodStartStr = periodStart ? periodStart.toISOString().slice(0, 10) : null;
    const periodEndStr = (period === "custom" && req.query.to) ? req.query.to : null;

    const sheets = await getSheets();
    const batchRes = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SPREADSHEET_ID,
      ranges: ["Sectors", "Command_Center", "Definitions", "Quests"],
    });
    const [sectorsRaw, ccRaw, defRaw, questsRaw] = batchRes.data.valueRanges;
    const allMinions = parseTable(sectorsRaw.values || []);
    const commandCenter = parseTable(ccRaw.values || []);
    const definitions = parseTable(defRaw.values || []);
    const quests = parseTable(questsRaw.values || []);

    // Stats
    const intel = getStat(commandCenter, "Intel");
    const stamina = getStat(commandCenter, "Stamina");
    const tempo = getStat(commandCenter, "Tempo");
    const reputation = getStat(commandCenter, "Reputation");
    const confidence = getStat(commandCenter, "Confidence");

    // Survival info
    let survivalCol = null;
    if (allMinions.length > 0) {
      survivalCol = Object.keys(allMinions[0]).find((k) => k.toLowerCase().includes("survival"));
    }
    const survivalBossKeys = new Set();
    if (survivalCol) {
      for (const r of allMinions) {
        if ((r[survivalCol] || "").toUpperCase() === "X") {
          survivalBossKeys.add(`${r["Sector"]}|${r["Boss"]}`);
        }
      }
    }

    // Boss conquest stats
    const bossMap = {};
    for (const m of allMinions) {
      const key = `${m["Sector"]}|${m["Boss"]}`;
      if (!bossMap[key]) bossMap[key] = { sector: m["Sector"], boss: m["Boss"], total: 0, enslaved: 0, isSurvival: survivalBossKeys.has(key) };
      bossMap[key].total++;
      if (m["Status"] === "Enslaved") bossMap[key].enslaved++;
    }
    const allBosses = Object.values(bossMap);
    const completedBosses = allBosses.filter((b) => b.enslaved >= b.total);
    const totalBosses = allBosses.length;
    const survivalBosses = allBosses.filter((b) => b.isSurvival);
    const completedSurvival = survivalBosses.filter((b) => b.enslaved >= b.total).length;

    // Sector summary
    const sectorStats = {};
    for (const m of allMinions) {
      const s = m["Sector"] || "Unknown";
      if (!sectorStats[s]) sectorStats[s] = { total: 0, enslaved: 0, engaged: 0, locked: 0 };
      sectorStats[s].total++;
      if (m["Status"] === "Enslaved") sectorStats[s].enslaved++;
      else if (m["Status"] === "Engaged") sectorStats[s].engaged++;
      else sectorStats[s].locked++;
    }

    // Subject summary (school-focused view) with period start/end comparison
    const subjectStats = {};
    for (const m of allMinions) {
      const subj = m["Subject"] || "Unassigned";
      if (!subjectStats[subj]) subjectStats[subj] = { total: 0, enslaved: 0, enslavedAtStart: 0, engaged: 0, locked: 0, sectors: new Set(), bosses: new Set() };
      subjectStats[subj].total++;
      subjectStats[subj].sectors.add(m["Sector"] || "");
      subjectStats[subj].bosses.add(m["Boss"] || "");
      if (m["Status"] === "Enslaved") {
        subjectStats[subj].enslaved++;
        // Count how many were already enslaved BEFORE the period started
        const completedDate = (m["Date Quest Completed"] || "").slice(0, 10);
        if (periodStartStr && completedDate && completedDate < periodStartStr) {
          subjectStats[subj].enslavedAtStart++;
        } else if (!periodStartStr) {
          // "All time" — no period, start value is 0
        }
      } else if (m["Status"] === "Engaged") subjectStats[subj].engaged++;
      else subjectStats[subj].locked++;
    }

    // Quest stats
    const approvedQuests = quests.filter((q) => q["Status"] === "Approved");
    const activeQuests = quests.filter((q) => q["Status"] === "Active");
    const submittedQuests = quests.filter((q) => q["Status"] === "Submitted");
    const totalEnslaved = allMinions.filter((m) => m["Status"] === "Enslaved").length;
    const totalMinions = allMinions.length;
    const overallPct = totalMinions > 0 ? Math.round((totalEnslaved / totalMinions) * 100) : 0;

    // Timeline entries (approved quests with dates), filtered by period
    const timelineEntries = approvedQuests
      .filter((q) => q["Date Completed"] && (!periodStartStr || q["Date Completed"].slice(0, 10) >= periodStartStr) && (!periodEndStr || q["Date Completed"].slice(0, 10) <= periodEndStr))
      .map((q) => ({
        date: q["Date Completed"],
        boss: q["Boss"],
        minion: q["Minion"],
        sector: q["Sector"],
        type: q["Proof Type"] || "",
        feedback: q["Feedback"] || "",
      }))
      .sort((a, b) => b.date.localeCompare(a.date)); // newest first

    // -- SUMMARY SPARKLINES --
    const today = new Date();
    const sparklineEnd = periodEndStr ? new Date(periodEndStr + "T00:00:00") : today;
    // Determine sparkline date range: use period or default 30 days
    const sparklineStart = periodStart || new Date(sparklineEnd.getFullYear(), sparklineEnd.getMonth(), sparklineEnd.getDate() - 29);
    const sparklineDays = Math.min(Math.max(1, Math.round((sparklineEnd - sparklineStart) / (1000 * 60 * 60 * 24))), 365) + 1;
    const maxBars = 30;
    const bucketSize = sparklineDays > maxBars ? Math.ceil(sparklineDays / maxBars) : 1;
    const numBuckets = Math.ceil(sparklineDays / bucketSize);

    const days30 = [];
    for (let i = numBuckets - 1; i >= 0; i--) {
      const d = new Date(sparklineEnd);
      d.setDate(d.getDate() - i * bucketSize);
      days30.push(d.toISOString().slice(0, 10));
    }

    // Collect completion dates for enslaved minions
    const minionCompleteDates = allMinions
      .filter((m) => m["Status"] === "Enslaved" && m["Date Quest Completed"])
      .map((m) => m["Date Quest Completed"].slice(0, 10))
      .sort();

    // Quest completion dates (approved)
    const questCompleteDates = approvedQuests
      .filter((q) => q["Date Completed"])
      .map((q) => q["Date Completed"].slice(0, 10))
      .sort();

    // Helper: count items <= date (cumulative)
    const countUpTo = (sortedDates, date) => {
      let count = 0;
      for (const d of sortedDates) { if (d <= date) count++; else break; }
      return count;
    };

    // Build daily series for each metric
    // 1. Minions Enslaved (cumulative by completion date)
    const enslavedByDay = days30.map((d) => countUpTo(minionCompleteDates, d));
    // For minions without dates, they exist but aren't in minionCompleteDates.
    // Current total = totalEnslaved, so scale: if no dates, show flat at current value
    const enslavedWithDates = minionCompleteDates.length;
    const enslavedNoDates = totalEnslaved - enslavedWithDates;
    const enslavedSeries = enslavedByDay.map((v) => v + enslavedNoDates);

    // 2. Overall Conquest % (derived)
    const conquestSeries = enslavedSeries.map((v) => totalMinions > 0 ? Math.round((v / totalMinions) * 100) : 0);

    // 3. Bosses Conquered — need per-boss completion date (latest minion date for fully conquered bosses)
    const bossCompleteDates = [];
    for (const b of allBosses) {
      if (b.enslaved < b.total) continue;
      // Find latest Date Quest Completed among this boss's minions
      const bossMinions = allMinions.filter((m) => m["Sector"] === b.sector && m["Boss"] === b.boss && m["Status"] === "Enslaved");
      const dates = bossMinions.map((m) => (m["Date Quest Completed"] || "").slice(0, 10)).filter(Boolean);
      if (dates.length > 0) bossCompleteDates.push(dates.sort().pop());
    }
    bossCompleteDates.sort();
    const bossesWithDates = bossCompleteDates.length;
    const bossesNoDates = completedBosses.length - bossesWithDates;
    const bossSeries = days30.map((d) => countUpTo(bossCompleteDates, d) + bossesNoDates);

    // 4. Guardians Defeated — same logic for survival bosses
    const guardianCompleteDates = [];
    for (const b of survivalBosses) {
      if (b.enslaved < b.total) continue;
      const bossMinions = allMinions.filter((m) => m["Sector"] === b.sector && m["Boss"] === b.boss && m["Status"] === "Enslaved");
      const dates = bossMinions.map((m) => (m["Date Quest Completed"] || "").slice(0, 10)).filter(Boolean);
      if (dates.length > 0) guardianCompleteDates.push(dates.sort().pop());
    }
    guardianCompleteDates.sort();
    const guardiansWithDates = guardianCompleteDates.length;
    const guardiansNoDates = completedSurvival - guardiansWithDates;
    const guardianSeries = days30.map((d) => countUpTo(guardianCompleteDates, d) + guardiansNoDates);

    // 5. Quests Completed (cumulative approved)
    const questsWithDates = questCompleteDates.length;
    const questsNoDates = approvedQuests.length - questsWithDates;
    const questCompleteSeries = days30.map((d) => countUpTo(questCompleteDates, d) + questsNoDates);

    // 6. Quests In Progress — historical active count using Date Added / Date Resolved
    const questInProgressCurrent = activeQuests.length + submittedQuests.length;
    // Count active/submitted quests missing Date Added — they exist but we don't know when they started
    const undatedActiveCount = quests.filter((q) =>
      (q["Status"] === "Active" || q["Status"] === "Submitted") && !q["Date Added"]
    ).length;
    const questInProgressSeries = days30.map((day) => {
      let count = undatedActiveCount;
      for (const q of quests) {
        const added = (q["Date Added"] || "").slice(0, 10);
        if (!added || added > day) continue;
        const resolved = (q["Date Resolved"] || "").slice(0, 10);
        if (!resolved || resolved > day) count++;
      }
      return count;
    });

    // 7. Guardian Minion Enslaved % (for Survival Mode section)
    const totalGuardianMinions = allMinions.filter((m) => {
      return survivalBosses.some((b) => b.sector === m["Sector"] && b.boss === m["Boss"]);
    }).length;
    const guardianMinionCompleteDates = allMinions
      .filter((m) => {
        if (m["Status"] !== "Enslaved") return false;
        return survivalBosses.some((b) => b.sector === m["Sector"] && b.boss === m["Boss"]);
      })
      .map((m) => (m["Date Quest Completed"] || "").slice(0, 10))
      .filter(Boolean)
      .sort();
    const guardianMinionsWithDates = guardianMinionCompleteDates.length;
    const guardianMinionEnslaved = allMinions.filter((m) => {
      return m["Status"] === "Enslaved" && survivalBosses.some((b) => b.sector === m["Sector"] && b.boss === m["Boss"]);
    }).length;
    const guardianMinionsNoDates = guardianMinionEnslaved - guardianMinionsWithDates;
    const guardianMinionPctSeries = days30.map((d) => {
      const count = countUpTo(guardianMinionCompleteDates, d) + guardianMinionsNoDates;
      return totalGuardianMinions > 0 ? Math.round((count / totalGuardianMinions) * 100) : 0;
    });

    // 8. Boss Conquest % (all bosses, % fully conquered by day)
    const bossConquestPctSeries = days30.map((d) => {
      const conqueredByDay = countUpTo(bossCompleteDates, d) + bossesNoDates;
      return totalBosses > 0 ? Math.round((conqueredByDay / totalBosses) * 100) : 0;
    });

    // Nice max: data max + 10%, rounded to a clean number
    function niceMax(rawMax) {
      if (rawMax <= 0) return 1;
      const target = rawMax * 1.1;
      if (target <= 5) return Math.ceil(target);
      if (target <= 20) return Math.ceil(target / 2) * 2;
      if (target <= 100) return Math.ceil(target / 5) * 5;
      if (target <= 500) return Math.ceil(target / 25) * 25;
      return Math.ceil(target / 50) * 50;
    }

    // SVG daily bar chart builder for progress sections (larger format with title + y-axis)
    function buildDailyChart(series, color, chartTitle, yAxisLabel, maxOverride) {
      const chartW = 700;
      const chartH = 180;
      const padL = 45;
      const padR = 10;
      const padT = 30;
      const padB = 22;
      const plotW = chartW - padL - padR;
      const plotH = chartH - padT - padB;
      const maxV = maxOverride != null ? maxOverride : niceMax(Math.max(...series, 1));
      const barW = plotW / series.length;

      let svg = `<svg width="100%" viewBox="0 0 ${chartW} ${chartH}" xmlns="http://www.w3.org/2000/svg" style="max-width:${chartW}px;">`;

      // Title
      svg += `<text x="${chartW / 2}" y="18" fill="${color}" font-size="16" font-weight="bold" text-anchor="middle" font-family="'Courier New',monospace" letter-spacing="2" style="text-shadow:0 0 6px ${color};">${chartTitle}</text>`;

      // Y-axis label (vertical)
      svg += `<text x="10" y="${padT + plotH / 2}" fill="#aaa" font-size="12" text-anchor="middle" font-family="'Courier New',monospace" transform="rotate(-90, 10, ${padT + plotH / 2})">${yAxisLabel}</text>`;

      // Grid lines + y labels
      for (let i = 0; i <= 4; i++) {
        const y = padT + plotH - (i / 4) * plotH;
        const val = Math.round((i / 4) * maxV);
        svg += `<line x1="${padL}" y1="${y}" x2="${chartW - padR}" y2="${y}" stroke="rgba(255,255,255,0.08)" stroke-width="1"/>`;
        svg += `<text x="${padL - 5}" y="${y + 4}" fill="#aaa" font-size="11" text-anchor="end" font-family="monospace">${val}</text>`;
      }

      // Bars
      for (let i = 0; i < series.length; i++) {
        const barH = maxV > 0 ? (series[i] / maxV) * plotH : 0;
        const x = padL + i * barW;
        const y = padT + plotH - barH;
        const opacity = i === series.length - 1 ? 1 : 0.6;
        svg += `<rect x="${x}" y="${y}" width="${Math.max(barW - 1, 2)}" height="${barH}" fill="${color}" opacity="${opacity}" rx="1">`;
        svg += `<title>${days30[i]}: ${series[i]}${maxOverride === 100 ? '%' : ''}</title></rect>`;
      }

      // X-axis labels (every 5 days)
      for (let i = 0; i < series.length; i += 5) {
        const x = padL + i * barW + barW / 2;
        svg += `<text x="${x}" y="${chartH - 5}" fill="#999" font-size="10" text-anchor="middle" font-family="monospace">${days30[i].slice(5)}</text>`;
      }
      // Last day label
      const lastX = padL + (series.length - 1) * barW + barW / 2;
      svg += `<text x="${lastX}" y="${chartH - 5}" fill="#bbb" font-size="10" text-anchor="middle" font-family="monospace">${days30[series.length - 1].slice(5)}</text>`;

      svg += `</svg>`;
      return svg;
    }

    // SVG sparkline bar builder with title, subtitle, y-axis labels, and +10% headroom
    function buildSparkBars(series, color, title, subtitle, unit) {
      const w = 280;
      const h = subtitle ? 115 : 105;
      const titleH = title ? 18 : 0;
      const subH = subtitle ? 14 : 0;
      const padT = titleH + subH + 4;
      const padL = 38;
      const padR = 4;
      const padB = 16;
      const plotW = w - padL - padR;
      const plotH = h - padT - padB;
      const rawMax = Math.max(...series, 1);
      const maxV = niceMax(rawMax);
      const barW = plotW / series.length;
      const suffix = unit || "";
      let svg = `<svg width="100%" viewBox="0 0 ${w} ${h}" style="display:block;margin:6px auto 0;max-width:${w}px;">`;
      if (title) {
        svg += `<text x="${padL + plotW / 2}" y="15" fill="${color}" font-size="14" font-weight="bold" text-anchor="middle" font-family="'Courier New',monospace" letter-spacing="2" style="text-shadow:0 0 6px ${color};">${title}</text>`;
      }
      if (subtitle) {
        svg += `<text x="${padL + plotW / 2}" y="${titleH + 12}" fill="#bbb" font-size="10" text-anchor="middle" font-family="'Courier New',monospace" letter-spacing="1.5">${subtitle}</text>`;
      }
      // Y-axis: max, mid, 0
      svg += `<text x="${padL - 4}" y="${padT + 8}" fill="#ccc" font-size="10" text-anchor="end" font-family="monospace">${maxV}${suffix}</text>`;
      const midVal = Math.round(maxV / 2);
      svg += `<text x="${padL - 4}" y="${padT + plotH / 2 + 3}" fill="#aaa" font-size="10" text-anchor="end" font-family="monospace">${midVal}${suffix}</text>`;
      svg += `<text x="${padL - 4}" y="${padT + plotH}" fill="#aaa" font-size="10" text-anchor="end" font-family="monospace">0</text>`;
      for (let i = 0; i < series.length; i++) {
        const barH = maxV > 0 ? (series[i] / maxV) * plotH : 0;
        const x = padL + i * barW;
        const y = padT + plotH - barH;
        const opacity = i === series.length - 1 ? 1 : 0.5;
        svg += `<rect x="${x}" y="${y}" width="${Math.max(barW - 0.5, 1)}" height="${barH}" fill="${color}" opacity="${opacity}" rx="0.5"><title>${days30[i]}: ${series[i]}${suffix}</title></rect>`;
      }
      // X-axis label
      svg += `<text x="${padL + plotW / 2}" y="${h - 2}" fill="#aaa" font-size="9" text-anchor="middle" font-family="monospace" letter-spacing="1">${periodLabel}</text>`;
      svg += `</svg>`;
      return svg;
    }

    // Build stat growth timeline chart
    // Collect stat contributions from enslaved minions with completion dates
    const statNames = ["INTELLIGENCE", "STAMINA", "TEMPO", "REPUTATION"];
    const statLabels = { INTELLIGENCE: "INTEL", STAMINA: "STAMINA", TEMPO: "TEMPO", REPUTATION: "REPUTATION" };
    const statColors = { INTELLIGENCE: "#00f2ff", STAMINA: "#00ff9d", TEMPO: "#ff00ff", REPUTATION: "#ff8800" };

    // Build sector weight lookup from definitions
    const sectorWeightsMap = {};
    for (const d of definitions) {
      if (d["Sector"]) {
        sectorWeightsMap[d["Sector"].toUpperCase()] = {
          INTELLIGENCE: parseFloat(d["INTELLIGENCE"]) || 0,
          STAMINA: parseFloat(d["STAMINA"]) || 0,
          TEMPO: parseFloat(d["TEMPO"]) || 0,
          REPUTATION: parseFloat(d["REPUTATION"]) || 0,
        };
      }
    }

    // Collect dated stat events from enslaved minions (filtered by period)
    const statEvents = [];
    for (const m of allMinions) {
      if (m["Status"] !== "Enslaved") continue;
      const dateStr = m["Date Quest Completed"] || "";
      if (!dateStr || dateStr.length < 10) continue; // need YYYY-MM-DD
      if (periodStartStr && dateStr.slice(0, 10) < periodStartStr) continue;
      if (periodEndStr && dateStr.slice(0, 10) > periodEndStr) continue;
      const impact = parseFloat(m["Impact(1-3)"]) || 1;
      const sw = sectorWeightsMap[(m["Sector"] || "").toUpperCase()];
      if (!sw) continue;
      for (const stat of statNames) {
        statEvents.push({ date: dateStr.slice(0, 10), stat, points: impact * sw[stat] });
      }
    }

    // Determine time range and bucket size
    let statBarsHtml = "";
    if (statEvents.length === 0) {
      // Show 7-day chart using current stat values as flat baseline
      const curStatVals = {
        INTELLIGENCE: intel.value,
        STAMINA: stamina.value,
        TEMPO: tempo.value,
        REPUTATION: reputation.value,
      };
      const emptyChartW = 800, emptyChartH = 220;
      const ePadL = 50, ePadB = 35, ePadT = 30, ePadR = 100;
      const ePlotW = emptyChartW - ePadL - ePadR;
      const ePlotH = emptyChartH - ePadB - ePadT;
      const eGroupW = ePlotW / 7;
      const eBarW = Math.max(3, Math.min(16, (eGroupW - 4) / statNames.length));
      const eMaxRaw = Math.max(...statNames.map(s => curStatVals[s]), 1);
      const eMaxV = niceMax(eMaxRaw);
      let eSvg = `<svg width="100%" viewBox="0 0 ${emptyChartW} ${emptyChartH}" xmlns="http://www.w3.org/2000/svg" style="max-width:${emptyChartW}px;">`;
      // Title
      eSvg += `<text x="${emptyChartW / 2}" y="20" fill="#888" font-size="14" font-weight="bold" text-anchor="middle" font-family="'Courier New',monospace" letter-spacing="2">STAT LEVELS — 7 DAY TREND</text>`;
      // Grid lines
      for (let i = 0; i <= 4; i++) {
        const y = ePadT + ePlotH - (i / 4) * ePlotH;
        const val = Math.round((i / 4) * eMaxV);
        eSvg += `<line x1="${ePadL}" y1="${y}" x2="${emptyChartW - ePadR}" y2="${y}" stroke="rgba(255,255,255,0.06)" stroke-width="1"/>`;
        eSvg += `<text x="${ePadL - 5}" y="${y + 4}" fill="#aaa" font-size="11" text-anchor="end" font-family="monospace">${val}</text>`;
      }
      // Bars with current stat values for each of the last 7 days
      for (let di = 0; di < 7; di++) {
        const dayDate = new Date(); dayDate.setDate(dayDate.getDate() - 6 + di);
        const label = String(dayDate.getMonth() + 1).padStart(2, "0") + "-" + String(dayDate.getDate()).padStart(2, "0");
        const groupX = ePadL + di * eGroupW + 2;
        for (let si = 0; si < statNames.length; si++) {
          const stat = statNames[si];
          const val = curStatVals[stat];
          const barH = eMaxV > 0 ? (val / eMaxV) * ePlotH : 0;
          const x = groupX + si * (eBarW + 1);
          const y = ePadT + ePlotH - barH;
          eSvg += `<rect x="${x}" y="${y}" width="${eBarW}" height="${barH}" fill="${statColors[stat]}" opacity="0.85" rx="1">`;
          eSvg += `<title>${statLabels[stat]}: ${val.toFixed(1)}</title></rect>`;
        }
        const labelX = groupX + (statNames.length * (eBarW + 1)) / 2;
        eSvg += `<text x="${labelX}" y="${emptyChartH - 8}" fill="#999" font-size="10" text-anchor="middle" font-family="monospace">${label}</text>`;
      }
      // Legend (stacked vertically on right, outside plot area)
      const eLegX = emptyChartW - ePadR + 5;
      for (let si = 0; si < statNames.length; si++) {
        const ly = ePadT + 8 + si * 18;
        eSvg += `<rect x="${eLegX}" y="${ly}" width="10" height="10" fill="${statColors[statNames[si]]}" rx="1"/>`;
        eSvg += `<text x="${eLegX + 14}" y="${ly + 9}" fill="${statColors[statNames[si]]}" font-size="10" font-family="monospace">${statLabels[statNames[si]]}</text>`;
      }
      eSvg += `</svg>`;
      const curSummary = statNames.map((s) =>
        `<span style="color:${statColors[s]}">${curStatVals[s].toFixed(1)} ${statLabels[s]}</span>`
      ).join(' <span style="color:#333;">|</span> ');
      statBarsHtml = `
        <div style="text-align:center;color:#888;font-size:0.65em;letter-spacing:2px;margin-bottom:8px;">CURRENT LEVELS (7 DAYS)</div>
        <div style="overflow-x:auto;">${eSvg}</div>
        <div style="text-align:center;font-size:0.7em;letter-spacing:1px;margin-top:10px;color:#888;">CURRENT: ${curSummary}</div>`;
    } else {
      // Find date range
      const allDates = statEvents.map((e) => e.date).sort();
      const minDate = new Date(allDates[0]);
      const maxDate = new Date(allDates[allDates.length - 1]);
      const daySpan = Math.max(1, Math.round((maxDate - minDate) / (1000 * 60 * 60 * 24)));

      // Choose bucket: daily ≤7 days, weekly ≤60 days, monthly otherwise
      let bucketMode, bucketLabel;
      if (daySpan <= 7) { bucketMode = "day"; bucketLabel = "DAILY"; }
      else if (daySpan <= 60) { bucketMode = "week"; bucketLabel = "WEEKLY"; }
      else { bucketMode = "month"; bucketLabel = "MONTHLY"; }

      const toBucket = (dateStr) => {
        const d = new Date(dateStr);
        if (bucketMode === "day") return dateStr.slice(0, 10);
        if (bucketMode === "week") {
          const day = d.getDay();
          const weekStart = new Date(d);
          weekStart.setDate(d.getDate() - day);
          return weekStart.toISOString().slice(0, 10);
        }
        return dateStr.slice(0, 7); // YYYY-MM
      };

      const formatBucket = (b) => {
        if (bucketMode === "day") return b.slice(5); // MM-DD
        if (bucketMode === "week") return "W" + b.slice(5); // WMM-DD
        return b; // YYYY-MM
      };

      // Accumulate points EARNED per bucket per stat (not cumulative)
      const bucketTotals = {}; // { bucket: { INTELLIGENCE: pts, ... } }
      for (const e of statEvents) {
        const b = toBucket(e.date);
        if (!bucketTotals[b]) bucketTotals[b] = { INTELLIGENCE: 0, STAMINA: 0, TEMPO: 0, REPUTATION: 0 };
        bucketTotals[b][e.stat] += e.points;
      }

      const buckets = Object.keys(bucketTotals).sort();
      const periodData = buckets.map((b) => ({ bucket: b, ...bucketTotals[b] }));

      // Find max value across any single period for scaling
      let maxVal = 0;
      for (const p of periodData) {
        for (const stat of statNames) {
          if (p[stat] > maxVal) maxVal = p[stat];
        }
      }
      maxVal = niceMax(maxVal || 1);

      // Total points earned across all periods (for summary)
      const totalEarned = { INTELLIGENCE: 0, STAMINA: 0, TEMPO: 0, REPUTATION: 0 };
      for (const p of periodData) {
        for (const stat of statNames) totalEarned[stat] += p[stat];
      }

      // Build SVG bar chart — points earned per period
      const chartW = 800;
      const chartH = 220;
      const padL = 50;
      const padB = 35;
      const padT = 10;
      const padR = 100;
      const plotW = chartW - padL - padR;
      const plotH = chartH - padB - padT;
      const barGroupW = periodData.length > 0 ? plotW / periodData.length : plotW;
      const barW = Math.max(3, Math.min(16, (barGroupW - 4) / statNames.length));

      let svg = `<svg width="100%" viewBox="0 0 ${chartW} ${chartH}" xmlns="http://www.w3.org/2000/svg" style="max-width:${chartW}px;">`;
      // Grid lines
      for (let i = 0; i <= 4; i++) {
        const y = padT + plotH - (i / 4) * plotH;
        const val = ((i / 4) * maxVal).toFixed(0);
        svg += `<line x1="${padL}" y1="${y}" x2="${chartW - padR}" y2="${y}" stroke="rgba(255,255,255,0.08)" stroke-width="1"/>`;
        svg += `<text x="${padL - 5}" y="${y + 4}" fill="#aaa" font-size="11" text-anchor="end" font-family="monospace">${val}</text>`;
      }

      // Bars — points earned per period (not cumulative)
      for (let ci = 0; ci < periodData.length; ci++) {
        const p = periodData[ci];
        const groupX = padL + ci * barGroupW + 2;
        for (let si = 0; si < statNames.length; si++) {
          const stat = statNames[si];
          const h = (p[stat] / maxVal) * plotH;
          const x = groupX + si * (barW + 1);
          const y = padT + plotH - h;
          svg += `<rect x="${x}" y="${y}" width="${barW}" height="${h}" fill="${statColors[stat]}" opacity="0.85" rx="1">`;
          svg += `<title>${statLabels[stat]}: +${p[stat].toFixed(1)} pts</title></rect>`;
        }
        // Bucket label
        const labelX = groupX + (statNames.length * (barW + 1)) / 2;
        svg += `<text x="${labelX}" y="${chartH - 8}" fill="#999" font-size="10" text-anchor="middle" font-family="monospace">${formatBucket(p.bucket)}</text>`;
      }

      // Legend (stacked vertically on right)
      const legendX = chartW - padR + 10;
      for (let si = 0; si < statNames.length; si++) {
        const ly = padT + 8 + si * 18;
        svg += `<rect x="${legendX}" y="${ly}" width="10" height="10" fill="${statColors[statNames[si]]}" rx="1"/>`;
        svg += `<text x="${legendX + 14}" y="${ly + 9}" fill="${statColors[statNames[si]]}" font-size="10" font-family="monospace">${statLabels[statNames[si]]}</text>`;
      }

      svg += `</svg>`;

      // Summary: total points earned
      const earnedSummary = statNames.map((s) =>
        `<span style="color:${statColors[s]}">+${totalEarned[s].toFixed(1)} ${statLabels[s]}</span>`
      ).join(' <span style="color:#333;">|</span> ');

      statBarsHtml = `
        <div style="text-align:center;color:#888;font-size:0.65em;letter-spacing:2px;margin-bottom:8px;">POINTS EARNED PER ${bucketMode.toUpperCase()} (${periodData.length} ${bucketMode === "day" ? "DAYS" : bucketMode === "week" ? "WEEKS" : "MONTHS"})</div>
        <div style="overflow-x:auto;">${svg}</div>
        <div style="text-align:center;font-size:0.7em;letter-spacing:1px;margin-top:10px;color:#888;">TOTAL EARNED: ${earnedSummary}</div>`;
    }

    // Sector donut chart data (pure CSS)
    let sectorChartHtml = "";
    const sectorNames = Object.keys(sectorStats).sort();
    const sectorColors = ["#00f2ff", "#00ff9d", "#ff00ff", "#ff8800", "#ffea00", "#ff4444", "#aa88ff", "#88ff88"];
    for (let i = 0; i < sectorNames.length; i++) {
      const s = sectorNames[i];
      const stat = sectorStats[s];
      const pct = stat.total > 0 ? Math.round((stat.enslaved / stat.total) * 100) : 0;
      const color = sectorColors[i % sectorColors.length];
      sectorChartHtml += `
        <div class="sector-prog-item">
          <div class="sector-prog-ring" style="background: conic-gradient(${color} 0% ${pct}%, #1a1d26 ${pct}% 100%);"></div>
          <div class="sector-prog-label">${escHtml(s)}</div>
          <div class="sector-prog-pct" style="color:${color};">${pct}%</div>
          <div class="sector-prog-detail">${stat.enslaved}/${stat.total}</div>
        </div>`;
    }

    // Subject progress table (school-focused) — with period start/end comparison
    const subjectNames = Object.keys(subjectStats).sort();
    const hasPeriod = !!periodStartStr;
    let subjectTableHtml = "";
    for (const subj of subjectNames) {
      const st = subjectStats[subj];
      // Skip subjects with zero enslaved for the entire period
      if (st.enslaved === 0 && st.enslavedAtStart === 0) continue;
      const startPct = st.total > 0 ? Math.round((st.enslavedAtStart / st.total) * 100) : 0;
      const endPct = st.total > 0 ? Math.round((st.enslaved / st.total) * 100) : 0;
      const gained = st.enslaved - st.enslavedAtStart;
      const endColor = endPct >= 100 ? "#00ff9d" : endPct >= 50 ? "#ffea00" : "#00f2ff";
      subjectTableHtml += `
        <div class="subj-prog-row">
          <div class="subj-prog-name">${escHtml(subj)}</div>
          <div class="subj-prog-bars">
            ${hasPeriod ? `<div class="subj-bar-pair">
              <div class="subj-prog-bar"><div class="subj-prog-fill" style="width:${startPct}%;background:#555;"></div></div>
              <div class="subj-bar-label" style="color:#888;">START ${startPct}%</div>
            </div>` : ""}
            <div class="subj-bar-pair">
              <div class="subj-prog-bar"><div class="subj-prog-fill" style="width:${endPct}%;background:${endColor};"></div></div>
              <div class="subj-bar-label" style="color:${endColor};">${hasPeriod ? "NOW" : ""} ${endPct}%<span style="color:#888;margin-left:6px;">${st.enslaved}/${st.total}</span>${hasPeriod && gained > 0 ? `<span style="color:#00ff9d;margin-left:6px;">+${gained}</span>` : ""}</div>
            </div>
          </div>
          <div class="subj-prog-meta">${st.bosses.size} TOPIC${st.bosses.size !== 1 ? "S" : ""}</div>
        </div>`;
    }

    // Boss conquest daily % chart
    const currentBossPct = totalBosses > 0 ? Math.round((completedBosses.length / totalBosses) * 100) : 0;
    const bossConquestHtml = `
      <div style="text-align:center;color:#888;font-size:0.65em;margin-bottom:6px;">${completedBosses.length}/${totalBosses} BOSSES FULLY CONQUERED (${currentBossPct}%)</div>
      <div style="overflow-x:auto;">${buildDailyChart(bossConquestPctSeries, "#ff6600", "BOSS CONQUEST — DAILY %", "% CONQUERED", 100)}</div>`;

    // Survival mode daily % chart
    const currentGuardianMinionPct = totalGuardianMinions > 0 ? Math.round((guardianMinionEnslaved / totalGuardianMinions) * 100) : 0;
    let survivalHtml = "";
    if (survivalBosses.length > 0) {
      survivalHtml = `
        <div style="text-align:center;color:#888;font-size:0.65em;margin-bottom:6px;">${guardianMinionEnslaved}/${totalGuardianMinions} GUARDIAN MINIONS ENSLAVED (${currentGuardianMinionPct}%)</div>
        <div style="overflow-x:auto;">${buildDailyChart(guardianMinionPctSeries, "#ff4444", "GUARDIAN MINIONS ENSLAVED — DAILY %", "% ENSLAVED", 100)}</div>`;
    }

    // Timeline
    let timelineHtml = "";
    if (timelineEntries.length === 0) {
      timelineHtml = `<div class="timeline-empty">NO COMPLETED QUESTS YET. START CONQUERING!</div>`;
    } else {
      for (const e of timelineEntries.slice(0, 30)) {
        timelineHtml += `
          <div class="timeline-entry">
            <div class="timeline-date">${escHtml(e.date)}</div>
            <div class="timeline-dot"></div>
            <div class="timeline-content">
              <span class="timeline-minion">${escHtml(e.minion)}</span>
              <span class="timeline-arrow">&rarr;</span>
              <span class="timeline-boss">${escHtml(e.boss)}</span>
              <span class="timeline-sector">${escHtml(e.sector)}</span>
              ${e.type ? `<span class="timeline-type">${escHtml(e.type)}</span>` : ''}
              ${e.feedback ? `<div class="timeline-feedback">"${escHtml(e.feedback)}"</div>` : ''}
            </div>
          </div>`;
      }
    }

    res.send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Progress Report - Sovereign HUD</title>
    <style>
    body {
        background: #0a0b10;
        color: #00f2ff;
        font-family: 'Courier New', monospace;
        padding: 20px;
        text-transform: uppercase;
        overflow-x: hidden;
    }
    .hud-container {
        max-width: 960px;
        margin: auto;
    }
    .back-link { display: inline-block; color: #00f2ff; text-decoration: none; border: 1px solid #00f2ff; padding: 6px 15px; margin-bottom: 15px; font-size: 0.8em; transition: all 0.2s; }
    .back-link:hover { background: #00f2ff; color: #0a0b10; }
    h1 {
        text-align: center;
        color: #ffea00;
        text-shadow: 2px 2px #ff00ff;
        letter-spacing: 4px;
        border-bottom: 2px solid #ffea00;
        padding-bottom: 10px;
    }
    .progress-subtitle {
        text-align: center;
        color: #888;
        font-size: 0.75em;
        letter-spacing: 3px;
        margin-bottom: 30px;
    }

    /* Summary cards */
    .summary-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 14px;
        margin-bottom: 30px;
        border: 1px solid rgba(0, 242, 255, 0.12);
        background: rgba(0, 242, 255, 0.02);
        border-radius: 10px;
        padding: 18px;
    }
    .summary-card {
        border: 1px solid rgba(0, 242, 255, 0.18);
        background: rgba(0, 10, 20, 0.5);
        border-radius: 8px;
        padding: 10px 6px 4px;
        text-align: center;
    }

    /* Sections */
    .progress-section {
        border: 1px solid rgba(0, 242, 255, 0.15);
        background: rgba(0, 242, 255, 0.02);
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 25px;
    }
    .progress-section h2 {
        color: #ff00ff;
        font-size: 0.9em;
        letter-spacing: 3px;
        margin: 0 0 15px 0;
        padding-bottom: 8px;
        border-bottom: 1px solid rgba(255, 0, 255, 0.3);
    }

    /* Stat bars */
    .prog-stat {
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 10px;
    }
    .prog-stat-label {
        width: 220px;
        font-size: 0.75em;
        font-weight: bold;
        letter-spacing: 1px;
        flex-shrink: 0;
    }
    .prog-stat-bar {
        flex: 1;
        height: 14px;
        background: #1a1d26;
        border-radius: 2px;
        overflow: hidden;
    }
    .prog-stat-fill {
        height: 100%;
        border-radius: 2px;
        transition: width 0.5s;
    }
    .prog-stat-pct {
        width: 40px;
        text-align: right;
        font-size: 0.75em;
        font-weight: bold;
    }

    /* Sector donut rings */
    .sector-prog-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(100px, 1fr));
        gap: 15px;
        text-align: center;
    }
    .sector-prog-item { display: flex; flex-direction: column; align-items: center; }
    .sector-prog-ring {
        width: 70px; height: 70px; border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        position: relative;
    }
    .sector-prog-ring::after {
        content: '';
        position: absolute;
        width: 40px; height: 40px;
        background: #0a0b10;
        border-radius: 50%;
    }
    .sector-prog-label { font-size: 0.6em; color: #888; margin-top: 6px; letter-spacing: 1px; }
    .sector-prog-pct { font-size: 0.85em; font-weight: bold; }
    .sector-prog-detail { font-size: 0.6em; color: #555; }

    /* Subject progress bars */
    .subj-prog-row {
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 12px;
        padding-bottom: 12px;
        border-bottom: 1px solid rgba(0,242,255,0.08);
    }
    .subj-prog-row:last-child { border-bottom: none; margin-bottom: 0; }
    .subj-prog-name {
        width: 140px;
        font-size: 0.75em;
        color: #ccc;
        font-weight: bold;
        letter-spacing: 1px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        flex-shrink: 0;
    }
    .subj-prog-bars {
        flex: 1;
        display: flex;
        flex-direction: column;
        gap: 3px;
    }
    .subj-bar-pair {
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .subj-prog-bar {
        flex: 1;
        height: 10px;
        background: #1a1d26;
        border-radius: 2px;
        overflow: hidden;
    }
    .subj-prog-fill {
        height: 100%;
        border-radius: 2px;
        transition: width 0.5s;
    }
    .subj-bar-label {
        font-size: 0.65em;
        white-space: nowrap;
        min-width: 100px;
        letter-spacing: 1px;
    }
    .subj-prog-meta {
        width: 65px;
        text-align: right;
        font-size: 0.6em;
        color: #555;
    }

    /* Boss conquest bars */
    .boss-prog-row {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 6px;
    }
    .boss-prog-name {
        width: 180px;
        font-size: 0.7em;
        color: #ccc;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        flex-shrink: 0;
    }
    .boss-prog-bar {
        flex: 1;
        height: 10px;
        background: #1a1d26;
        border-radius: 2px;
        overflow: hidden;
    }
    .boss-prog-fill {
        height: 100%;
        border-radius: 2px;
    }
    .boss-prog-pct {
        width: 35px;
        text-align: right;
        font-size: 0.7em;
        color: #888;
    }

    /* Timeline */
    .timeline {
        position: relative;
        padding-left: 20px;
    }
    .timeline::before {
        content: '';
        position: absolute;
        left: 8px;
        top: 0;
        bottom: 0;
        width: 2px;
        background: rgba(255, 0, 255, 0.3);
    }
    .timeline-entry {
        display: flex;
        align-items: flex-start;
        gap: 12px;
        margin-bottom: 12px;
        position: relative;
    }
    .timeline-dot {
        width: 10px;
        height: 10px;
        border-radius: 50%;
        background: #ff00ff;
        flex-shrink: 0;
        margin-top: 3px;
        box-shadow: 0 0 6px rgba(255, 0, 255, 0.5);
        position: absolute;
        left: -16px;
    }
    .timeline-date {
        font-size: 0.65em;
        color: #555;
        width: 80px;
        flex-shrink: 0;
    }
    .timeline-content { font-size: 0.8em; }
    .timeline-minion { color: #00f2ff; font-weight: bold; }
    .timeline-arrow { color: #ff00ff; margin: 0 4px; }
    .timeline-boss { color: #ff6600; }
    .timeline-sector { color: #888; font-size: 0.85em; margin-left: 6px; }
    .timeline-type { color: #ffea00; font-size: 0.8em; margin-left: 6px; border: 1px solid rgba(255, 234, 0, 0.3); padding: 1px 5px; }
    .timeline-feedback { color: #ffea00; font-size: 0.85em; margin-top: 3px; font-style: italic; text-transform: none; }
    .timeline-empty { text-align: center; color: #555; padding: 20px; }

    /* Period filter toolbar */
    .period-toolbar {
        display: flex; align-items: center; justify-content: center; gap: 8px;
        margin-bottom: 25px; flex-wrap: wrap;
    }
    .period-btn {
        padding: 6px 14px; font-family: 'Courier New', monospace; font-size: 0.7em;
        letter-spacing: 1px; text-decoration: none; border: 1px solid #444;
        color: #888; transition: all 0.2s; text-transform: uppercase;
    }
    .period-btn:hover { border-color: #00f2ff; color: #00f2ff; }
    .period-btn.active { border-color: #ffea00; color: #ffea00; background: rgba(255,234,0,0.1); }
    .print-btn {
        padding: 6px 14px; font-family: 'Courier New', monospace; font-size: 0.7em;
        letter-spacing: 1px; border: 1px solid #ff00ff; color: #ff00ff;
        background: none; cursor: pointer; transition: all 0.2s; text-transform: uppercase;
    }
    .print-btn:hover { background: #ff00ff; color: #0a0b10; box-shadow: 0 0 10px rgba(255,0,255,0.5); }

    /* Print header (hidden on screen) */
    .print-header { display: none; }

    @media (max-width: 600px) {
        body { padding: 10px; }
        .prog-stat-label { width: 140px; font-size: 0.65em; }
        .boss-prog-name { width: 120px; }
        .summary-grid { grid-template-columns: repeat(2, 1fr); }
    }

    @media print {
        body { background: white !important; color: #222 !important; padding: 10px !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        .hud-container { max-width: 100% !important; }
        .back-link, .period-toolbar, .print-btn, #custom-range { display: none !important; }
        .print-header { display: block !important; text-align: center; margin-bottom: 20px; border-bottom: 2px solid #333; padding-bottom: 10px; }
        .print-header h2 { color: #222; font-size: 1.2em; margin: 0 0 4px 0; text-shadow: none; }
        .print-header .print-meta { color: #666; font-size: 0.7em; letter-spacing: 1px; }
        h1 { color: #222 !important; text-shadow: none !important; border-color: #333 !important; font-size: 1.3em; }
        .progress-subtitle { color: #666 !important; }
        .summary-grid { border-color: #ccc !important; background: #f9f9f9 !important; }
        .summary-card { border-color: #ddd !important; background: white !important; break-inside: avoid; }
        .progress-section { border-color: #ccc !important; background: white !important; break-inside: avoid; page-break-inside: avoid; }
        .progress-section h2 { color: #333 !important; border-color: #ccc !important; }
        .sector-prog-ring::after { background: white !important; }
        .timeline-entry { color: #222; }
        .timeline-dot { background: #333 !important; box-shadow: none !important; }
        .timeline::before { background: #ccc !important; }
        .timeline-date { color: #666 !important; }
        .timeline-content { color: #222 !important; }
        .timeline-minion { color: #0066cc !important; }
        .timeline-boss { color: #cc6600 !important; }
        .subj-prog-name { color: #222 !important; }
        .subj-prog-bar { background: #ddd !important; }
        .subj-bar-label { color: #333 !important; }
        .subj-bar-label span { color: #666 !important; }
        .subj-prog-meta { color: #666 !important; }
        .subj-prog-row { border-color: #ddd !important; }
        .timeline-sector { color: #666 !important; }
        .timeline-feedback { color: #666 !important; }
        svg text { fill: #333 !important; }
    }
    </style>
</head>
<body>
    <div class="hud-container">
        <div class="print-header">
            <h2>SOVEREIGN HUD &mdash; PROGRESS REPORT</h2>
            <div class="print-meta">STUDENT: HENRY &bull; PERIOD: ${periodLabel} &bull; PRINTED: ${now.toISOString().slice(0, 10)}</div>
        </div>
        <a class="back-link" href="/">&lt; BACK TO HUD</a>
        <h1>&#x1F4CA; Progress Report</h1>
        <div class="progress-subtitle">STUDENT ACCOMPLISHMENTS &amp; GROWTH TRACKER</div>

        <div class="period-toolbar">
            <a class="period-btn${period === "30d" ? " active" : ""}" href="/progress?period=30d">LAST 30 DAYS</a>
            <a class="period-btn${period === "semester" ? " active" : ""}" href="/progress?period=semester">THIS SEMESTER</a>
            <a class="period-btn${period === "year" ? " active" : ""}" href="/progress?period=year">THIS YEAR</a>
            <a class="period-btn${!period ? " active" : ""}" href="/progress">ALL TIME</a>
            <button class="period-btn${period === "custom" ? " active" : ""}" onclick="document.getElementById('custom-range').style.display=document.getElementById('custom-range').style.display==='flex'?'none':'flex'">CUSTOM...</button>
            <button class="print-btn" onclick="window.print()">&#x1F5A8; PRINT / SAVE PDF</button>
        </div>
        <div id="custom-range" style="display:${period === "custom" ? "flex" : "none"};align-items:center;justify-content:center;gap:10px;margin:-15px 0 20px;flex-wrap:wrap;">
            <label style="font-size:0.7em;color:#888;letter-spacing:1px;">FROM</label>
            <input type="date" id="cr-from" value="${req.query.from || ""}" style="background:#1a1d26;border:1px solid #444;color:#00f2ff;padding:5px 8px;font-family:'Courier New',monospace;font-size:0.8em;">
            <label style="font-size:0.7em;color:#888;letter-spacing:1px;">TO</label>
            <input type="date" id="cr-to" value="${req.query.to || ""}" style="background:#1a1d26;border:1px solid #444;color:#00f2ff;padding:5px 8px;font-family:'Courier New',monospace;font-size:0.8em;">
            <button onclick="var f=document.getElementById('cr-from').value,t=document.getElementById('cr-to').value;if(f)window.location='/progress?period=custom&from='+f+(t?'&to='+t:'')" style="padding:5px 12px;background:#ffea00;color:#0a0b10;border:none;font-family:'Courier New',monospace;font-size:0.75em;font-weight:bold;letter-spacing:1px;cursor:pointer;">GO</button>
        </div>

        <div class="summary-grid">
            <div class="summary-card">
                ${buildSparkBars(enslavedSeries, "#00ff9d", "MINIONS ENSLAVED", "OF " + totalMinions + " TOTAL")}
            </div>
            <div class="summary-card">
                ${buildSparkBars(conquestSeries, "#ffea00", "OVERALL CONQUEST", "% OF ALL MINIONS", "%")}
            </div>
            <div class="summary-card">
                ${buildSparkBars(bossSeries, "#ff6600", "BOSSES CONQUERED", "OF " + totalBosses + " BOSSES")}
            </div>
            <div class="summary-card">
                ${buildSparkBars(guardianSeries, "#ff4444", "GUARDIANS DEFEATED", "OF " + survivalBosses.length + " GUARDIANS")}
            </div>
            <div class="summary-card">
                ${buildSparkBars(questCompleteSeries, "#ff00ff", "QUESTS COMPLETED", "APPROVED QUESTS")}
            </div>
            <div class="summary-card">
                ${buildSparkBars(questInProgressSeries, "#00f2ff", "QUESTS ACTIVE", "IN PROGRESS NOW")}
            </div>
        </div>

        <div class="progress-section">
            <h2>&#x1F4DA; SUBJECT PROGRESS</h2>
            <div style="text-align:center;color:#888;font-size:0.65em;margin-bottom:12px;letter-spacing:1px;">ACADEMIC SUBJECTS &mdash; ${hasPeriod ? `<span style="color:#555;">GRAY = START OF PERIOD</span> &bull; <span style="color:#00f2ff;">COLOR = CURRENT</span>` : "SCHOOL-ALIGNED VIEW"}</div>
            ${subjectTableHtml || '<div style="text-align:center;color:#555;padding:10px;">NO SUBJECTS ASSIGNED YET</div>'}
        </div>

        <div class="progress-section">
            <h2>&#x2694; STAT GROWTH</h2>
            ${statBarsHtml}
        </div>

        <div class="progress-section">
            <h2>&#x1F30D; SECTOR CONQUEST</h2>
            <div class="sector-prog-grid">
                ${sectorChartHtml}
            </div>
        </div>

        <div class="progress-section">
            <h2>&#x1F6E1; SURVIVAL MODE PROGRESS</h2>
            ${survivalHtml || '<div style="text-align:center;color:#555;padding:10px;">NO SURVIVAL BOSSES FOUND</div>'}
        </div>

        <div class="progress-section">
            <h2>&#x1F451; BOSS CONQUEST</h2>
            ${bossConquestHtml}
        </div>

        <div class="progress-section">
            <h2>&#x1F4C5; ACCOMPLISHMENT TIMELINE</h2>
            <div class="timeline">
                ${timelineHtml}
            </div>
        </div>
    </div>
</body>
</html>`);
  } catch (err) {
    console.error("Progress page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// Admin: Landing page
// ---------------------------------------------------------------------------
function buildAdminPage(pendingCount) {
  pendingCount = pendingCount || 0;
  const questTitle = pendingCount > 0 ? `QUEST APPROVAL <span style="color:#ff4444;">(${pendingCount})</span>` : "QUEST APPROVAL";
  const functions = [
    { id: "quests", title: questTitle, desc: pendingCount > 0 ? `${pendingCount} quest${pendingCount > 1 ? "s" : ""} awaiting approval.` : "No quests pending. Check back later.", href: "/admin/quests", active: true },
    { id: "manual", title: "MANUAL ENTRY", desc: "Add new objectives (minions) directly without opening Google Sheets.", href: "/admin/manual", active: true },
    { id: "locks", title: "LOCK/UNLOCK", desc: "Manage prerequisites and locked objectives.", href: "/admin/locks", active: true },
    { id: "import", title: "PHOTO IMPORT", desc: "Upload lesson photos for AI classification and auto-import to the tracker.", href: "/admin/import", active: true },
    { id: "notes", title: "TEACHER NOTES", desc: "Leave notes, observations, and communication for other teachers.", href: "/admin/notes", active: true },
  ];

  const cards = functions.map((f) => {
    if (f.active) {
      return `<a href="${f.href}" class="admin-card active">
        <div class="admin-card-status">ACTIVE</div>
        <h2>${f.title}</h2>
        <p>${f.desc}</p>
        <div class="admin-card-action">LAUNCH &gt;&gt;</div>
      </a>`;
    }
    return `<div class="admin-card locked">
      <div class="admin-card-status">LOCKED</div>
      <h2>${f.title}</h2>
      <p>${f.desc}</p>
      <div class="admin-card-action">COMING SOON</div>
    </div>`;
  }).join("\n");

  return `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Parent Admin - Sovereign HUD</title>
    <style>
    body {
        background: #0a0b10;
        color: #ffea00;
        font-family: 'Courier New', monospace;
        padding: 20px;
        text-transform: uppercase;
        overflow-x: hidden;
    }
    .hud-container {
        border: 2px solid #ffea00;
        padding: 20px;
        box-shadow: 0 0 15px rgba(255, 234, 0, 0.5);
        max-width: 960px;
        margin: auto;
    }
    h1 {
        text-shadow: 2px 2px #ff00ff;
        border-bottom: 1px solid #ffea00;
        letter-spacing: 2px;
        text-align: center;
        margin-bottom: 5px;
        color: #ffea00;
    }
    .admin-subtitle {
        text-align: center;
        color: #888;
        font-size: 0.8em;
        margin-bottom: 30px;
        letter-spacing: 3px;
    }
    .back-link {
        display: inline-block;
        color: #00f2ff;
        text-decoration: none;
        border: 1px solid #00f2ff;
        padding: 6px 15px;
        margin-bottom: 15px;
        font-size: 0.8em;
        transition: all 0.2s;
    }
    .back-link:hover { background: #00f2ff; color: #0a0b10; }
    .admin-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 20px;
    }
    .admin-card {
        border: 2px solid #333;
        border-radius: 6px;
        padding: 20px;
        text-decoration: none;
        color: inherit;
        transition: all 0.3s;
        display: block;
    }
    .admin-card.active {
        border-color: #ffea00;
        box-shadow: 0 0 10px rgba(255, 234, 0, 0.2);
    }
    .admin-card.active:hover {
        background: rgba(255, 234, 0, 0.05);
        box-shadow: 0 0 20px rgba(255, 234, 0, 0.4);
    }
    .admin-card.locked {
        opacity: 0.4;
    }
    .admin-card h2 {
        margin: 8px 0;
        font-size: 1em;
        letter-spacing: 2px;
    }
    .admin-card.active h2 { color: #ffea00; }
    .admin-card.locked h2 { color: #555; }
    .admin-card p {
        font-size: 0.75em;
        color: #888;
        text-transform: none;
        margin: 8px 0 15px;
        line-height: 1.4;
    }
    .admin-card-status {
        font-size: 0.65em;
        letter-spacing: 3px;
        color: #00ff9d;
    }
    .admin-card.locked .admin-card-status { color: #444; }
    .admin-card-action {
        font-size: 0.8em;
        letter-spacing: 2px;
        color: #ffea00;
        font-weight: bold;
    }
    .admin-card.locked .admin-card-action { color: #333; }
    @media (max-width: 600px) {
        .admin-grid { grid-template-columns: 1fr; }
        body { padding: 10px; }
    }
    </style>
</head>
<body>
    <div class="hud-container">
        <a class="back-link" href="/">&lt; BACK TO HUD</a>
        <h1>&#x2699; Parent Admin Console</h1>
        <div class="admin-subtitle">COMMAND AUTHORITY: ACTIVE</div>
        <div class="admin-grid">
            ${cards}
        </div>
    </div>
</body>
</html>`;
}

app.get("/admin", async (req, res) => {
  try {
    const sheets = await getSheets();
    const quests = await fetchQuestsData(sheets);
    const pendingCount = quests.filter((q) => q["Status"] === "Submitted").length;
    res.send(buildAdminPage(pendingCount));
  } catch (err) {
    // Fallback without count on error
    res.send(buildAdminPage(0));
  }
});

// ---------------------------------------------------------------------------
// Admin: Quest Approval page
// ---------------------------------------------------------------------------
app.get("/admin/quests", async (req, res) => {
  try {
    const sheets = await getSheets();
    const quests = await fetchQuestsData(sheets);

    const statusColors = { Active: "#ff6600", Submitted: "#ffea00", Approved: "#00ff9d", Rejected: "#ff0044" };
    // Only show Submitted quests for approval
    const actionable = quests.filter((q) => q["Status"] === "Submitted");
    // Recently approved (last 5 days) for undo
    const recentCutoff = new Date(Date.now() - 5 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
    const recentlyApproved = quests
      .filter((q) => q["Status"] === "Approved" && (q["Date Resolved"] || "") >= recentCutoff)
      .sort((a, b) => (b["Date Resolved"] || "").localeCompare(a["Date Resolved"] || ""));
    const sorted = [...actionable];

    let cards = "";
    const counts = { Submitted: 0, Active: 0, Rejected: 0, Approved: 0 };
    for (const q of quests) {
      counts[q["Status"]] = (counts[q["Status"]] || 0) + 1;
    }
    for (const q of sorted) {
      const sc = statusColors[q["Status"]] || "#555";
      const proofLink = q["Proof Link"] || "";
      const proofDisplay = proofLink
        ? (proofLink.startsWith("http")
          ? `<a href="${escHtml(proofLink)}" target="_blank" style="color:#00f2ff;">${escHtml(proofLink)}</a>`
          : escHtml(proofLink))
        : '<span style="color:#555;">No proof submitted</span>';

      const actions = `
          <div class="qa-feedback-row">
            <input type="text" id="fb-${escHtml(q["Quest ID"])}" placeholder="OPTIONAL NOTE FOR HENRY..." class="qa-feedback">
          </div>
          <div class="qa-actions">
            <form method="POST" action="/admin/quests/approve" style="display:inline;">
              <input type="hidden" name="questId" value="${escHtml(q["Quest ID"])}">
              <input type="hidden" name="feedback" value="">
              <button type="submit" class="qa-btn qa-approve" onclick="this.form.feedback.value=document.getElementById('fb-${escHtml(q["Quest ID"])}').value;return confirm('Approve this quest? The minion will be marked as Enslaved.')">&#x2713; APPROVE</button>
            </form>
            <form method="POST" action="/admin/quests/reject" style="display:inline;">
              <input type="hidden" name="questId" value="${escHtml(q["Quest ID"])}">
              <input type="hidden" name="feedback" value="">
              <button type="submit" class="qa-btn qa-reject" onclick="this.form.feedback.value=document.getElementById('fb-${escHtml(q["Quest ID"])}').value;return confirm('Reject this quest? Henry will need to re-submit.')">&#x2717; REJECT</button>
            </form>
          </div>`;

      cards += `
        <div class="qa-card" style="border-color: ${sc};">
          <div class="qa-header">
            <input type="checkbox" class="qa-bulk-check" value="${escHtml(q["Quest ID"])}">
            <span class="qa-status" style="color:${sc};">${q["Status"]}</span>
            <span class="qa-id">${q["Quest ID"]}</span>
          </div>
          <div class="qa-target">
            <span class="qa-boss">${escHtml(q["Boss"])}</span>
            <span class="qa-arrow">&gt;</span>
            <span class="qa-minion">${escHtml(q["Minion"])}</span>
          </div>
          <div class="qa-sector">SECTOR: ${escHtml(q["Sector"])}</div>
          <div class="qa-task"><span class="qa-task-label">TASK:</span> ${q["Suggested By AI"] || "No details"}</div>
          <div class="qa-proof">
            <span class="qa-proof-label">PROOF:</span>
            ${q["Proof Type"] ? '<span class="qa-proof-type">[' + escHtml(q["Proof Type"]) + ']</span>' : ''}
            ${proofDisplay}
          </div>
          ${q["Date Completed"] ? '<div class="qa-date">DATE: ' + q["Date Completed"] + '</div>' : ''}
          ${actions}
        </div>`;
    }

    res.send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Quest Approval</title>
    <style>
    body { background: #0a0b10; color: #00f2ff; font-family: 'Courier New', monospace; padding: 20px; text-transform: uppercase; }
    .hud-container { border: 2px solid #00f2ff; padding: 20px; box-shadow: 0 0 15px #00f2ff; max-width: 900px; margin: auto; }
    .back-link { display: inline-block; color: #00f2ff; text-decoration: none; border: 1px solid #00f2ff; padding: 6px 15px; margin-bottom: 15px; font-size: 0.8em; transition: all 0.2s; }
    .back-link:hover { background: #00f2ff; color: #0a0b10; }
    h1 { text-align: center; color: #ffea00; text-shadow: 2px 2px #ff00ff; letter-spacing: 4px; margin: 15px 0 5px; }
    .qa-summary { text-align: center; font-size: 0.75em; color: #888; letter-spacing: 2px; margin-bottom: 25px; }
    .qa-summary span { margin: 0 8px; }
    .qa-card {
        border: 1px solid #333; padding: 15px; margin-bottom: 12px;
        background: rgba(255,255,255,0.02); border-left-width: 3px;
    }
    .qa-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; }
    .qa-status { font-weight: bold; font-size: 0.9em; letter-spacing: 2px; }
    .qa-id { font-size: 0.7em; color: #555; }
    .qa-target { font-size: 1.1em; margin-bottom: 4px; }
    .qa-boss { color: #ff6600; font-weight: bold; }
    .qa-arrow { color: #555; margin: 0 6px; }
    .qa-minion { color: #ccc; }
    .qa-sector { font-size: 0.7em; color: #666; letter-spacing: 1px; margin-bottom: 8px; }
    .qa-task { font-size: 0.8em; color: #ccc; border-left: 2px solid #ff00ff; padding-left: 10px; margin-bottom: 8px; text-transform: none; }
    .qa-task-label { color: #ff00ff; font-weight: bold; }
    .qa-proof { font-size: 0.8em; color: #aaa; margin-bottom: 6px; text-transform: none; }
    .qa-proof-label { color: #ffea00; font-weight: bold; }
    .qa-proof-type { color: #ffea00; margin: 0 4px; }
    .qa-date { font-size: 0.7em; color: #555; margin-bottom: 8px; }
    .qa-actions { margin-top: 10px; display: flex; gap: 10px; }
    .qa-btn {
        padding: 8px 16px; font-family: 'Courier New', monospace; font-weight: bold;
        font-size: 0.85em; cursor: pointer; letter-spacing: 1px; border: 1px solid; transition: all 0.2s;
    }
    .qa-approve { background: rgba(0,255,157,0.1); border-color: #00ff9d; color: #00ff9d; }
    .qa-approve:hover { background: #00ff9d; color: #0a0b10; box-shadow: 0 0 10px rgba(0,255,157,0.5); }
    .qa-reject { background: rgba(255,0,68,0.1); border-color: #ff0044; color: #ff0044; }
    .qa-reject:hover { background: #ff0044; color: #0a0b10; box-shadow: 0 0 10px rgba(255,0,68,0.5); }
    .qa-reopen { background: rgba(0,242,255,0.1); border-color: #00f2ff; color: #00f2ff; }
    .qa-reopen:hover { background: #00f2ff; color: #0a0b10; box-shadow: 0 0 10px rgba(0,242,255,0.5); }
    .qa-done { color: #00ff9d; font-weight: bold; letter-spacing: 2px; font-size: 0.85em; }
    .qa-empty { text-align: center; color: #555; padding: 40px; font-size: 0.9em; letter-spacing: 2px; }
    .qa-recent-section { margin-top: 30px; border-top: 1px solid #333; padding-top: 15px; }
    .qa-recent-title { color: #00ff9d; font-size: 0.8em; letter-spacing: 3px; margin: 0 0 12px 0; text-align: center; }
    .qa-recent-card { opacity: 0.6; }
    .qa-recent-card:hover { opacity: 1; }
    .qa-feedback-row { margin-top: 8px; }
    .qa-feedback {
        width: 100%; box-sizing: border-box; padding: 8px 12px;
        background: rgba(255,234,0,0.05); border: 1px solid #444; color: #ffea00;
        font-family: 'Courier New', monospace; font-size: 0.8em; letter-spacing: 1px;
        text-transform: none;
    }
    .qa-feedback::placeholder { color: #666; text-transform: uppercase; }
    .qa-feedback:focus { outline: none; border-color: #ffea00; box-shadow: 0 0 6px rgba(255,234,0,0.3); }
    .qa-bulk-bar {
        display: flex; align-items: center; gap: 12px; padding: 12px 15px;
        border: 1px solid #00ff9d; background: rgba(0,255,157,0.05);
        margin-bottom: 18px; position: sticky; top: 0; z-index: 10;
    }
    .qa-bulk-select-all {
        display: flex; align-items: center; gap: 6px; color: #00ff9d;
        font-size: 0.75em; letter-spacing: 2px; cursor: pointer; white-space: nowrap;
    }
    .qa-bulk-select-all input { accent-color: #00ff9d; width: 16px; height: 16px; cursor: pointer; }
    .qa-bulk-check { accent-color: #00ff9d; width: 16px; height: 16px; cursor: pointer; margin-right: 8px; }
    .qa-bulk-btn:disabled { opacity: 0.4; cursor: not-allowed; }
    </style>
</head>
<body>
    <div class="hud-container">
        <a class="back-link" href="/admin">&lt; BACK TO ADMIN</a>
        <h1>Quest Approval</h1>
        <div class="qa-summary">
            <span style="color:#ffea00;">${counts.Submitted} PENDING</span>
            <span style="color:#ff6600;">${counts.Active} ACTIVE</span>
            <span style="color:#ff0044;">${counts.Rejected} REJECTED</span>
            <span style="color:#00ff9d;">${counts.Approved} APPROVED</span>
        </div>
        <div style="text-align:center;margin-bottom:20px;">
            <form method="POST" action="/admin/quests/sync" style="display:inline;">
                <button type="submit" class="qa-btn qa-reopen" onclick="return confirm('This will re-sync all quest statuses to the Sectors sheet. Use this if statuses look wrong (e.g. after editing the sheet manually). Safe to run anytime. Continue?')">&#x21BB; SYNC TO SECTORS</button>
            </form>
            <div style="color:#666;font-size:0.6em;margin-top:6px;text-transform:none;letter-spacing:0;">Repair tool &mdash; use if Sectors sheet statuses look out of sync with quests. Safe to run anytime.</div>
        </div>
        ${actionable.length > 0 ? `
        <div class="qa-bulk-bar">
            <label class="qa-bulk-select-all"><input type="checkbox" id="qa-select-all"> SELECT ALL</label>
            <input type="text" id="qa-bulk-feedback" placeholder="BULK FEEDBACK (OPTIONAL)..." class="qa-feedback" style="flex:1;max-width:350px;">
            <form method="POST" action="/admin/quests/bulk-approve" id="qa-bulk-form" style="display:inline;">
                <input type="hidden" name="questIds" id="qa-bulk-ids" value="">
                <input type="hidden" name="feedback" id="qa-bulk-fb" value="">
                <button type="submit" class="qa-btn qa-approve qa-bulk-btn" id="qa-bulk-submit" disabled>&#x2713; BULK APPROVE (0)</button>
            </form>
        </div>` : ''}
        ${cards || '<div class="qa-empty">NO QUESTS TO REVIEW</div>'}
        ${recentlyApproved.length > 0 ? `
        <div class="qa-recent-section">
          <h2 class="qa-recent-title">RECENTLY APPROVED</h2>
          ${recentlyApproved.map((q) => `
            <div class="qa-card qa-recent-card" style="border-color: rgba(0,255,157,0.3);">
              <div class="qa-header">
                <span class="qa-status" style="color:#00ff9d;">APPROVED</span>
                <span class="qa-id">${escHtml(q["Quest ID"])} &mdash; ${q["Date Resolved"] || ""}</span>
              </div>
              <div class="qa-target">
                <span class="qa-boss">${escHtml(q["Boss"])}</span>
                <span class="qa-arrow">&gt;</span>
                <span class="qa-minion">${escHtml(q["Minion"])}</span>
              </div>
              <div class="qa-actions">
                <form method="POST" action="/admin/quests/unapprove" style="display:inline;">
                  <input type="hidden" name="questId" value="${escHtml(q["Quest ID"])}">
                  <input type="hidden" name="feedback" value="">
                  <button type="submit" class="qa-btn qa-reject" onclick="return confirm('Undo approval? This will un-enslave the minion and return it here for review.')">&#x21BB; UNDO APPROVAL</button>
                </form>
              </div>
            </div>
          `).join("")}
        </div>` : ""}
    </div>
    <script>
    (function() {
        var selectAll = document.getElementById('qa-select-all');
        var bulkBtn = document.getElementById('qa-bulk-submit');
        var bulkForm = document.getElementById('qa-bulk-form');
        var bulkIds = document.getElementById('qa-bulk-ids');
        var bulkFb = document.getElementById('qa-bulk-fb');
        var bulkFeedback = document.getElementById('qa-bulk-feedback');
        var checks = document.querySelectorAll('.qa-bulk-check');
        if (!selectAll || !checks.length) return;
        function updateBulk() {
            var checked = document.querySelectorAll('.qa-bulk-check:checked');
            var count = checked.length;
            bulkBtn.textContent = '\\u2713 BULK APPROVE (' + count + ')';
            bulkBtn.disabled = count === 0;
            selectAll.checked = count === checks.length && count > 0;
        }
        selectAll.addEventListener('change', function() {
            checks.forEach(function(c) { c.checked = selectAll.checked; });
            updateBulk();
        });
        checks.forEach(function(c) { c.addEventListener('change', updateBulk); });
        bulkForm.addEventListener('submit', function(e) {
            var checked = document.querySelectorAll('.qa-bulk-check:checked');
            if (checked.length === 0) { e.preventDefault(); return; }
            var ids = [];
            checked.forEach(function(c) { ids.push(c.value); });
            bulkIds.value = ids.join(',');
            bulkFb.value = bulkFeedback ? bulkFeedback.value : '';
            if (!confirm('Approve ' + ids.length + ' quest(s)? All selected minions will be marked as Enslaved.')) {
                e.preventDefault();
            }
        });
    })();
    </script>
</body>
</html>`);
  } catch (err) {
    console.error("Quest approval page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// Admin: Quest Approval actions (approve, reject, reopen)
// ---------------------------------------------------------------------------
async function findQuestAndUpdate(req, res, newStatus, clearDate) {
  const { questId } = req.body;
  if (!questId) return res.status(400).send("Missing questId");

  const sheets = await getSheets();
  const questsRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Quests",
  });
  const rows = questsRes.data.values;
  if (!rows || rows.length < 2) return res.status(404).send("No quests found");

  const headers = rows[0];
  const idCol = headers.indexOf("Quest ID");
  const statusCol = headers.indexOf("Status");
  const bossCol = headers.indexOf("Boss");
  const minionCol = headers.indexOf("Minion");
  const sectorCol = headers.indexOf("Sector");
  const dateCol = headers.indexOf("Date Completed");
  const dateAddedCol = headers.indexOf("Date Added");
  const dateResolvedCol = headers.indexOf("Date Resolved");
  const feedbackCol = headers.indexOf("Feedback");
  const colRef = (col) => String.fromCharCode(65 + col);

  let targetRowIdx = -1;
  let questRow = null;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idCol] === questId) {
      targetRowIdx = i + 1; // 1-based
      questRow = rows[i];
      break;
    }
  }
  if (targetRowIdx === -1) return res.status(404).send("Quest not found");

  const now = new Date().toISOString().slice(0, 10);
  const updates = [{
    range: "Quests!" + colRef(statusCol) + targetRowIdx,
    values: [[newStatus]],
  }];
  if (clearDate && dateCol >= 0) {
    updates.push({
      range: "Quests!" + colRef(dateCol) + targetRowIdx,
      values: [[""]],
    });
  }

  // Approve/Reject: set Date Resolved (quest is no longer active)
  if ((newStatus === "Approved" || newStatus === "Rejected") && dateResolvedCol >= 0) {
    updates.push({
      range: "Quests!" + colRef(dateResolvedCol) + targetRowIdx,
      values: [[now]],
    });
  }

  // Reopen or un-approve back to Submitted: clear Date Resolved
  if ((clearDate || newStatus === "Submitted") && dateResolvedCol >= 0) {
    updates.push({
      range: "Quests!" + colRef(dateResolvedCol) + targetRowIdx,
      values: [[""]],
    });
  }

  // Write feedback (from approve/reject form, or clear on reopen)
  if (feedbackCol >= 0) {
    const feedback = clearDate ? "" : (req.body.feedback || "");
    updates.push({
      range: "Quests!" + colRef(feedbackCol) + targetRowIdx,
      values: [[feedback]],
    });
  }

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { valueInputOption: "RAW", data: updates },
  });

  // Sync Quest Status to Sectors
  const qBoss = questRow[bossCol];
  const qMinion = questRow[minionCol];
  const qSector = questRow[sectorCol];
  if (qBoss && qMinion && qSector) {
    await updateSectorsQuestStatus(sheets, qSector, qBoss, qMinion, newStatus);
  }

  res.redirect("/admin/quests");
}

app.post("/admin/quests/approve", async (req, res) => {
  try {
    await findQuestAndUpdate(req, res, "Approved", false);
  } catch (err) {
    console.error("Quest approve error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

app.post("/admin/quests/reject", async (req, res) => {
  try {
    await findQuestAndUpdate(req, res, "Rejected", false);
  } catch (err) {
    console.error("Quest reject error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

app.post("/admin/quests/reopen", async (req, res) => {
  try {
    await findQuestAndUpdate(req, res, "Active", true);
  } catch (err) {
    console.error("Quest reopen error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

app.post("/admin/quests/unapprove", async (req, res) => {
  try {
    await findQuestAndUpdate(req, res, "Submitted", false);
  } catch (err) {
    console.error("Quest unapprove error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

app.post("/admin/quests/bulk-approve", async (req, res) => {
  try {
    const { questIds, feedback } = req.body;
    if (!questIds) return res.status(400).send("Missing questIds");
    const ids = questIds.split(",").map((s) => s.trim()).filter(Boolean);
    if (ids.length === 0) return res.status(400).send("No quest IDs provided");

    const sheets = await getSheets();
    const questsRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Quests",
    });
    const rows = questsRes.data.values;
    if (!rows || rows.length < 2) return res.status(404).send("No quests found");

    const headers = rows[0];
    const idCol = headers.indexOf("Quest ID");
    const statusCol = headers.indexOf("Status");
    const bossCol = headers.indexOf("Boss");
    const minionCol = headers.indexOf("Minion");
    const sectorCol = headers.indexOf("Sector");
    const dateResolvedCol = headers.indexOf("Date Resolved");
    const feedbackCol = headers.indexOf("Feedback");
    const colRef = (col) => String.fromCharCode(65 + col);
    const now = new Date().toISOString().slice(0, 10);

    const updates = [];
    const sectorsToSync = [];

    for (const qId of ids) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][idCol] !== qId) continue;
        const rowNum = i + 1;
        updates.push({ range: "Quests!" + colRef(statusCol) + rowNum, values: [["Approved"]] });
        if (dateResolvedCol >= 0) {
          updates.push({ range: "Quests!" + colRef(dateResolvedCol) + rowNum, values: [[now]] });
        }
        if (feedbackCol >= 0) {
          updates.push({ range: "Quests!" + colRef(feedbackCol) + rowNum, values: [[feedback || ""]] });
        }
        sectorsToSync.push({
          sector: rows[i][sectorCol],
          boss: rows[i][bossCol],
          minion: rows[i][minionCol],
        });
        break;
      }
    }

    if (updates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { valueInputOption: "RAW", data: updates },
      });
    }

    for (const s of sectorsToSync) {
      if (s.boss && s.minion && s.sector) {
        await updateSectorsQuestStatus(sheets, s.sector, s.boss, s.minion, "Approved");
      }
    }

    res.redirect("/admin/quests");
  } catch (err) {
    console.error("Bulk approve error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

app.post("/admin/quests/sync", async (req, res) => {
  try {
    const sheets = await getSheets();
    const quests = await fetchQuestsData(sheets);

    let synced = 0;
    for (const q of quests) {
      const boss = q["Boss"];
      const minion = q["Minion"];
      const sector = q["Sector"];
      const status = q["Status"];
      if (boss && minion && sector && status) {
        await updateSectorsQuestStatus(sheets, sector, boss, minion, status);
        synced++;
      }
    }

    res.redirect("/admin/quests");
  } catch (err) {
    console.error("Quest sync error:", err);
    res.status(500).send(`<pre style="color:red">Error syncing quests: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// Admin: Lock/Unlock Management — manage prerequisites for locked minions
// ---------------------------------------------------------------------------
app.get("/admin/locks", async (req, res) => {
  try {
    const sheets = await getSheets();
    const secRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Sectors",
    });
    const rows = secRes.data.values;
    if (!rows || rows.length < 2) return res.send("No data in Sectors sheet");

    const headers = rows[0];
    const sectorCol = headers.indexOf("Sector");
    const bossCol = headers.indexOf("Boss");
    const minionCol = headers.indexOf("Minion");
    const statusCol = headers.indexOf("Status");
    const lockedCol = headers.indexOf("Locked for what?");

    const subjectCol = headers.indexOf("Subject");
    const allMinions = [];
    const prereqData = {}; // { sector: { boss: { subject, minions: [] } } }
    const subjectToSectors = {}; // { subject: Set(sectors) }
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const sector = (r[sectorCol] || "").trim();
      const boss = (r[bossCol] || "").trim();
      const minion = (r[minionCol] || "").trim();
      const status = (r[statusCol] || "").trim();
      const prereq = lockedCol >= 0 ? (r[lockedCol] || "").trim() : "";
      const subject = subjectCol >= 0 ? (r[subjectCol] || "").trim() : "";
      if (!sector || !boss || !minion) continue;
      allMinions.push({ sector, boss, minion, status, prereq });
      if (!prereqData[sector]) prereqData[sector] = {};
      if (!prereqData[sector][boss]) prereqData[sector][boss] = { subject: "", minions: [] };
      if (subject && !prereqData[sector][boss].subject) prereqData[sector][boss].subject = subject;
      if (!prereqData[sector][boss].minions.includes(minion)) prereqData[sector][boss].minions.push(minion);
      if (subject) {
        if (!subjectToSectors[subject]) subjectToSectors[subject] = new Set();
        subjectToSectors[subject].add(sector);
      }
    }
    // Convert sets to arrays for JSON
    const subjectMap = {};
    for (const [subj, secs] of Object.entries(subjectToSectors)) {
      subjectMap[subj] = [...secs].sort();
    }

    // Sort: Locked first, then Engaged, then Enslaved
    const statusOrder = { Locked: 0, Engaged: 1, Enslaved: 2 };
    allMinions.sort((a, b) => (statusOrder[a.status] ?? 3) - (statusOrder[b.status] ?? 3) || a.sector.localeCompare(b.sector) || a.boss.localeCompare(b.boss));

    const statusColors = { Locked: "#555", Engaged: "#ff6600", Enslaved: "#00ff9d" };

    // Track engaged counts per boss for "Lock All" buttons
    const engagedPerBoss = {};
    for (const m of allMinions) {
      if (m.status === "Engaged") {
        const key = m.sector + "|" + m.boss;
        if (!engagedPerBoss[key]) engagedPerBoss[key] = { sector: m.sector, boss: m.boss, count: 0 };
        engagedPerBoss[key].count++;
      }
    }
    const shownLockAll = new Set();

    let tableRows = "";
    for (const m of allMinions) {
      const sc = statusColors[m.status] || "#555";
      const isLocked = m.status === "Locked";
      const isEngaged = m.status === "Engaged";

      let actions = "";
      if (isLocked) {
        actions = `
          <button type="button" class="lock-btn lock-edit" onclick="editPrereq(this)" data-sector="${escHtml(m.sector)}" data-boss="${escHtml(m.boss)}" data-minion="${escHtml(m.minion)}" data-prereq="${escHtml(m.prereq)}">EDIT</button>
          <form method="POST" action="/admin/locks/update" style="display:inline;">
            <input type="hidden" name="sector" value="${escHtml(m.sector)}">
            <input type="hidden" name="boss" value="${escHtml(m.boss)}">
            <input type="hidden" name="minion" value="${escHtml(m.minion)}">
            <input type="hidden" name="status" value="Engaged">
            <input type="hidden" name="prerequisite" value="">
            <button type="submit" class="lock-btn lock-unlock" onclick="return confirm('Unlock this minion? It will become Engaged.')">UNLOCK</button>
          </form>`;
      } else if (isEngaged) {
        const bossKey = m.sector + "|" + m.boss;
        const lockAllBtn = (!shownLockAll.has(bossKey) && engagedPerBoss[bossKey] && engagedPerBoss[bossKey].count > 1)
          ? `<button type="button" class="lock-btn lock-lock" onclick="lockBoss(this)" data-sector="${escHtml(m.sector)}" data-boss="${escHtml(m.boss)}" style="font-size:0.6em;">LOCK ALL ${engagedPerBoss[bossKey].count}</button>`
          : "";
        if (lockAllBtn) shownLockAll.add(bossKey);
        actions = `
          <button type="button" class="lock-btn lock-lock" onclick="lockMinion(this)" data-sector="${escHtml(m.sector)}" data-boss="${escHtml(m.boss)}" data-minion="${escHtml(m.minion)}">LOCK</button>
          ${lockAllBtn}`;
      } else {
        actions = `<span style="color:#00ff9d;font-size:0.7em;">ENSLAVED</span>`;
      }

      tableRows += `
        <tr>
          <td style="color:#888;font-size:0.75em;">${escHtml(m.sector)}</td>
          <td style="color:#ff6600;">${escHtml(m.boss)}</td>
          <td>${escHtml(m.minion)}</td>
          <td style="color:${sc};font-weight:bold;">${m.status}</td>
          <td class="prereq-cell" style="font-size:0.75em;color:#ffea00;">${escHtml(m.prereq)}</td>
          <td class="lock-actions">${actions}</td>
        </tr>`;
    }

    res.send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Lock/Unlock Management</title>
    <style>
    body { background: #0a0b10; color: #00f2ff; font-family: 'Courier New', monospace; padding: 20px; text-transform: uppercase; }
    .hud-container { border: 2px solid #00f2ff; padding: 20px; box-shadow: 0 0 15px #00f2ff; max-width: 1100px; margin: auto; }
    .back-link { display: inline-block; color: #00f2ff; text-decoration: none; border: 1px solid #00f2ff; padding: 6px 15px; margin-bottom: 15px; font-size: 0.8em; transition: all 0.2s; }
    .back-link:hover { background: #00f2ff; color: #0a0b10; }
    h1 { text-align: center; color: #ffea00; text-shadow: 2px 2px #ff00ff; letter-spacing: 4px; margin: 15px 0 5px; }
    .lock-subtitle { text-align: center; font-size: 0.7em; color: #888; letter-spacing: 2px; margin-bottom: 20px; }
    table { width: 100%; border-collapse: collapse; font-size: 0.8em; }
    th { background: #1a1d26; color: #ffea00; padding: 10px 8px; text-align: left; border-bottom: 2px solid #00f2ff; }
    td { padding: 8px; border-bottom: 1px solid #1a1d26; }
    tr:hover td { background: rgba(0, 242, 255, 0.05); }
    .lock-btn {
        padding: 4px 10px; font-family: 'Courier New', monospace; font-weight: bold;
        font-size: 0.75em; cursor: pointer; letter-spacing: 1px; border: 1px solid; transition: all 0.2s;
        background: none;
    }
    .lock-edit { border-color: #ffea00; color: #ffea00; }
    .lock-edit:hover { background: #ffea00; color: #0a0b10; }
    .lock-unlock { border-color: #00ff9d; color: #00ff9d; }
    .lock-unlock:hover { background: #00ff9d; color: #0a0b10; }
    .lock-lock { border-color: #ff0044; color: #ff0044; }
    .lock-lock:hover { background: #ff0044; color: #0a0b10; }
    .lock-actions { white-space: nowrap; }
    .lock-modal-overlay {
        display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0;
        background: rgba(0,0,0,0.8); z-index: 100; align-items: center; justify-content: center;
    }
    .lock-modal-overlay.active { display: flex; }
    .lock-modal {
        background: #0a0b10; border: 2px solid #ffea00; padding: 25px; max-width: 500px; width: 90%;
        box-shadow: 0 0 20px rgba(255,234,0,0.3);
    }
    .lock-modal h3 { color: #ffea00; margin: 0 0 15px 0; letter-spacing: 2px; }
    .lock-modal-actions { margin-top: 15px; display: flex; gap: 10px; }
    .lock-modal .lock-btn { padding: 8px 16px; }
    .prereq-picker { display: flex; gap: 8px; flex-wrap: wrap; align-items: flex-end; margin-bottom: 10px; }
    .prereq-picker-group { flex: 1; min-width: 120px; }
    .prereq-picker-group label { display: block; font-size: 0.6em; color: #888; letter-spacing: 1px; margin-bottom: 3px; }
    .prereq-picker-group select {
        width: 100%; padding: 8px; background: #1a1d26; border: 1px solid #444; color: #00f2ff;
        font-family: 'Courier New', monospace; font-size: 0.8em;
    }
    .prereq-picker-group select:focus { outline: none; border-color: #ffea00; }
    .prereq-picker-group select:disabled { opacity: 0.4; }
    .prereq-add-btn {
        padding: 8px 12px; background: rgba(0,255,157,0.1); border: 1px solid #00ff9d; color: #00ff9d;
        font-family: 'Courier New', monospace; font-size: 0.75em; font-weight: bold;
        cursor: pointer; letter-spacing: 1px; transition: all 0.2s; white-space: nowrap;
    }
    .prereq-add-btn:hover { background: #00ff9d; color: #0a0b10; }
    .prereq-add-btn:disabled { opacity: 0.4; cursor: not-allowed; }
    .prereq-chips { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 10px; min-height: 24px; }
    .prereq-chip {
        display: inline-flex; align-items: center; gap: 6px; padding: 4px 10px;
        background: rgba(255,234,0,0.1); border: 1px solid #ffea00; color: #ffea00;
        font-size: 0.7em; font-family: 'Courier New', monospace; letter-spacing: 1px;
    }
    .prereq-chip-remove {
        cursor: pointer; color: #ff0044; font-weight: bold; font-size: 1.2em; line-height: 1;
    }
    .prereq-chip-remove:hover { color: #ff4444; }
    .prereq-empty { font-size: 0.65em; color: #555; letter-spacing: 1px; }
    .lm-target-box {
        margin-bottom: 12px; padding: 10px; border: 1px solid #00f2ff;
        background: rgba(0,242,255,0.05);
    }
    .lm-target-label { font-size: 0.6em; color: #888; letter-spacing: 2px; margin-bottom: 2px; }
    .lm-target-sector { font-size: 0.7em; color: #888; letter-spacing: 1px; }
    .lm-target-main { font-size: 1em; font-weight: bold; letter-spacing: 1px; }
    .lock-find-toggle, .lock-find-toggle-b {
        border-color: #555; color: #555; background: none; cursor: pointer; transition: all 0.2s;
    }
    .lock-find-toggle.active, .lock-find-toggle-b.active { border-color: #ffea00; color: #ffea00; background: rgba(255,234,0,0.1); }
    </style>
</head>
<body>
    <div class="hud-container">
        <a class="back-link" href="/admin">&lt; BACK TO ADMIN</a>
        <h1>Lock/Unlock Management</h1>
        <div class="lock-subtitle">${allMinions.filter((m) => m.status === "Locked").length} LOCKED &bull; ${allMinions.filter((m) => m.status === "Engaged").length} ENGAGED &bull; ${allMinions.filter((m) => m.status === "Enslaved").length} ENSLAVED</div>
        <table>
            <thead>
                <tr><th>Sector</th><th>Boss</th><th>Minion</th><th>Status</th><th>Prerequisite</th><th>Actions</th></tr>
            </thead>
            <tbody>${tableRows}</tbody>
        </table>
    </div>

    <div class="lock-modal-overlay" id="lock-modal">
        <div class="lock-modal">
            <h3 id="lock-modal-title">EDIT PREREQUISITE</h3>
            <form method="POST" action="/admin/locks/update" id="lock-modal-form">
                <input type="hidden" name="sector" id="lm-sector">
                <input type="hidden" name="boss" id="lm-boss">
                <input type="hidden" name="minion" id="lm-minion">
                <input type="hidden" name="status" id="lm-status">
                <input type="hidden" name="prerequisite" id="lm-prereq">
                <div class="lm-target-box" id="lm-target-box">
                    <div class="lm-target-label">TARGET:</div>
                    <div class="lm-target-sector" id="lm-target-sector"></div>
                    <div class="lm-target-main" id="lm-target-main"></div>
                </div>
                <div style="font-size:0.65em;color:#ffea00;letter-spacing:1px;margin-bottom:6px;">PREREQUISITES:</div>
                <div class="prereq-chips" id="lm-chips"><span class="prereq-empty">NONE — ADD BELOW</span></div>
                <div style="margin-bottom:6px;">
                    <span style="font-size:0.6em;color:#888;letter-spacing:1px;">FIND BY: </span>
                    <button type="button" class="lock-btn lock-find-toggle active" id="lm-find-subject" style="font-size:0.6em;padding:3px 8px;">SUBJECT</button>
                    <button type="button" class="lock-btn lock-find-toggle" id="lm-find-sector" style="font-size:0.6em;padding:3px 8px;">SECTOR</button>
                </div>
                <div class="prereq-picker">
                    <div class="prereq-picker-group" id="lm-group-subject">
                        <label>SUBJECT</label>
                        <select id="lm-step-subject"><option value="">Select subject...</option></select>
                    </div>
                    <div class="prereq-picker-group" id="lm-group-sector">
                        <label>SECTOR</label>
                        <select id="lm-step-sector"><option value="">Select...</option></select>
                    </div>
                    <div class="prereq-picker-group">
                        <label>BOSS</label>
                        <select id="lm-step-boss" disabled><option value="">Select above first...</option></select>
                    </div>
                    <div class="prereq-picker-group">
                        <label>REQUIRE</label>
                        <select id="lm-step-level" disabled><option value="">Select boss first...</option></select>
                    </div>
                    <button type="button" class="prereq-add-btn" id="lm-add-btn" disabled>+ ADD</button>
                </div>
                <div class="lock-modal-actions">
                    <button type="submit" class="lock-btn lock-unlock">SAVE</button>
                    <button type="button" class="lock-btn lock-edit" onclick="closeModal()">CANCEL</button>
                </div>
            </form>
        </div>
    </div>

    <!-- Bulk lock modal -->
    <div class="lock-modal-overlay" id="lock-boss-modal">
        <div class="lock-modal">
            <h3>LOCK ALL MINIONS OF BOSS</h3>
            <form method="POST" action="/admin/locks/bulk-lock" id="lock-boss-form">
                <input type="hidden" name="sector" id="lb-sector">
                <input type="hidden" name="boss" id="lb-boss">
                <input type="hidden" name="prerequisite" id="lb-prereq">
                <div class="lm-target-box">
                    <div class="lm-target-label">BOSS:</div>
                    <div class="lm-target-main" id="lb-target"></div>
                </div>
                <div style="font-size:0.65em;color:#ffea00;letter-spacing:1px;margin-bottom:6px;">PREREQUISITE FOR ALL:</div>
                <div class="prereq-chips" id="lb-chips"><span class="prereq-empty">NONE — ADD BELOW</span></div>
                <div style="margin-bottom:6px;">
                    <span style="font-size:0.6em;color:#888;letter-spacing:1px;">FIND BY: </span>
                    <button type="button" class="lock-btn lock-find-toggle-b active" id="lb-find-subject" style="font-size:0.6em;padding:3px 8px;">SUBJECT</button>
                    <button type="button" class="lock-btn lock-find-toggle-b" id="lb-find-sector" style="font-size:0.6em;padding:3px 8px;">SECTOR</button>
                </div>
                <div class="prereq-picker">
                    <div class="prereq-picker-group" id="lb-group-subject">
                        <label>SUBJECT</label>
                        <select id="lb-step-subject"><option value="">Select subject...</option></select>
                    </div>
                    <div class="prereq-picker-group" id="lb-group-sector">
                        <label>SECTOR</label>
                        <select id="lb-step-sector"><option value="">Select...</option></select>
                    </div>
                    <div class="prereq-picker-group">
                        <label>BOSS</label>
                        <select id="lb-step-boss" disabled><option value="">Select above first...</option></select>
                    </div>
                    <div class="prereq-picker-group">
                        <label>REQUIRE</label>
                        <select id="lb-step-level" disabled><option value="">Select boss first...</option></select>
                    </div>
                    <button type="button" class="prereq-add-btn" id="lb-add-btn" disabled>+ ADD</button>
                </div>
                <div class="lock-modal-actions">
                    <button type="submit" class="lock-btn lock-unlock">LOCK ALL</button>
                    <button type="button" class="lock-btn lock-edit" onclick="closeBossModal()">CANCEL</button>
                </div>
            </form>
        </div>
    </div>

    <script>
    var PD = ${JSON.stringify(prereqData)};
    var SM = ${JSON.stringify(subjectMap)};

    // --- Shared prereq picker logic ---
    function PrereqPicker(opts) {
        var self = this;
        self.stepSubject = document.getElementById(opts.subjectId);
        self.stepSector = document.getElementById(opts.sectorId);
        self.stepBoss = document.getElementById(opts.bossId);
        self.stepLevel = document.getElementById(opts.levelId);
        self.addBtn = document.getElementById(opts.addBtnId);
        self.chipsEl = document.getElementById(opts.chipsId);
        self.hiddenInput = document.getElementById(opts.hiddenId);
        self.groupSubject = document.getElementById(opts.groupSubjectId);
        self.groupSector = document.getElementById(opts.groupSectorId);
        self.findSubjectBtn = document.getElementById(opts.findSubjectBtnId);
        self.findSectorBtn = document.getElementById(opts.findSectorBtnId);
        self.chips = [];
        self.mode = 'subject'; // or 'sector'

        // Populate subjects
        var subjects = Object.keys(SM).sort();
        self.stepSubject.innerHTML = '<option value="">Select subject...</option>' + subjects.map(function(s) {
            return '<option value="' + s + '">' + s + '</option>';
        }).join('');

        // Populate sectors
        var sectors = Object.keys(PD).sort();
        self.stepSector.innerHTML = '<option value="">Select sector...</option>' + sectors.map(function(s) {
            return '<option value="' + s + '">' + s + '</option>';
        }).join('');

        // Mode toggle
        self.setMode = function(mode) {
            self.mode = mode;
            if (mode === 'subject') {
                self.groupSubject.style.display = '';
                self.groupSector.style.display = 'none';
                self.findSubjectBtn.classList.add('active');
                self.findSectorBtn.classList.remove('active');
            } else {
                self.groupSubject.style.display = 'none';
                self.groupSector.style.display = '';
                self.findSectorBtn.classList.add('active');
                self.findSubjectBtn.classList.remove('active');
            }
            self.resetPicker();
        };
        self.findSubjectBtn.addEventListener('click', function() { self.setMode('subject'); });
        self.findSectorBtn.addEventListener('click', function() { self.setMode('sector'); });

        // Subject change → populate sector(s) → then boss
        self.stepSubject.addEventListener('change', function() {
            var subj = this.value;
            self.stepLevel.innerHTML = '<option value="">Select boss first...</option>';
            self.stepLevel.disabled = true;
            self.addBtn.disabled = true;
            if (!subj || !SM[subj]) {
                self.stepBoss.innerHTML = '<option value="">Select subject first...</option>';
                self.stepBoss.disabled = true;
                return;
            }
            // Find all bosses across matching sectors for this subject
            var secs = SM[subj];
            var bossOpts = '<option value="">Select boss...</option>';
            for (var si = 0; si < secs.length; si++) {
                var sec = secs[si];
                if (!PD[sec]) continue;
                var bosses = Object.keys(PD[sec]).sort();
                for (var bi = 0; bi < bosses.length; bi++) {
                    var b = bosses[bi];
                    if (PD[sec][b].subject === subj) {
                        bossOpts += '<option value="' + sec + '|' + b + '">' + b + ' (' + sec + ')</option>';
                    }
                }
            }
            self.stepBoss.innerHTML = bossOpts;
            self.stepBoss.disabled = false;
        });

        // Sector change → populate boss
        self.stepSector.addEventListener('change', function() {
            var sec = this.value;
            self.stepLevel.innerHTML = '<option value="">Select boss first...</option>';
            self.stepLevel.disabled = true;
            self.addBtn.disabled = true;
            if (!sec || !PD[sec]) {
                self.stepBoss.innerHTML = '<option value="">Select sector first...</option>';
                self.stepBoss.disabled = true;
                return;
            }
            var bosses = Object.keys(PD[sec]).sort();
            self.stepBoss.innerHTML = '<option value="">Select boss...</option>' + bosses.map(function(b) {
                var subj = PD[sec][b].subject;
                return '<option value="' + sec + '|' + b + '">' + b + (subj ? ' (' + subj + ')' : '') + '</option>';
            }).join('');
            self.stepBoss.disabled = false;
        });

        // Boss change → populate level
        self.stepBoss.addEventListener('change', function() {
            var val = this.value;
            self.addBtn.disabled = true;
            if (!val) { self.stepLevel.innerHTML = '<option value="">Select boss first...</option>'; self.stepLevel.disabled = true; return; }
            var parts = val.split('|');
            var sec = parts[0], boss = parts[1];
            if (!PD[sec] || !PD[sec][boss]) { self.stepLevel.disabled = true; return; }
            var minions = PD[sec][boss].minions || [];
            var levelOpts = '<option value="">Choose...</option>';
            levelOpts += '<option value="Boss:' + boss + '">ENTIRE BOSS: ' + boss + ' (all minions)</option>';
            for (var i = 0; i < minions.length; i++) {
                levelOpts += '<option value="Minion:' + boss + '>' + minions[i] + '">MINION: ' + minions[i] + '</option>';
            }
            self.stepLevel.innerHTML = levelOpts;
            self.stepLevel.disabled = false;
        });

        self.stepLevel.addEventListener('change', function() { self.addBtn.disabled = !this.value; });

        self.addBtn.addEventListener('click', function() {
            var val = self.stepLevel.value;
            if (!val || self.chips.indexOf(val) >= 0) return;
            self.chips.push(val);
            self.renderChips();
            self.resetPicker();
        });

        self.renderChips = function() {
            if (self.chips.length === 0) {
                self.chipsEl.innerHTML = '<span class="prereq-empty">NONE \\u2014 ADD BELOW</span>';
            } else {
                self.chipsEl.innerHTML = self.chips.map(function(c, i) {
                    return '<span class="prereq-chip">' + c + ' <span class="prereq-chip-remove" data-idx="' + i + '">\\u00d7</span></span>';
                }).join('');
            }
            self.hiddenInput.value = self.chips.join(';');
        };

        self.chipsEl.addEventListener('click', function(e) {
            var rm = e.target.closest('.prereq-chip-remove');
            if (!rm) return;
            self.chips.splice(parseInt(rm.dataset.idx), 1);
            self.renderChips();
        });

        self.resetPicker = function() {
            self.stepSubject.value = '';
            self.stepSector.value = '';
            self.stepBoss.innerHTML = '<option value="">Select above first...</option>';
            self.stepBoss.disabled = true;
            self.stepLevel.innerHTML = '<option value="">Select boss first...</option>';
            self.stepLevel.disabled = true;
            self.addBtn.disabled = true;
        };

        self.open = function(chips) {
            self.chips = chips;
            self.renderChips();
            self.setMode('subject');
        };
    }

    // --- Single minion modal ---
    var modal = document.getElementById('lock-modal');
    var picker = new PrereqPicker({
        subjectId: 'lm-step-subject', sectorId: 'lm-step-sector', bossId: 'lm-step-boss',
        levelId: 'lm-step-level', addBtnId: 'lm-add-btn', chipsId: 'lm-chips',
        hiddenId: 'lm-prereq', groupSubjectId: 'lm-group-subject', groupSectorId: 'lm-group-sector',
        findSubjectBtnId: 'lm-find-subject', findSectorBtnId: 'lm-find-sector'
    });

    function openModal(btn, title) {
        document.getElementById('lm-sector').value = btn.dataset.sector;
        document.getElementById('lm-boss').value = btn.dataset.boss;
        document.getElementById('lm-minion').value = btn.dataset.minion;
        document.getElementById('lm-status').value = 'Locked';
        document.getElementById('lm-target-sector').textContent = btn.dataset.sector;
        document.getElementById('lm-target-main').innerHTML = '<span style="color:#ff6600;">' + btn.dataset.boss + '</span> <span style="color:#555;">&gt;</span> <span style="color:#00f2ff;">' + btn.dataset.minion + '</span>';
        document.getElementById('lock-modal-title').textContent = title;
        var existing = (btn.dataset.prereq || '').split(';').map(function(s) { return s.trim(); }).filter(Boolean);
        picker.open(existing);
        modal.classList.add('active');
    }
    function editPrereq(btn) { openModal(btn, 'EDIT PREREQUISITE'); }
    function lockMinion(btn) { openModal(btn, 'LOCK MINION'); }
    function closeModal() { modal.classList.remove('active'); }
    modal.addEventListener('click', function(e) { if (e.target === modal) closeModal(); });

    // --- Bulk lock boss modal ---
    var bossModal = document.getElementById('lock-boss-modal');
    var bossPicker = new PrereqPicker({
        subjectId: 'lb-step-subject', sectorId: 'lb-step-sector', bossId: 'lb-step-boss',
        levelId: 'lb-step-level', addBtnId: 'lb-add-btn', chipsId: 'lb-chips',
        hiddenId: 'lb-prereq', groupSubjectId: 'lb-group-subject', groupSectorId: 'lb-group-sector',
        findSubjectBtnId: 'lb-find-subject', findSectorBtnId: 'lb-find-sector'
    });

    function lockBoss(btn) {
        document.getElementById('lb-sector').value = btn.dataset.sector;
        document.getElementById('lb-boss').value = btn.dataset.boss;
        document.getElementById('lb-target').innerHTML = '<span style="color:#ff6600;">' + btn.dataset.boss + '</span> <span style="color:#888;">(' + btn.dataset.sector + ')</span>';
        bossPicker.open([]);
        bossModal.classList.add('active');
    }
    function closeBossModal() { bossModal.classList.remove('active'); }
    bossModal.addEventListener('click', function(e) { if (e.target === bossModal) closeBossModal(); });
    </script>
</body>
</html>`);
  } catch (err) {
    console.error("Lock/unlock page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

app.post("/admin/locks/update", async (req, res) => {
  try {
    const { sector, boss, minion, status, prerequisite } = req.body;
    if (!sector || !boss || !minion || !status) return res.status(400).send("Missing fields");

    const sheets = await getSheets();
    const secRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Sectors",
    });
    const rows = secRes.data.values;
    if (!rows || rows.length < 2) return res.status(404).send("No sectors data");

    const headers = rows[0];
    const sectorCol = headers.indexOf("Sector");
    const bossCol = headers.indexOf("Boss");
    const minionCol = headers.indexOf("Minion");
    const statusCol = headers.indexOf("Status");
    const lockedCol = headers.indexOf("Locked for what?");

    const colLetter = (idx) => {
      if (idx < 26) return String.fromCharCode(65 + idx);
      return String.fromCharCode(64 + Math.floor(idx / 26)) + String.fromCharCode(65 + (idx % 26));
    };

    let found = false;
    const updates = [];
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][sectorCol] || "").trim() === sector &&
          (rows[i][bossCol] || "").trim() === boss &&
          (rows[i][minionCol] || "").trim() === minion) {
        const rowNum = i + 1;
        updates.push({ range: `Sectors!${colLetter(statusCol)}${rowNum}`, values: [[status]] });
        if (lockedCol >= 0) {
          updates.push({ range: `Sectors!${colLetter(lockedCol)}${rowNum}`, values: [[prerequisite || ""]] });
        }
        found = true;
        break;
      }
    }

    if (!found) return res.status(404).send("Minion not found in Sectors sheet");

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "RAW", data: updates },
    });

    res.redirect("/admin/locks");
  } catch (err) {
    console.error("Lock update error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

app.post("/admin/locks/bulk-lock", async (req, res) => {
  try {
    const { sector, boss, prerequisite } = req.body;
    if (!sector || !boss) return res.status(400).send("Missing sector or boss");

    const sheets = await getSheets();
    const secRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Sectors",
    });
    const rows = secRes.data.values;
    if (!rows || rows.length < 2) return res.status(404).send("No sectors data");

    const headers = rows[0];
    const sectorCol = headers.indexOf("Sector");
    const bossCol = headers.indexOf("Boss");
    const statusCol = headers.indexOf("Status");
    const lockedCol = headers.indexOf("Locked for what?");

    const colLetter = (idx) => {
      if (idx < 26) return String.fromCharCode(65 + idx);
      return String.fromCharCode(64 + Math.floor(idx / 26)) + String.fromCharCode(65 + (idx % 26));
    };

    const updates = [];
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][sectorCol] || "").trim() === sector &&
          (rows[i][bossCol] || "").trim() === boss &&
          (rows[i][statusCol] || "").trim() === "Engaged") {
        const rowNum = i + 1;
        updates.push({ range: `Sectors!${colLetter(statusCol)}${rowNum}`, values: [["Locked"]] });
        if (lockedCol >= 0) {
          updates.push({ range: `Sectors!${colLetter(lockedCol)}${rowNum}`, values: [[prerequisite || ""]] });
        }
      }
    }

    if (updates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { valueInputOption: "RAW", data: updates },
      });
    }

    res.redirect("/admin/locks");
  } catch (err) {
    console.error("Bulk lock error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// Helper: Check and auto-unlock minions whose prerequisites are met
// ---------------------------------------------------------------------------
async function checkAndUnlockPrerequisites(sheets) {
  const secRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sectors",
  });
  const rows = secRes.data.values;
  if (!rows || rows.length < 2) return;

  const headers = rows[0];
  const sectorCol = headers.indexOf("Sector");
  const bossCol = headers.indexOf("Boss");
  const minionCol = headers.indexOf("Minion");
  const statusCol = headers.indexOf("Status");
  const lockedCol = headers.indexOf("Locked for what?");

  if (lockedCol < 0 || statusCol < 0) return;

  const colLetter = (idx) => {
    if (idx < 26) return String.fromCharCode(65 + idx);
    return String.fromCharCode(64 + Math.floor(idx / 26)) + String.fromCharCode(65 + (idx % 26));
  };

  // Build lookup of all minions by status
  const minionStatus = {}; // "BOSS>MINION" -> status
  const bossMinions = {}; // "BOSS" -> [{ minion, status }]
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const b = (r[bossCol] || "").trim();
    const m = (r[minionCol] || "").trim();
    const s = (r[statusCol] || "").trim();
    if (!b || !m) continue;
    minionStatus[b + ">" + m] = s;
    if (!bossMinions[b]) bossMinions[b] = [];
    bossMinions[b].push({ minion: m, status: s });
  }

  // Check if a boss is fully conquered (all minions Enslaved)
  function isBossConquered(bossName) {
    const minions = bossMinions[bossName];
    if (!minions || minions.length === 0) return false;
    return minions.every((m) => m.status === "Enslaved");
  }

  // Check if a specific minion is Enslaved
  function isMinionEnslaved(key) {
    return minionStatus[key] === "Enslaved";
  }

  // Find locked minions whose prerequisites are now met
  const updates = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const status = (r[statusCol] || "").trim();
    if (status !== "Locked") continue;
    const prereq = (r[lockedCol] || "").trim();
    if (!prereq) continue;

    const reqs = prereq.split(";").map((s) => s.trim()).filter(Boolean);
    const allMet = reqs.every((req) => {
      if (req.startsWith("Boss:")) {
        return isBossConquered(req.slice(5).trim());
      } else if (req.startsWith("Minion:")) {
        return isMinionEnslaved(req.slice(7).trim());
      }
      return false;
    });

    if (allMet) {
      const rowNum = i + 1;
      updates.push({ range: `Sectors!${colLetter(statusCol)}${rowNum}`, values: [["Engaged"]] });
      updates.push({ range: `Sectors!${colLetter(lockedCol)}${rowNum}`, values: [[""]] });
    }
  }

  if (updates.length > 0) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "RAW", data: updates },
    });
  }
}

// ---------------------------------------------------------------------------
// Admin: Manual Entry page — add new minions without opening Google Sheets
// ---------------------------------------------------------------------------
app.get("/admin/manual", async (req, res) => {
  try {
    const sheets = await getSheets();
    const sectorsRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Sectors",
    });
    const allMinions = parseTable(sectorsRes.data.values || []);

    // Build hierarchical data for cascading dropdowns
    const sectors = [...new Set(allMinions.map((r) => r["Sector"]).filter(Boolean))].sort();
    const sectorData = {};
    const meSubjectMap = {}; // subject → [sectors]
    for (const m of allMinions) {
      const sec = m["Sector"] || "", boss = m["Boss"] || "", subj = m["Subject"] || "", min = m["Minion"] || "";
      if (!sec || !boss) continue;
      if (!sectorData[sec]) sectorData[sec] = {};
      if (!sectorData[sec][boss]) sectorData[sec][boss] = { subject: "", minions: [] };
      if (subj && !sectorData[sec][boss].subject) sectorData[sec][boss].subject = subj;
      if (min && !sectorData[sec][boss].minions.includes(min)) sectorData[sec][boss].minions.push(min);
      if (subj) {
        if (!meSubjectMap[subj]) meSubjectMap[subj] = new Set();
        meSubjectMap[subj].add(sec);
      }
    }
    // Convert sets for JSON
    const meSubjects = {};
    for (const [subj, secs] of Object.entries(meSubjectMap)) {
      meSubjects[subj] = [...secs].sort();
    }

    const addedCount = parseInt(req.query.success) || 0;
    const successMsg = addedCount > 0
      ? `<div class="success-msg">&#x2714; ${addedCount} MINION${addedCount > 1 ? "S" : ""} ADDED SUCCESSFULLY!</div>` : "";
    const addedToQuest = req.query.quest === "1"
      ? `<div class="success-msg" style="color:#ff6600;">&#x2605; ALSO ADDED TO QUEST BOARD</div>` : "";

    res.send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Manual Entry - Sovereign HUD</title>
    <style>
    body {
        background: #0a0b10;
        color: #00f2ff;
        font-family: 'Courier New', monospace;
        padding: 20px;
        text-transform: uppercase;
        overflow-x: hidden;
    }
    .hud-container {
        border: 2px solid #ffea00;
        padding: 20px;
        box-shadow: 0 0 15px rgba(255, 234, 0, 0.5);
        max-width: 700px;
        margin: auto;
    }
    h1 {
        text-shadow: 2px 2px #ff00ff;
        border-bottom: 1px solid #ffea00;
        letter-spacing: 2px;
        text-align: center;
        margin-bottom: 5px;
        color: #ffea00;
    }
    .back-link {
        display: inline-block;
        color: #00f2ff;
        text-decoration: none;
        border: 1px solid #00f2ff;
        padding: 6px 15px;
        margin-bottom: 15px;
        font-size: 0.8em;
        transition: all 0.2s;
    }
    .back-link:hover { background: #00f2ff; color: #0a0b10; }
    .manual-subtitle {
        text-align: center;
        color: #888;
        font-size: 0.75em;
        margin-bottom: 25px;
        letter-spacing: 3px;
    }
    .form-group {
        margin-bottom: 15px;
    }
    .form-group label {
        display: block;
        font-size: 0.75em;
        color: #ffea00;
        letter-spacing: 2px;
        margin-bottom: 5px;
    }
    .form-group input, .form-group select, .form-group textarea {
        width: 100%;
        padding: 10px;
        background: #1a1d26;
        border: 1px solid #333;
        color: #00f2ff;
        font-family: 'Courier New', monospace;
        font-size: 0.85em;
        box-sizing: border-box;
    }
    .form-group input:focus, .form-group select:focus, .form-group textarea:focus {
        outline: none;
        border-color: #ffea00;
        box-shadow: 0 0 5px rgba(255, 234, 0, 0.3);
    }
    .form-group textarea { resize: vertical; min-height: 60px; text-transform: none; }
    .form-group .hint {
        font-size: 0.6em;
        color: #555;
        margin-top: 3px;
        letter-spacing: 1px;
        text-transform: none;
    }
    .form-row {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 15px;
    }
    .quest-toggle {
        display: flex;
        align-items: center;
        gap: 10px;
        margin-top: 15px;
        margin-bottom: 20px;
        padding: 12px;
        border: 1px solid rgba(255, 102, 0, 0.3);
        background: rgba(255, 102, 0, 0.05);
        border-radius: 4px;
    }
    .quest-toggle input[type="checkbox"] {
        width: 20px;
        height: 20px;
        accent-color: #ff6600;
        cursor: pointer;
    }
    .quest-toggle label {
        color: #ff6600;
        font-size: 0.8em;
        letter-spacing: 2px;
        cursor: pointer;
    }
    .submit-btn {
        width: 100%;
        padding: 14px;
        background: #ffea00;
        color: #0a0b10;
        border: none;
        font-family: 'Courier New', monospace;
        font-size: 1em;
        font-weight: bold;
        letter-spacing: 3px;
        cursor: pointer;
        transition: all 0.2s;
    }
    .submit-btn:hover {
        background: #00ff9d;
        box-shadow: 0 0 15px rgba(0, 255, 157, 0.5);
    }
    .success-msg {
        text-align: center;
        color: #00ff9d;
        font-weight: bold;
        letter-spacing: 2px;
        margin-bottom: 15px;
        padding: 10px;
        border: 1px solid rgba(0, 255, 157, 0.3);
        background: rgba(0, 255, 157, 0.05);
    }
    .required { color: #ff4444; }
    .combo-input {
        width: 100%;
        padding: 10px;
        background: #1a1d26;
        border: 1px solid #ffea00;
        color: #00f2ff;
        font-family: 'Courier New', monospace;
        font-size: 0.85em;
        box-sizing: border-box;
        margin-top: 5px;
    }
    .combo-input:focus {
        outline: none;
        border-color: #ffea00;
        box-shadow: 0 0 5px rgba(255, 234, 0, 0.3);
    }
    .combo-back-btn {
        background: none;
        border: 1px solid #555;
        color: #888;
        padding: 4px 10px;
        font-family: 'Courier New', monospace;
        font-size: 0.65em;
        cursor: pointer;
        margin-top: 5px;
        transition: all 0.2s;
    }
    .combo-back-btn:hover { border-color: #00f2ff; color: #00f2ff; }
    .me-minion-row {
        border-bottom: 1px solid rgba(0,242,255,0.15);
        padding-bottom: 10px;
        margin-bottom: 10px;
    }
    .me-minion-row:last-child { border-bottom: none; margin-bottom: 0; }
    .me-add-row-btn {
        display: block;
        width: 100%;
        padding: 10px;
        background: transparent;
        border: 1px dashed #00f2ff;
        color: #00f2ff;
        font-family: 'Courier New', monospace;
        font-size: 0.8em;
        letter-spacing: 2px;
        cursor: pointer;
        margin-top: 10px;
        transition: all 0.2s;
    }
    .me-add-row-btn:hover { background: rgba(0,242,255,0.1); border-style: solid; }
    .me-remove-row {
        background: none;
        border: 1px solid #ff4444;
        color: #ff4444;
        width: 28px;
        height: 28px;
        font-size: 1em;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        transition: all 0.2s;
        padding: 0;
    }
    .me-remove-row:hover { background: rgba(255,68,68,0.2); }
    @media (max-width: 600px) {
        body { padding: 10px; }
        .form-row { grid-template-columns: 1fr; }
    }
    </style>
</head>
<body>
    <div class="hud-container">
        <a class="back-link" href="/admin">&lt; BACK TO ADMIN</a>
        <h1>&#x270F; Manual Entry</h1>
        <div class="manual-subtitle">ADD NEW OBJECTIVES WITHOUT OPENING GOOGLE SHEETS</div>
        ${successMsg}${addedToQuest}
        <form method="POST" action="/admin/manual" id="manualForm">
            <div style="border:1px solid rgba(255,234,0,0.3);padding:15px;margin-bottom:20px;background:rgba(255,234,0,0.03);">
                <div style="font-size:0.7em;color:#ffea00;letter-spacing:2px;margin-bottom:12px;">STEP 1: SELECT SUBJECT, SECTOR &amp; BOSS</div>
                <div class="form-row" style="grid-template-columns:1fr 1fr 1fr;">
                    <div class="form-group">
                        <label>SUBJECT</label>
                        <select id="me-subject">
                            <option value="">Select subject...</option>
                            <option value="__new__">+ New subject...</option>
                        </select>
                        <input type="text" class="combo-input" id="subject-new" name="subject" placeholder="Type new subject..." style="display:none;">
                    </div>
                    <div class="form-group">
                        <label>SECTOR <span class="required">*</span></label>
                        <select id="me-sector" name="sector" required>
                            <option value="" disabled selected>Select subject or sector...</option>
                            <option value="__new__">+ New sector...</option>
                        </select>
                        <input type="text" class="combo-input" id="sector-new" placeholder="Type new sector..." style="display:none;">
                    </div>
                    <div class="form-group">
                        <label>BOSS <span class="required">*</span></label>
                        <select id="me-boss" name="boss" required disabled>
                            <option value="" disabled selected>Select sector first...</option>
                        </select>
                        <input type="text" class="combo-input" id="boss-new" placeholder="Type new boss..." style="display:none;">
                    </div>
                </div>
            </div>

            <div style="border:1px solid rgba(0,242,255,0.3);padding:15px;margin-bottom:20px;background:rgba(0,242,255,0.03);">
                <div style="font-size:0.7em;color:#00f2ff;letter-spacing:2px;margin-bottom:12px;">STEP 2: ADD MINIONS</div>
                <div id="me-minion-rows">
                    <div class="me-minion-row">
                        <div class="form-row" style="grid-template-columns:2fr 3fr 1fr 1fr auto;">
                            <div class="form-group">
                                <label>MINION NAME <span class="required">*</span></label>
                                <input type="text" name="minions[]" required placeholder="Minion name...">
                            </div>
                            <div class="form-group">
                                <label>TASK</label>
                                <input type="text" name="tasks[]" placeholder="Task description...">
                            </div>
                            <div class="form-group">
                                <label>IMPACT <span class="required">*</span></label>
                                <select name="impacts[]" required>
                                    <option value="1">1</option>
                                    <option value="2" selected>2</option>
                                    <option value="3">3</option>
                                </select>
                            </div>
                            <div class="form-group">
                                <label>STATUS</label>
                                <select name="statuses[]">
                                    <option value="Engaged" selected>Engaged</option>
                                    <option value="Locked">Locked</option>
                                </select>
                            </div>
                            <div class="form-group" style="display:flex;align-items:flex-end;">
                                <button type="button" class="me-remove-row" onclick="removeRow(this)" title="Remove" style="visibility:hidden;">&#x2715;</button>
                            </div>
                        </div>
                    </div>
                </div>
                <button type="button" id="me-add-row" class="me-add-row-btn">+ ADD ANOTHER MINION</button>
            </div>

            <div class="quest-toggle">
                <input type="checkbox" id="addToQuest" name="addToQuest" value="1">
                <label for="addToQuest">ALSO ADD ENGAGED MINIONS TO QUEST BOARD</label>
            </div>
            <button type="submit" class="submit-btn">ADD MINIONS</button>
        </form>
    </div>
    <script>
    var SD = ${JSON.stringify(sectorData)};
    var SM = ${JSON.stringify(meSubjects)};
    var meSubject = document.getElementById('me-subject');
    var meSector = document.getElementById('me-sector');
    var meBoss = document.getElementById('me-boss');
    var rowContainer = document.getElementById('me-minion-rows');

    // Populate subject dropdown
    var subjNames = Object.keys(SM).sort();
    meSubject.innerHTML = '<option value="">Select subject...</option>' +
        subjNames.map(function(s) { return '<option value="' + s + '">' + s + '</option>'; }).join('') +
        '<option value="__new__">+ New subject...</option>';

    // Populate all sectors initially
    var allSectors = Object.keys(SD).sort();
    function populateSectors(filterSectors) {
        var secs = filterSectors || allSectors;
        meSector.innerHTML = '<option value="" disabled selected>Select sector...</option>' +
            secs.map(function(s) { return '<option value="' + s + '">' + s + '</option>'; }).join('') +
            '<option value="__new__">+ New sector...</option>';
        meSector.disabled = false;
    }
    populateSectors();

    // Combo-select helpers for new entries
    function setupCombo(selId, inputId, onNew) {
        var sel = document.getElementById(selId);
        var input = document.getElementById(inputId);
        if (!sel || !input) return;
        sel.addEventListener('change', function() {
            if (sel.value === '__new__') {
                sel.style.display = 'none';
                input.style.display = 'block';
                input.focus();
                if (onNew) onNew();
            }
        });
        var backBtn = document.createElement('button');
        backBtn.type = 'button';
        backBtn.textContent = '\\u2190 BACK';
        backBtn.className = 'combo-back-btn';
        backBtn.style.display = 'none';
        input.parentNode.insertBefore(backBtn, input.nextSibling);
        input.addEventListener('focus', function() { backBtn.style.display = 'inline-block'; });
        backBtn.addEventListener('click', function() {
            input.style.display = 'none'; input.value = '';
            backBtn.style.display = 'none';
            sel.style.display = 'block'; sel.value = '';
        });
    }

    setupCombo('me-subject', 'subject-new');
    setupCombo('me-sector', 'sector-new', function() {
        // New sector → enable boss for free entry
        meBoss.innerHTML = '<option value="" disabled selected>Select boss...</option><option value="__new__">+ New boss...</option>';
        meBoss.disabled = false;
    });
    setupCombo('me-boss', 'boss-new');

    // Subject change → filter sectors
    meSubject.addEventListener('change', function() {
        if (meSubject.value === '__new__') return;
        var subj = meSubject.value;
        meBoss.innerHTML = '<option value="" disabled selected>Select sector first...</option>';
        meBoss.disabled = true;
        if (!subj || !SM[subj]) {
            populateSectors();
            return;
        }
        populateSectors(SM[subj]);
        // If only one sector, auto-select it
        if (SM[subj].length === 1) {
            meSector.value = SM[subj][0];
            meSector.dispatchEvent(new Event('change'));
        }
    });

    // Sector change → populate boss
    meSector.addEventListener('change', function() {
        if (meSector.value === '__new__') return;
        var sec = meSector.value;
        var bossNew = document.getElementById('boss-new');
        if (bossNew) { bossNew.style.display = 'none'; bossNew.value = ''; }
        meBoss.style.display = 'block';

        if (!sec || !SD[sec]) {
            meBoss.innerHTML = '<option value="" disabled selected>Select sector first...</option>';
            meBoss.disabled = true;
            return;
        }
        var bosses = Object.keys(SD[sec]).sort();
        var subj = meSubject.value && meSubject.value !== '__new__' ? meSubject.value : null;
        // Filter bosses by subject if one is selected
        if (subj) {
            bosses = bosses.filter(function(b) { return SD[sec][b].subject === subj; });
        }
        var html = '<option value="" disabled selected>Select boss...</option>';
        for (var i = 0; i < bosses.length; i++) {
            var b = bosses[i];
            var bs = SD[sec][b].subject;
            html += '<option value="' + b + '">' + b + (bs ? ' (' + bs + ')' : '') + '</option>';
        }
        html += '<option value="__new__">+ New boss...</option>';
        meBoss.innerHTML = html;
        meBoss.disabled = false;
        // Auto-select subject if boss determines it
    });

    // Boss change → auto-set subject if not already set
    meBoss.addEventListener('change', function() {
        if (meBoss.value === '__new__') return;
        var sec = meSector.value;
        var boss = meBoss.value;
        if (sec && boss && SD[sec] && SD[sec][boss] && SD[sec][boss].subject) {
            var subjInput = document.getElementById('subject-new');
            if (subjInput && subjInput.style.display !== 'none') return; // user typing new
            if (!meSubject.value || meSubject.value === '') {
                meSubject.value = SD[sec][boss].subject;
            }
        }
    });

    // Add minion row
    var rowTemplate = rowContainer.querySelector('.me-minion-row').innerHTML;
    document.getElementById('me-add-row').addEventListener('click', function() {
        var div = document.createElement('div');
        div.className = 'me-minion-row';
        div.innerHTML = rowTemplate;
        // Make remove button visible for all rows after first
        var rmBtn = div.querySelector('.me-remove-row');
        if (rmBtn) rmBtn.style.visibility = 'visible';
        // Reset values
        div.querySelectorAll('input[type="text"]').forEach(function(inp) { inp.value = ''; });
        rowContainer.appendChild(div);
        updateRemoveButtons();
        div.querySelector('input[name="minions[]"]').focus();
    });

    function removeRow(btn) {
        var row = btn.closest('.me-minion-row');
        if (row) row.remove();
        updateRemoveButtons();
    }
    window.removeRow = removeRow;

    function updateRemoveButtons() {
        var rows = rowContainer.querySelectorAll('.me-minion-row');
        rows.forEach(function(r, i) {
            var btn = r.querySelector('.me-remove-row');
            if (btn) btn.style.visibility = rows.length > 1 ? 'visible' : 'hidden';
        });
    }

    // On submit: copy combo-input values to hidden fields
    document.getElementById('manualForm').addEventListener('submit', function(e) {
        // Subject: if text input visible, use it
        var subjNew = document.getElementById('subject-new');
        if (subjNew && subjNew.style.display !== 'none' && subjNew.value.trim()) {
            // subject-new already has name="subject"
        } else if (meSubject.value && meSubject.value !== '__new__') {
            // Remove name from text input to avoid duplicate field
            if (subjNew) subjNew.removeAttribute('name');
            var h = document.createElement('input'); h.type = 'hidden'; h.name = 'subject'; h.value = meSubject.value;
            this.appendChild(h);
        }
        // Sector: if text input visible, use it
        var secNew = document.getElementById('sector-new');
        if (secNew && secNew.style.display !== 'none' && secNew.value.trim()) {
            var h2 = document.createElement('input'); h2.type = 'hidden'; h2.name = 'sector'; h2.value = secNew.value.trim();
            this.appendChild(h2);
            meSector.removeAttribute('name');
        }
        // Boss: if text input visible, use it
        var bossNew = document.getElementById('boss-new');
        if (bossNew && bossNew.style.display !== 'none' && bossNew.value.trim()) {
            var h3 = document.createElement('input'); h3.type = 'hidden'; h3.name = 'boss'; h3.value = bossNew.value.trim();
            this.appendChild(h3);
            meBoss.removeAttribute('name');
        }
    });
    </script>
</body>
</html>`);
  } catch (err) {
    console.error("Manual entry page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

app.post("/admin/manual", async (req, res) => {
  try {
    const { sector, boss, subject, addToQuest } = req.body;
    // Support both single and array form fields
    const minions = [].concat(req.body["minions[]"] || req.body.minion || []).filter(Boolean);
    const tasks = [].concat(req.body["tasks[]"] || req.body.task || []);
    const impacts = [].concat(req.body["impacts[]"] || req.body.impact || []);
    const statuses = [].concat(req.body["statuses[]"] || req.body.status || []);

    if (!sector || !boss || minions.length === 0) {
      return res.status(400).send("Missing required fields: sector, boss, at least one minion");
    }

    const sheets = await getSheets();

    // Get Sectors headers to build rows in correct order
    const sectorsRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Sectors!1:1",
    });
    const headers = sectorsRes.data.values ? sectorsRes.data.values[0] : [];
    if (headers.length === 0) {
      return res.status(500).send("Sectors sheet has no headers");
    }

    const statusIdx = headers.indexOf("Status");
    const rows = [];
    for (let i = 0; i < minions.length; i++) {
      const minionName = minions[i].trim();
      if (!minionName) continue;
      const minionStatus = statuses[i] || "Engaged";
      const row = buildSectorsRow(headers, {
        sector,
        boss,
        minion: minionName,
        subject: subject || "",
        task: tasks[i] || "",
        impact: parseInt(impacts[i]) || 1,
      });
      if (statusIdx !== -1) row[statusIdx] = minionStatus;
      rows.push({ row, minionName, task: tasks[i] || "", status: minionStatus });
    }

    if (rows.length === 0) {
      return res.status(400).send("No valid minions provided");
    }

    // Batch append all rows at once
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Sectors!A:Z",
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: { values: rows.map(r => r.row) },
    });

    // Optionally add engaged minions to quest board
    let questAdded = false;
    if (addToQuest === "1") {
      const engagedRows = rows.filter(r => r.status === "Engaged");
      if (engagedRows.length > 0) {
        await ensureQuestsSheet(sheets);
        const defRes = await sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: "Definitions",
        });
        const definitions = parseTable(defRes.data.values);
        const { proofType, suggestion } = generateProofSuggestion(sector, definitions);
        const today = new Date().toISOString().slice(0, 10);
        const questRows = [];

        for (const er of engagedRows) {
          const questId = generateQuestId();
          const taskDetail = er.task || suggestion;
          questRows.push([questId, boss, er.minionName, sector, "Active", proofType, "", taskDetail, "", today, "", ""]);
        }

        await sheets.spreadsheets.values.append({
          spreadsheetId: SPREADSHEET_ID,
          range: "Quests!A:L",
          valueInputOption: "RAW",
          insertDataOption: "INSERT_ROWS",
          requestBody: { values: questRows },
        });

        for (const er of engagedRows) {
          await updateSectorsQuestStatus(sheets, sector, boss, er.minionName, "Active");
        }
        questAdded = true;
      }
    }

    res.redirect(`/admin/manual?success=${rows.length}${questAdded ? '&quest=1' : ''}`);
  } catch (err) {
    console.error("Manual entry error:", err);
    res.status(500).send(`<pre style="color:red">Error adding minion(s): ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// Admin: Teacher Notes
// ---------------------------------------------------------------------------
app.get("/admin/notes", async (req, res) => {
  try {
    const sheets = await getSheets();
    await ensureTeacherNotesSheet(sheets);
    const notesRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Teacher_Notes",
    });
    const notes = parseTable(notesRes.data.values || []);
    const sorted = notes.sort((a, b) => (b["Date"] || "").localeCompare(a["Date"] || ""));
    const authorName = req.cookies.userName || "Unknown";

    let notesHtml = "";
    if (sorted.length === 0) {
      notesHtml = `<div style="text-align:center;color:#555;padding:30px;">NO NOTES YET. ADD THE FIRST ONE BELOW.</div>`;
    } else {
      for (const n of sorted) {
        notesHtml += `
          <div class="note-card">
            <div class="note-header">
              <span class="note-date">${escHtml(n["Date"] || "")}</span>
              <span class="note-author">${escHtml(n["Author"] || "")}</span>
              ${n["Subject"] ? `<span class="note-subject">${escHtml(n["Subject"])}</span>` : ""}
            </div>
            <div class="note-body">${escHtml(n["Note"] || "")}</div>
          </div>`;
      }
    }

    res.send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Teacher Notes - Sovereign HUD</title>
    <style>
    body { background: #0a0b10; color: #00f2ff; font-family: 'Courier New', monospace; padding: 20px; text-transform: uppercase; }
    .hud-container { border: 2px solid #ff00ff; padding: 20px; box-shadow: 0 0 15px rgba(255,0,255,0.3); max-width: 800px; margin: auto; }
    .back-link { display: inline-block; color: #00f2ff; text-decoration: none; border: 1px solid #00f2ff; padding: 6px 15px; margin-bottom: 15px; font-size: 0.8em; transition: all 0.2s; }
    .back-link:hover { background: #00f2ff; color: #0a0b10; }
    h1 { text-align: center; color: #ffea00; text-shadow: 2px 2px #ff00ff; letter-spacing: 4px; margin: 15px 0 5px; }
    .notes-subtitle { text-align: center; font-size: 0.7em; color: #888; letter-spacing: 2px; margin-bottom: 20px; }
    .note-form { border: 1px solid rgba(255,0,255,0.3); padding: 15px; margin-bottom: 25px; background: rgba(255,0,255,0.03); }
    .note-form label { display: block; font-size: 0.7em; color: #ff00ff; letter-spacing: 2px; margin-bottom: 5px; }
    .note-form input, .note-form textarea { width: 100%; padding: 10px; background: #1a1d26; border: 1px solid #333; color: #00f2ff; font-family: 'Courier New', monospace; font-size: 0.85em; box-sizing: border-box; margin-bottom: 10px; }
    .note-form textarea { min-height: 80px; resize: vertical; text-transform: none; }
    .note-form input:focus, .note-form textarea:focus { outline: none; border-color: #ff00ff; box-shadow: 0 0 5px rgba(255,0,255,0.3); }
    .note-submit { width: 100%; padding: 12px; background: #ff00ff; color: #0a0b10; border: none; font-family: 'Courier New', monospace; font-size: 0.9em; font-weight: bold; letter-spacing: 3px; cursor: pointer; transition: all 0.2s; }
    .note-submit:hover { background: #ffea00; box-shadow: 0 0 15px rgba(255,234,0,0.5); }
    .note-card { border: 1px solid rgba(255,0,255,0.2); padding: 12px; margin-bottom: 10px; background: rgba(255,0,255,0.02); }
    .note-header { display: flex; gap: 12px; align-items: center; margin-bottom: 8px; flex-wrap: wrap; }
    .note-date { font-size: 0.7em; color: #888; }
    .note-author { font-size: 0.7em; color: #ff00ff; font-weight: bold; letter-spacing: 1px; }
    .note-subject { font-size: 0.7em; color: #ffea00; border: 1px solid rgba(255,234,0,0.3); padding: 1px 6px; }
    .note-body { font-size: 0.85em; color: #ccc; text-transform: none; line-height: 1.5; }
    .success-msg { text-align: center; color: #00ff9d; font-size: 0.8em; padding: 8px; border: 1px solid rgba(0,255,157,0.3); background: rgba(0,255,157,0.05); margin-bottom: 15px; }
    @media (max-width: 600px) { body { padding: 10px; } }
    </style>
</head>
<body>
    <div class="hud-container">
        <a class="back-link" href="/admin">&lt; BACK TO ADMIN</a>
        <h1>&#x1F4DD; Teacher Notes</h1>
        <div class="notes-subtitle">OBSERVATIONS, PLANS &amp; COMMUNICATION</div>
        ${req.query.added === "1" ? '<div class="success-msg">&#x2714; NOTE ADDED</div>' : ""}
        <div class="note-form">
            <form method="POST" action="/admin/notes">
                <label>SUBJECT (optional)</label>
                <input type="text" name="subject" placeholder="e.g. Math progress, Behavior, Schedule change...">
                <label>NOTE</label>
                <textarea name="note" required placeholder="Write your note here..."></textarea>
                <button type="submit" class="note-submit">ADD NOTE</button>
            </form>
        </div>
        <div style="font-size:0.7em;color:#888;letter-spacing:2px;margin-bottom:12px;">${sorted.length} NOTE${sorted.length !== 1 ? "S" : ""}</div>
        ${notesHtml}
    </div>
</body>
</html>`);
  } catch (err) {
    console.error("Teacher notes page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

app.post("/admin/notes", async (req, res) => {
  try {
    const note = (req.body.note || "").trim();
    const subject = (req.body.subject || "").trim();
    if (!note) return res.redirect("/admin/notes");

    const author = req.cookies.userName || "Unknown";
    const today = new Date().toISOString().slice(0, 10);

    const sheets = await getSheets();
    await ensureTeacherNotesSheet(sheets);
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Teacher_Notes!A:D",
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: { values: [[today, author, subject, note]] },
    });

    res.redirect("/admin/notes?added=1");
  } catch (err) {
    console.error("Add note error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// Admin: Photo Import page
// ---------------------------------------------------------------------------
function buildImportPage() {
  return `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Photo Import - Sovereign HUD</title>
    <style>
    body {
        background: #0a0b10;
        color: #ffea00;
        font-family: 'Courier New', monospace;
        padding: 20px;
        text-transform: uppercase;
        overflow-x: hidden;
    }
    .hud-container {
        border: 2px solid #ffea00;
        padding: 20px;
        box-shadow: 0 0 15px rgba(255, 234, 0, 0.5);
        max-width: 960px;
        margin: auto;
    }
    h1 {
        text-shadow: 2px 2px #ff00ff;
        border-bottom: 1px solid #ffea00;
        letter-spacing: 2px;
        text-align: center;
        margin-bottom: 5px;
        color: #ffea00;
    }
    .back-link {
        display: inline-block;
        color: #00f2ff;
        text-decoration: none;
        border: 1px solid #00f2ff;
        padding: 6px 15px;
        margin-bottom: 15px;
        font-size: 0.8em;
        transition: all 0.2s;
    }
    .back-link:hover { background: #00f2ff; color: #0a0b10; }
    .import-subtitle {
        text-align: center;
        color: #888;
        font-size: 0.8em;
        margin-bottom: 25px;
        letter-spacing: 3px;
    }
    .drop-zone {
        border: 2px dashed #ffea00;
        border-radius: 8px;
        padding: 50px 20px;
        text-align: center;
        cursor: pointer;
        transition: all 0.3s;
        margin-bottom: 20px;
    }
    .drop-zone:hover, .drop-zone.dragover {
        background: rgba(255, 234, 0, 0.05);
        box-shadow: 0 0 20px rgba(255, 234, 0, 0.3);
    }
    .drop-zone-text {
        font-size: 0.9em;
        letter-spacing: 2px;
        margin-bottom: 8px;
    }
    .drop-zone-hint {
        font-size: 0.7em;
        color: #666;
        text-transform: none;
    }
    .file-list {
        margin: 15px 0;
        font-size: 0.8em;
    }
    .file-item {
        display: inline-block;
        background: rgba(255, 234, 0, 0.1);
        border: 1px solid #333;
        border-radius: 4px;
        padding: 4px 10px;
        margin: 3px;
        font-size: 0.85em;
    }
    .btn {
        display: inline-block;
        background: #0a0b10;
        border: 2px solid #ffea00;
        color: #ffea00;
        padding: 12px 25px;
        font-family: 'Courier New', monospace;
        font-size: 0.9em;
        font-weight: bold;
        letter-spacing: 2px;
        text-transform: uppercase;
        cursor: pointer;
        transition: all 0.3s;
    }
    .btn:hover:not(:disabled) {
        background: #ffea00;
        color: #0a0b10;
    }
    .btn:disabled {
        opacity: 0.3;
        cursor: not-allowed;
    }
    .btn-approve {
        border-color: #00ff9d;
        color: #00ff9d;
    }
    .btn-approve:hover:not(:disabled) {
        background: #00ff9d;
        color: #0a0b10;
    }
    .btn-small {
        padding: 5px 12px;
        font-size: 0.7em;
    }
    .progress {
        text-align: center;
        padding: 30px;
        font-size: 0.9em;
        letter-spacing: 2px;
    }
    .progress .spinner {
        display: inline-block;
        animation: spin 1s linear infinite;
        margin-right: 8px;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    .results-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin: 20px 0 15px;
        padding-bottom: 10px;
        border-bottom: 1px solid #333;
    }
    .results-stats {
        font-size: 0.8em;
        color: #888;
    }
    .result-card {
        border: 1px solid #333;
        border-radius: 6px;
        padding: 15px;
        margin-bottom: 12px;
        transition: all 0.3s;
    }
    .result-card.approved {
        border-color: #00ff9d;
        background: rgba(0, 255, 157, 0.03);
    }
    .result-card.skipped {
        border-color: #333;
        opacity: 0.4;
    }
    .result-card.error {
        border-color: #ff3333;
        background: rgba(255, 51, 51, 0.05);
    }
    .result-card-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 10px;
    }
    .result-card-file {
        font-size: 0.8em;
        color: #00f2ff;
    }
    .result-card-actions button {
        margin-left: 6px;
    }
    .result-fields {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 8px;
        margin-bottom: 10px;
    }
    .result-field label {
        display: block;
        font-size: 0.65em;
        color: #666;
        margin-bottom: 2px;
        letter-spacing: 1px;
    }
    .result-field select,
    .result-field input {
        width: 100%;
        background: #111;
        border: 1px solid #333;
        color: #ffea00;
        padding: 6px 8px;
        font-family: 'Courier New', monospace;
        font-size: 0.85em;
        box-sizing: border-box;
    }
    .result-field select:focus,
    .result-field input:focus {
        border-color: #ffea00;
        outline: none;
    }
    .result-reasoning {
        font-size: 0.7em;
        color: #666;
        text-transform: none;
        font-style: italic;
        margin-top: 6px;
    }
    .result-confidence {
        font-size: 0.7em;
        letter-spacing: 1px;
    }
    .result-confidence.high { color: #00ff9d; }
    .result-confidence.medium { color: #ffea00; }
    .result-confidence.low { color: #ff3333; }
    .result-dupe {
        font-size: 0.7em;
        color: #ff3333;
        letter-spacing: 1px;
        margin-left: 8px;
    }
    .status-msg {
        text-align: center;
        padding: 20px;
        font-size: 0.9em;
        letter-spacing: 2px;
    }
    .status-msg.success { color: #00ff9d; }
    .status-msg.error { color: #ff3333; }
    .hidden { display: none; }
    @media (max-width: 600px) {
        body { padding: 10px; }
        .result-fields { grid-template-columns: 1fr; }
    }
    </style>
</head>
<body>
    <div class="hud-container">
        <a class="back-link" href="/admin">&lt; BACK TO ADMIN</a>
        <h1>&#x1F4F7; Lesson Photo Import</h1>
        <div class="import-subtitle">AI-POWERED CLASSIFICATION</div>

        <div id="upload-section">
            <div class="drop-zone" id="dropZone">
                <div class="drop-zone-text">DRAG &amp; DROP LESSON PHOTOS HERE</div>
                <div class="drop-zone-hint">or click to browse &mdash; accepts .jpg, .png, .webp</div>
            </div>
            <input type="file" id="fileInput" multiple accept=".jpg,.jpeg,.png,.webp" class="hidden">
            <div class="file-list" id="fileList"></div>
            <div style="text-align:center;">
                <button class="btn" id="analyzeBtn" disabled>ANALYZE WITH AI &gt;&gt;</button>
            </div>
        </div>

        <div id="progress-section" class="hidden">
            <div class="progress">
                <span class="spinner">&#x2699;</span>
                <span id="progressText">ANALYZING PHOTOS...</span>
            </div>
        </div>

        <div id="results-section" class="hidden">
            <div class="results-header">
                <span>CLASSIFICATION RESULTS</span>
                <span class="results-stats" id="resultsStats"></span>
            </div>
            <div id="resultCards"></div>
            <div style="text-align:center;margin-top:20px;">
                <button class="btn btn-approve" id="writeBtn">WRITE TO SHEETS &gt;&gt;</button>
            </div>
        </div>

        <div id="status-section" class="hidden">
            <div class="status-msg" id="statusMsg"></div>
            <div style="text-align:center;margin-top:15px;">
                <button class="btn" onclick="location.reload()">IMPORT MORE</button>
                <a href="/admin" class="btn" style="text-decoration:none;margin-left:10px;">BACK TO ADMIN</a>
            </div>
        </div>
    </div>

    <script>
    (function() {
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const fileList = document.getElementById('fileList');
        const analyzeBtn = document.getElementById('analyzeBtn');
        const progressSection = document.getElementById('progress-section');
        const progressText = document.getElementById('progressText');
        const uploadSection = document.getElementById('upload-section');
        const resultsSection = document.getElementById('results-section');
        const resultsStats = document.getElementById('resultsStats');
        const resultCards = document.getElementById('resultCards');
        const writeBtn = document.getElementById('writeBtn');
        const statusSection = document.getElementById('status-section');
        const statusMsg = document.getElementById('statusMsg');

        let selectedFiles = [];
        let analysisResults = [];
        let validSectors = [];
        let existingBosses = {};

        // Drag and drop
        dropZone.addEventListener('click', () => fileInput.click());
        dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            handleFiles(e.dataTransfer.files);
        });
        fileInput.addEventListener('change', () => handleFiles(fileInput.files));

        function handleFiles(files) {
            selectedFiles = Array.from(files).filter(f => /\\.(jpg|jpeg|png|webp)$/i.test(f.name));
            fileList.innerHTML = selectedFiles.map(f =>
                '<span class="file-item">' + esc(f.name) + '</span>'
            ).join('');
            analyzeBtn.disabled = selectedFiles.length === 0;
        }

        function esc(s) {
            const d = document.createElement('div');
            d.textContent = s;
            return d.innerHTML;
        }

        // Analyze
        analyzeBtn.addEventListener('click', async () => {
            if (selectedFiles.length === 0) return;
            uploadSection.classList.add('hidden');
            progressSection.classList.remove('hidden');
            progressText.textContent = 'UPLOADING ' + selectedFiles.length + ' PHOTO(S)...';

            const formData = new FormData();
            for (const f of selectedFiles) formData.append('photos', f);

            try {
                const resp = await fetch('/admin/import/upload', { method: 'POST', body: formData });
                const data = await resp.json();
                if (data.error) throw new Error(data.error);

                analysisResults = data.results;
                validSectors = data.validSectors || [];
                existingBosses = data.existingBosses || {};
                renderResults();
            } catch (err) {
                progressSection.classList.add('hidden');
                statusSection.classList.remove('hidden');
                statusMsg.className = 'status-msg error';
                statusMsg.textContent = 'ERROR: ' + err.message;
            }
        });

        function renderResults() {
            progressSection.classList.add('hidden');
            resultsSection.classList.remove('hidden');

            const totalInput = analysisResults.reduce((s, r) => s + (r.usage ? r.usage.input : 0), 0);
            const totalOutput = analysisResults.reduce((s, r) => s + (r.usage ? r.usage.output : 0), 0);
            const cost = (totalInput * 3 + totalOutput * 15) / 1000000;
            const errorCount = analysisResults.filter(r => r.error).length;
            const okCount = analysisResults.length - errorCount;
            resultsStats.textContent = okCount + ' ANALYZED | ' + errorCount + ' ERRORS | EST. COST: $' + cost.toFixed(4);

            const sectorOpts = validSectors.map(s => '<option value="' + esc(s) + '">' + esc(s) + '</option>').join('');

            let html = '';
            for (let i = 0; i < analysisResults.length; i++) {
                const r = analysisResults[i];
                if (r.error) {
                    html += '<div class="result-card error" data-idx="' + i + '" data-approved="false">' +
                        '<div class="result-card-header">' +
                            '<span class="result-card-file">#' + (i+1) + ' ' + esc(r.filename) + '</span>' +
                            '<span style="color:#ff3333;font-size:0.8em;">ERROR: ' + esc(r.error) + '</span>' +
                        '</div></div>';
                    continue;
                }
                const res = r.result;
                const confClass = (res.confidence || '').toLowerCase();
                const isDupe = r.isDuplicate ? '<span class="result-dupe">[DUPLICATE]</span>' : '';
                html += '<div class="result-card approved" data-idx="' + i + '" data-approved="true">' +
                    '<div class="result-card-header">' +
                        '<span class="result-card-file">#' + (i+1) + ' ' + esc(r.filename) +
                            ' <span class="result-confidence ' + confClass + '">' + esc(res.confidence || '') + '</span>' + isDupe + '</span>' +
                        '<span class="result-card-actions">' +
                            '<button class="btn btn-small btn-approve" onclick="toggleCard(' + i + ',true)">APPROVE</button>' +
                            '<button class="btn btn-small" onclick="toggleCard(' + i + ',false)">SKIP</button>' +
                        '</span>' +
                    '</div>' +
                    '<div class="result-fields">' +
                        '<div class="result-field"><label>SECTOR</label><select data-field="sector">' + sectorOpts + '</select></div>' +
                        '<div class="result-field"><label>BOSS</label><input data-field="boss" value="' + esc(res.boss || '') + '"></div>' +
                        '<div class="result-field"><label>MINION</label><input data-field="minion" value="' + esc(res.minion || '') + '"></div>' +
                        '<div class="result-field"><label>IMPACT (1-3)</label><select data-field="impact">' +
                            '<option value="1"' + (res.impact===1?' selected':'') + '>1 - Simple</option>' +
                            '<option value="2"' + (res.impact===2?' selected':'') + '>2 - Moderate</option>' +
                            '<option value="3"' + (res.impact===3?' selected':'') + '>3 - Complex</option>' +
                        '</select></div>' +
                    '</div>' +
                    '<div class="result-reasoning">' + esc(res.reasoning || '') + '</div>' +
                '</div>';
            }
            resultCards.innerHTML = html;

            // Set sector dropdown values
            for (let i = 0; i < analysisResults.length; i++) {
                const r = analysisResults[i];
                if (r.error) continue;
                const card = document.querySelector('[data-idx="' + i + '"]');
                const sel = card.querySelector('[data-field="sector"]');
                if (sel && r.result.sector) sel.value = r.result.sector;
            }

            updateWriteBtn();
        }

        window.toggleCard = function(idx, approved) {
            const card = document.querySelector('[data-idx="' + idx + '"]');
            card.setAttribute('data-approved', approved ? 'true' : 'false');
            card.className = 'result-card ' + (approved ? 'approved' : 'skipped');
            updateWriteBtn();
        };

        function updateWriteBtn() {
            const count = document.querySelectorAll('.result-card[data-approved="true"]').length;
            writeBtn.textContent = 'WRITE ' + count + ' APPROVED TO SHEETS >>';
            writeBtn.disabled = count === 0;
        }

        // Write to sheets
        writeBtn.addEventListener('click', async () => {
            const approved = [];
            document.querySelectorAll('.result-card[data-approved="true"]').forEach(card => {
                const idx = parseInt(card.getAttribute('data-idx'));
                const r = analysisResults[idx];
                if (!r || r.error) return;
                approved.push({
                    storedAs: r.storedAs,
                    sector: card.querySelector('[data-field="sector"]').value,
                    boss: card.querySelector('[data-field="boss"]').value,
                    minion: card.querySelector('[data-field="minion"]').value,
                    impact: parseInt(card.querySelector('[data-field="impact"]').value) || 2,
                });
            });

            if (approved.length === 0) return;
            resultsSection.classList.add('hidden');
            progressSection.classList.remove('hidden');
            progressText.textContent = 'WRITING ' + approved.length + ' ROW(S) TO GOOGLE SHEETS...';

            try {
                const resp = await fetch('/admin/import/approve', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ approved }),
                });
                const data = await resp.json();
                if (data.error) throw new Error(data.error);

                progressSection.classList.add('hidden');
                statusSection.classList.remove('hidden');
                statusMsg.className = 'status-msg success';
                statusMsg.textContent = data.count + ' MINION(S) ENSLAVED!';
            } catch (err) {
                progressSection.classList.add('hidden');
                statusSection.classList.remove('hidden');
                statusMsg.className = 'status-msg error';
                statusMsg.textContent = 'ERROR: ' + err.message;
            }
        });
    })();
    </script>
</body>
</html>`;
}

app.get("/admin/import", (req, res) => {
  res.send(buildImportPage());
});

// ---------------------------------------------------------------------------
// Admin: Upload + analyze photos
// ---------------------------------------------------------------------------
app.post("/admin/import/upload", upload.array("photos", 20), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: "No files uploaded" });
    }
    if (!ANTHROPIC_API_KEY) {
      return res.status(500).json({ error: "ANTHROPIC_API_KEY not configured in .env" });
    }

    const sheets = await getSheets();
    const context = await fetchImportContext(sheets);
    const prompt = buildImportPrompt(context);
    const anthropic = new Anthropic({ apiKey: ANTHROPIC_API_KEY });

    const results = [];
    for (let i = 0; i < req.files.length; i++) {
      const file = req.files[i];
      try {
        const analysis = await analyzeImageAI(anthropic, file.path, prompt);
        const isDupe = context.minionSet.has((analysis.result.minion || "").toLowerCase());
        results.push({
          filename: file.originalname,
          storedAs: file.filename,
          result: analysis.result,
          usage: analysis.usage,
          isDuplicate: isDupe,
        });
      } catch (err) {
        results.push({ filename: file.originalname, storedAs: file.filename, error: err.message });
      }
      if (i < req.files.length - 1) await new Promise((r) => setTimeout(r, 500));
    }

    res.json({
      results,
      validSectors: context.validSectors,
      existingBosses: Object.fromEntries(
        Object.entries(context.bossMap).map(([s, bs]) => [s, Array.from(bs)])
      ),
    });
  } catch (err) {
    console.error("Import upload error:", err);
    res.status(500).json({ error: err.message });
  }
});

// ---------------------------------------------------------------------------
// Admin: Approve + write to sheets
// ---------------------------------------------------------------------------
app.post("/admin/import/approve", async (req, res) => {
  try {
    const { approved } = req.body;
    if (!approved || approved.length === 0) {
      return res.status(400).json({ error: "No items to approve" });
    }

    const sheets = await getSheets();
    const context = await fetchImportContext(sheets);
    const rows = approved.map((item) => buildSectorsRow(context.headers, item));

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Sectors!A:Z",
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: { values: rows },
    });

    // Move processed images to done/
    for (const item of approved) {
      const src = path.join(IMPORT_DIR, item.storedAs);
      const dest = path.join(DONE_DIR, item.storedAs);
      if (fs.existsSync(src)) {
        fs.renameSync(src, dest);
      }
    }

    res.json({ success: true, count: approved.length });
  } catch (err) {
    console.error("Import approve error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Sovereign HUD online at http://localhost:${PORT}`);
});

const express = require("express");
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
// Ensure the Quests sheet tab exists (auto-create with headers if missing)
// ---------------------------------------------------------------------------
async function ensureQuestsSheet(sheets) {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: "sheets.properties.title",
  });
  const titles = meta.data.sheets.map((s) => s.properties.title);
  if (titles.includes("Quests")) return;

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [{ addSheet: { properties: { title: "Quests" } } }],
    },
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: "Quests!A1:I1",
    valueInputOption: "RAW",
    requestBody: {
      values: [["Quest ID", "Boss", "Minion", "Sector", "Status", "Proof Type", "Proof Link", "Suggested By AI", "Date Completed"]],
    },
  });
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

  if (questStatusCol < 0 || sectorCol < 0 || bossCol < 0 || minionCol < 0) return;

  const updates = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][sectorCol] === sector && rows[i][bossCol] === boss && rows[i][minionCol] === minion) {
      const rowNum = i + 1; // 1-based for Sheets API
      const qCol = String.fromCharCode(65 + questStatusCol);
      updates.push({ range: `Sectors!${qCol}${rowNum}`, values: [[questStatus]] });

      // Auto-enslave when approved
      if (questStatus === "Approved" && statusCol >= 0) {
        const sCol = String.fromCharCode(65 + statusCol);
        updates.push({ range: `Sectors!${sCol}${rowNum}`, values: [["Enslaved"]] });
      }
      break;
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
        <div class="survival-ring-title">${shieldIcon} RING OF GUARDIANS</div>
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
        border-bottom: 1px solid #00f2ff;
        letter-spacing: 2px;
        text-align: center;
        margin-bottom: 15px;
    }
    h1.hud-title {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 14px;
        padding: 15px 0;
        margin-bottom: 20px;
        position: relative;
    }
    h1.hud-title span {
        background: linear-gradient(90deg, #00f2ff, #ff00ff, #ffea00, #00ff9d, #00f2ff);
        background-size: 300% 100%;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        animation: titleShimmer 6s linear infinite;
        font-size: 1.4em;
        letter-spacing: 5px;
        filter: drop-shadow(0 0 12px rgba(0, 242, 255, 0.4)) drop-shadow(0 0 25px rgba(255, 0, 255, 0.2));
    }
    @keyframes titleShimmer {
        0% { background-position: 0% 50%; }
        100% { background-position: 300% 50%; }
    }
    h1.hud-title::before,
    h1.hud-title::after {
        content: "";
        flex: 1;
        height: 1px;
        max-width: 80px;
        background: linear-gradient(90deg, transparent, #00f2ff, transparent);
    }

    .mc-sprite { filter: drop-shadow(0 0 6px #00f2ff) drop-shadow(0 0 12px rgba(255,0,255,0.3)); flex-shrink: 0; }

    /* Confidence row: bar + radar side by side */
    .confidence-row {
        display: flex;
        align-items: center;
        gap: 25px;
        margin-bottom: 25px;
        border-bottom: 1px dashed #ffea00;
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
        border-left: 1px solid rgba(0, 242, 255, 0.2);
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

    /* Sector map */
    .sector-map {
        margin-top: 10px;
        border-top: 1px dashed #00f2ff;
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
        display: flex;
        flex-wrap: wrap;
        justify-content: center;
        gap: 40px;
    }
    .sector-zone {
        border: 1px solid rgba(0, 242, 255, 0.3);
        background: rgba(0, 242, 255, 0.03);
        border-radius: 8px;
        padding: 25px 15px 15px 15px;
        position: relative;
        min-width: 220px;
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 12px;
        flex: 1 1 auto;
        max-width: 420px;
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
        gap: 20px;
    }
    .boss-link {
        text-decoration: none;
        color: inherit;
    }
    .boss-orb-container {
        display: flex;
        flex-direction: column;
        align-items: center;
        width: 100px;
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
        font-size: 0.7em;
        color: #00f2ff;
        margin-top: 8px;
        text-align: center;
        word-wrap: break-word;
        line-height: 1.2;
        width: 100%;
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
        border-top: 1px dashed #ffd700;
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

    /* Army floating action button */
    .army-fab {
        position: fixed;
        bottom: 20px;
        left: 20px;
        text-decoration: none;
        background: #0a0b10;
        border: 2px solid #00ff9d;
        color: #00ff9d;
        padding: 10px 18px;
        font-family: 'Courier New', monospace;
        font-size: 0.8em;
        font-weight: bold;
        letter-spacing: 2px;
        text-transform: uppercase;
        z-index: 1000;
        box-shadow: 0 0 15px rgba(0, 255, 157, 0.3), 0 0 30px rgba(0, 255, 157, 0.1);
        transition: all 0.3s;
    }
    .army-fab:hover {
        background: #00ff9d;
        color: #0a0b10;
        box-shadow: 0 0 25px rgba(0, 255, 157, 0.6);
    }

    /* Quest floating action button */
    .quest-fab {
        position: fixed;
        bottom: 20px;
        right: 20px;
        text-decoration: none;
        background: #0a0b10;
        border: 2px solid #ff6600;
        color: #ff6600;
        padding: 10px 18px;
        font-family: 'Courier New', monospace;
        font-size: 0.8em;
        font-weight: bold;
        letter-spacing: 2px;
        text-transform: uppercase;
        z-index: 1000;
        box-shadow: 0 0 15px rgba(255, 102, 0, 0.4), 0 0 30px rgba(255, 102, 0, 0.1);
        transition: all 0.3s;
        animation: pulse 2s infinite;
    }
    .quest-fab:hover {
        background: #ff6600;
        color: #0a0b10;
        box-shadow: 0 0 25px rgba(255, 102, 0, 0.6);
    }
    .quest-fab.dim {
        border-color: #333;
        color: #666;
        box-shadow: none;
        animation: none;
    }

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
    }
    .survival-achieved .survival-ring-title {
        color: #ffd700;
        text-shadow: 0 0 8px rgba(255, 215, 0, 0.5);
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
            border-left: none;
            border-top: 1px solid rgba(0, 242, 255, 0.2);
            padding-left: 0;
            padding-top: 15px;
        }
        .stat-grid { grid-template-columns: 1fr; }
        .sector-zone { min-width: unset; max-width: 100%; }
        .guardian-ring { gap: 15px; }
        .guardian-orb { width: 65px; }
        .guardian-orb-circle { width: 40px; height: 40px; }
    }
    </style>
</head>
<body>
    <div class="hud-container">
        <h1 class="hud-title"><svg class="mc-sprite" width="36" height="36" viewBox="0 0 8 8" shape-rendering="crispEdges"><rect x="0" y="0" width="8" height="8" fill="#5b3a1a"/><rect x="1" y="0" width="6" height="1" fill="#3b2210"/><rect x="0" y="1" width="1" height="2" fill="#3b2210"/><rect x="7" y="1" width="1" height="2" fill="#3b2210"/><rect x="1" y="1" width="6" height="2" fill="#c69c6d"/><rect x="0" y="3" width="8" height="1" fill="#c69c6d"/><rect x="0" y="4" width="1" height="1" fill="#c69c6d"/><rect x="7" y="4" width="1" height="1" fill="#c69c6d"/><rect x="1" y="4" width="2" height="1" fill="#fff"/><rect x="5" y="4" width="2" height="1" fill="#fff"/><rect x="2" y="4" width="1" height="1" fill="#1a0a2e"/><rect x="5" y="4" width="1" height="1" fill="#1a0a2e"/><rect x="3" y="4" width="2" height="1" fill="#c69c6d"/><rect x="0" y="5" width="8" height="1" fill="#c69c6d"/><rect x="3" y="5" width="2" height="1" fill="#a0724a"/><rect x="0" y="6" width="1" height="1" fill="#c69c6d"/><rect x="7" y="6" width="1" height="1" fill="#c69c6d"/><rect x="1" y="6" width="6" height="1" fill="#b5825a"/><rect x="2" y="6" width="4" height="1" fill="#c69c6d"/><rect x="0" y="7" width="8" height="1" fill="#c69c6d"/></svg><span>Henry's Sovereign HUD</span><svg class="mc-sprite" width="36" height="36" viewBox="0 0 8 8" shape-rendering="crispEdges"><rect x="0" y="0" width="8" height="8" fill="#0a0"/><rect x="1" y="1" width="2" height="2" fill="#000"/><rect x="5" y="1" width="2" height="2" fill="#000"/><rect x="3" y="3" width="2" height="2" fill="#000"/><rect x="2" y="4" width="1" height="1" fill="#000"/><rect x="5" y="4" width="1" height="1" fill="#000"/><rect x="2" y="5" width="4" height="1" fill="#000"/><rect x="0" y="0" width="8" height="1" fill="#060"/><rect x="0" y="0" width="1" height="8" fill="#060"/><rect x="7" y="0" width="1" height="8" fill="#060"/><rect x="0" y="7" width="8" height="1" fill="#060"/><rect x="1" y="1" width="2" height="2" fill="#111"/><rect x="5" y="1" width="2" height="2" fill="#111"/><rect x="3" y="3" width="2" height="2" fill="#111"/><rect x="2" y="4" width="1" height="1" fill="#111"/><rect x="5" y="4" width="1" height="1" fill="#111"/><rect x="2" y="5" width="4" height="1" fill="#111"/></svg></h1>

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
    <a href="/army" class="army-fab">[[ARMY_LINK]]</a>
    <a href="/quests" class="quest-fab">[[QUEST_LINK]]</a>
</body>
</html>`;

// ---------------------------------------------------------------------------
// Boss Detail Page Template
// ---------------------------------------------------------------------------
function buildBossPage(bossName, sector, minions, totals, activeQuestKeys) {
  const statusColor = { Enslaved: "#00ff9d", Engaged: "#ff6600", Locked: "#555" };
  const norm = (raw, max) => ((parseFloat(raw) || 0) / max * 100).toFixed(1);

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
app.use(express.urlencoded({ extended: false }));
app.use(express.json());

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

app.get("/", async (req, res) => {
  try {
    const sheets = await getSheets();
    const data = await fetchSheetData(sheets);

    // Quest data for badge + active indicators
    let activeQuestCount = 0;
    let activeQuestKeys = new Set();
    try {
      const quests = await fetchQuestsData(sheets);
      const activeQuests = quests.filter((q) => q["Status"] === "Active");
      activeQuestCount = activeQuests.length;
      for (const q of activeQuests) {
        activeQuestKeys.add(q["Boss"] + "|" + q["Minion"]);
      }
    } catch { /* Quests sheet may not exist yet */ }

    let html = processAllData(HTML_TEMPLATE, data, activeQuestKeys);

    const questBadgeHtml = activeQuestCount > 0
      ? `<a href="/quests" class="quest-badge-link"><span class="quest-badge">${activeQuestCount} ACTIVE</span></a>`
      : `<a href="/quests" class="quest-badge-link"><span class="quest-badge dim">QUEST BOARD</span></a>`;
    html = html.split("[[QUEST_BADGE]]").join(questBadgeHtml);

    const questLinkText = activeQuestCount > 0
      ? `&#x2694; ${activeQuestCount} QUEST${activeQuestCount > 1 ? "S" : ""}`
      : `&#x2694; QUESTS`;
    const fabClass = activeQuestCount > 0 ? "" : " dim";
    html = html.split("[[QUEST_LINK]]").join(questLinkText);
    html = html.split('class="quest-fab"').join('class="quest-fab' + fabClass + '"');

    // Army count (Enslaved minions)
    html = html.split("[[ARMY_LINK]]").join(`&#x2694; Henry's Army`);

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

    // Fetch active quests to mark already-queued minions
    let activeQuestKeys = new Set();
    try {
      const quests = await fetchQuestsData(sheets);
      for (const q of quests.filter((q) => q["Status"] === "Active")) {
        activeQuestKeys.add(q["Boss"] + "|" + q["Minion"]);
      }
    } catch {}

    res.send(buildBossPage(bossName, sector, minions, totals, activeQuestKeys));
  } catch (err) {
    console.error("Boss page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// Sector Detail Page — shows all bosses in a sector with their minion tables
// ---------------------------------------------------------------------------
function buildSectorPage(sectorName, bosses, totals, activeQuestKeys) {
  const statusColor = { Enslaved: "#00ff9d", Engaged: "#ff6600", Locked: "#555" };
  const norm = (raw, max) => ((parseFloat(raw) || 0) / max * 100).toFixed(1);

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

    bossBlocks += `
      <div class="boss-block">
        <h2><a href="/boss/${encodeURIComponent(bossName)}?sector=${encodeURIComponent(sectorName)}" class="boss-link">${escHtml(bossName)}</a></h2>
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
    .back-link { color: #00f2ff; text-decoration: none; font-size: 0.75em; letter-spacing: 2px; }
    .back-link:hover { text-decoration: underline; }
    h1 { color: #ff00ff; text-align: center; margin: 20px 0 5px 0; text-shadow: 0 0 15px rgba(255,0,255,0.4); letter-spacing: 4px; }
    .sector-tag { text-align: center; font-size: 0.7em; color: #888; letter-spacing: 3px; margin-bottom: 30px; }
    .boss-block { margin-bottom: 35px; }
    .boss-block h2 { color: #ff6600; font-size: 1em; letter-spacing: 3px; margin-bottom: 8px; text-shadow: 0 0 8px rgba(255,102,0,0.3); }
    .boss-link { color: #ff6600; text-decoration: none; }
    .boss-link:hover { text-decoration: underline; color: #ffea00; }
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

    res.send(buildSectorPage(sectorName, bosses, totals, activeQuestKeys));
  } catch (err) {
    console.error("Sector page error:", err);
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
    .quest-card[data-status="Rejected"] { border-color: rgba(255, 0, 68, 0.3); opacity: 0.5; }
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

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Quests!A:I",
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[questId, boss, minion, sector, "Active", proofType, "", taskDetail, ""]],
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
      newRows.push([questId, boss, minion, sector, "Active", proofType, "", taskDetail, ""]);
    }

    if (newRows.length > 0) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Quests!A:I",
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

app.post("/quest/remove", async (req, res) => {
  try {
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
    const bossCol = headers.indexOf("Boss");
    const minionCol = headers.indexOf("Minion");
    const sectorCol = headers.indexOf("Sector");

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

    // Clear Quest Status on Sectors before deleting the quest row
    const qBoss = questRow[bossCol];
    const qMinion = questRow[minionCol];
    const qSector = questRow[sectorCol];
    if (qBoss && qMinion && qSector) {
      await updateSectorsQuestStatus(sheets, qSector, qBoss, qMinion, "");
    }

    // Get sheet ID for Quests tab
    const meta = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
      fields: "sheets.properties",
    });
    const questSheet = meta.data.sheets.find((s) => s.properties.title === "Quests");
    if (!questSheet) return res.status(404).send("Quests sheet not found");

    // Delete the row
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId: questSheet.properties.sheetId,
              dimension: "ROWS",
              startIndex: targetRowIdx, // 0-based, header is row 0
              endIndex: targetRowIdx + 1,
            },
          },
        }],
      },
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

    const sorted = [...quests].sort((a, b) =>
      statusOrder.indexOf(a["Status"]) - statusOrder.indexOf(b["Status"])
    );

    let questRows = "";
    for (const q of sorted) {
      const sc = statusColors[q["Status"]] || "#555";
      const isActive = q["Status"] === "Active";
      const abandonBtn = isActive
        ? `<form method="POST" action="/quest/remove" style="margin:0;" onsubmit="return confirm('Abandon this quest?')">
             <input type="hidden" name="questId" value="${q["Quest ID"]}">
             <button type="submit" class="quest-abandon-btn" title="Abandon quest">X</button>
           </form>`
        : "";

      // Build artifact type dropdown for submission
      const currentType = q["Proof Type"] || "";
      let submitForm = "";
      if (isActive) {
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
            <span class="quest-status" style="color:${sc}">${q["Status"]}</span>
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

    const activeCount = quests.filter((q) => q["Status"] === "Active").length;
    res.send(buildQuestBoardPage(questRows, activeCount, quests.length));
  } catch (err) {
    console.error("Quest board error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

app.get("/army", async (req, res) => {
  try {
    const sheets = await getSheets();
    const batchRes = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SPREADSHEET_ID,
      ranges: ["Sectors", "Command_Center"],
    });
    const [sectorsRaw, ccRaw] = batchRes.data.valueRanges;
    const allMinions = parseTable(sectorsRaw.values);
    const commandCenter = parseTable(ccRaw.values);
    const enslaved = allMinions.filter((r) => r["Status"] === "Enslaved");

    const totals = {
      intel: getStat(commandCenter, "Intel").totalPossible,
      stamina: getStat(commandCenter, "Stamina").totalPossible,
      tempo: getStat(commandCenter, "Tempo").totalPossible,
      rep: getStat(commandCenter, "Reputation").totalPossible,
    };
    const norm = (raw, max) => ((parseFloat(raw) || 0) / max * 100).toFixed(1);

    // Group by sector
    const bySector = {};
    for (const m of enslaved) {
      const s = m["Sector"] || "Unknown";
      if (!bySector[s]) bySector[s] = [];
      bySector[s].push(m);
    }

    let sectionsHtml = "";
    for (const sector of Object.keys(bySector).sort()) {
      const minions = bySector[sector];
      let rows = "";
      for (const m of minions) {
        const nInt = norm(m["INTELLIGENCE"], totals.intel);
        const nSta = norm(m["STAMINA"], totals.stamina);
        const nTmp = norm(m["TEMPO"], totals.tempo);
        const nRep = norm(m["REPUTATION"], totals.rep);
        const nTotal = (parseFloat(nInt) + parseFloat(nSta) + parseFloat(nTmp) + parseFloat(nRep)).toFixed(1);
        rows += `<tr>
          <td>${escHtml(m["Boss"])}</td>
          <td>${escHtml(m["Minion"])}</td>
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
              <th>Boss</th><th>Minion</th>
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
        border-bottom: 1px solid #00ff9d;
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
    .empty-army {
        text-align: center;
        color: #555;
        padding: 40px;
        font-size: 0.9em;
    }
    .admin-fab {
        position: fixed;
        bottom: 20px;
        right: 20px;
        text-decoration: none;
        background: #0a0b10;
        border: 2px solid #ffea00;
        color: #ffea00;
        padding: 10px 18px;
        font-family: 'Courier New', monospace;
        font-size: 0.8em;
        font-weight: bold;
        letter-spacing: 2px;
        text-transform: uppercase;
        z-index: 1000;
        box-shadow: 0 0 15px rgba(255, 234, 0, 0.3), 0 0 30px rgba(255, 234, 0, 0.1);
        transition: all 0.3s;
    }
    .admin-fab:hover {
        background: #ffea00;
        color: #0a0b10;
        box-shadow: 0 0 25px rgba(255, 234, 0, 0.6);
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
        <div class="army-subtitle">${enslaved.length} MINIONS ENSLAVED</div>
        ${sectionsHtml || '<div class="empty-army">NO MINIONS ENSLAVED YET. CONQUER THEM THROUGH THE QUEST BOARD.</div>'}
    </div>
    <a href="/admin" class="admin-fab">&#x2699; PARENT ADMIN</a>
</body>
</html>`);
  } catch (err) {
    console.error("Army page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// Admin: Landing page
// ---------------------------------------------------------------------------
function buildAdminPage() {
  const functions = [
    { id: "import", title: "PHOTO IMPORT", desc: "Upload lesson photos for AI classification and auto-import to the tracker.", href: "/admin/import", active: true },
    { id: "quests", title: "QUEST APPROVAL", desc: "Review and approve completed quests submitted by Henry.", href: "/admin/quests", active: true },
    { id: "manual", title: "MANUAL ENTRY", desc: "Directly add or edit minions in the Sectors sheet.", href: "#", active: false },
    { id: "reports", title: "PROGRESS REPORTS", desc: "Generate weekly and monthly stat summaries.", href: "#", active: false },
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

app.get("/admin", (req, res) => {
  res.send(buildAdminPage());
});

// ---------------------------------------------------------------------------
// Admin: Quest Approval page
// ---------------------------------------------------------------------------
app.get("/admin/quests", async (req, res) => {
  try {
    const sheets = await getSheets();
    const quests = await fetchQuestsData(sheets);

    const statusOrder = { Submitted: 0, Active: 1, Rejected: 2, Approved: 3 };
    const statusColors = { Active: "#ff6600", Submitted: "#ffea00", Approved: "#00ff9d", Rejected: "#ff0044" };
    const sorted = [...quests].sort((a, b) =>
      (statusOrder[a["Status"]] ?? 9) - (statusOrder[b["Status"]] ?? 9)
    );

    let cards = "";
    const counts = { Submitted: 0, Active: 0, Rejected: 0, Approved: 0 };
    for (const q of sorted) {
      counts[q["Status"]] = (counts[q["Status"]] || 0) + 1;
      const sc = statusColors[q["Status"]] || "#555";
      const proofLink = q["Proof Link"] || "";
      const proofDisplay = proofLink
        ? (proofLink.startsWith("http")
          ? `<a href="${escHtml(proofLink)}" target="_blank" style="color:#00f2ff;">${escHtml(proofLink)}</a>`
          : escHtml(proofLink))
        : '<span style="color:#555;">No proof submitted</span>';

      let actions = "";
      if (q["Status"] === "Submitted") {
        actions = `
          <div class="qa-actions">
            <form method="POST" action="/admin/quests/approve" style="display:inline;">
              <input type="hidden" name="questId" value="${escHtml(q["Quest ID"])}">
              <button type="submit" class="qa-btn qa-approve" onclick="return confirm('Approve this quest? The minion will be marked as Enslaved.')">&#x2713; APPROVE</button>
            </form>
            <form method="POST" action="/admin/quests/reject" style="display:inline;">
              <input type="hidden" name="questId" value="${escHtml(q["Quest ID"])}">
              <button type="submit" class="qa-btn qa-reject" onclick="return confirm('Reject this quest? Henry will need to re-submit.')">&#x2717; REJECT</button>
            </form>
          </div>`;
      } else if (q["Status"] === "Rejected") {
        actions = `
          <div class="qa-actions">
            <form method="POST" action="/admin/quests/reopen" style="display:inline;">
              <input type="hidden" name="questId" value="${escHtml(q["Quest ID"])}">
              <button type="submit" class="qa-btn qa-reopen">&#x21BB; REOPEN</button>
            </form>
          </div>`;
      } else if (q["Status"] === "Approved") {
        actions = `<div class="qa-actions"><span class="qa-done">&#x2713; ENSLAVED</span></div>`;
      }

      cards += `
        <div class="qa-card" style="border-color: ${sc};">
          <div class="qa-header">
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
    .back-link { color: #00f2ff; text-decoration: none; font-size: 0.75em; letter-spacing: 2px; }
    .back-link:hover { text-decoration: underline; }
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
                <button type="submit" class="qa-btn qa-reopen" onclick="return confirm('Sync all quest statuses to the Sectors sheet?')">&#x21BB; SYNC TO SECTORS</button>
            </form>
        </div>
        ${cards || '<div class="qa-empty">NO QUESTS TO REVIEW</div>'}
    </div>
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

  const updates = [{
    range: "Quests!" + String.fromCharCode(65 + statusCol) + targetRowIdx,
    values: [[newStatus]],
  }];
  if (clearDate && dateCol >= 0) {
    updates.push({
      range: "Quests!" + String.fromCharCode(65 + dateCol) + targetRowIdx,
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

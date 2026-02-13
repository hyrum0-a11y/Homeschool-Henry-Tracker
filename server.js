const express = require("express");
const { google } = require("googleapis");
const Anthropic = require("@anthropic-ai/sdk");
const multer = require("multer");
const fs = require("fs");
const path = require("path");

// ---------------------------------------------------------------------------
// Load .env manually (no dotenv dependency needed)
// ---------------------------------------------------------------------------
try {
  const envPath = path.join(__dirname, ".env");
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
const CATALOG_SPREADSHEET_ID = process.env.CATALOG_SPREADSHEET_ID;
const CREDENTIALS_PATH = process.env.GOOGLE_CREDENTIALS_PATH || "./credentials.json";
const PORT = parseInt(process.env.PORT, 10) || 3000;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const IMPORT_DIR = path.join(__dirname, "imports");
const DONE_DIR = path.join(IMPORT_DIR, "done");

// ---------------------------------------------------------------------------
// Google Sheets auth (readwrite for fix-formulas support)
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
// Rule-based proof suggestion from sector stat weights
// ---------------------------------------------------------------------------
function generateProofSuggestion(sectorName, definitions) {
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
    return { proofType: "Document", suggestion: "Write a research summary about this minion." };
  }

  const stats = [
    { name: "intel", value: weights.intel },
    { name: "stamina", value: weights.stamina },
    { name: "tempo", value: weights.tempo },
    { name: "rep", value: weights.rep },
  ];
  stats.sort((a, b) => b.value - a.value);

  const suggestions = {
    intel: { proofType: "Document", suggestion: "Write a research report or explanation demonstrating deep understanding." },
    stamina: { proofType: "Spreadsheet", suggestion: "Create a practice log or data tracker showing consistent effort over time." },
    tempo: { proofType: "Presentation", suggestion: "Deliver a timed presentation or speed drill demonstrating quick mastery." },
    rep: { proofType: "Video", suggestion: "Record a video teaching this concept to someone else or presenting it publicly." },
  };

  return suggestions[stats[0].name] || suggestions.intel;
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

  // Next Confidence Tier — extract just the sub-rank numeral (e.g. "III")
  let nextConfSub = "MAX";
  for (let i = 0; i < definitions.length; i++) {
    const tierName = definitions[i]["Name"];
    if (tierName && tierName === currentConfRank) {
      if (i + 1 < definitions.length && definitions[i + 1]["Name"]) {
        const nextFull = definitions[i + 1]["Name"]; // e.g. "Copper III"
        const parts = nextFull.split(" ");
        nextConfSub = parts.length > 1 ? parts[parts.length - 1] : nextFull;
      }
      break;
    }
  }

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

    bossHtml += `<div class="sector-zone" data-sector="${sector.toUpperCase()}">${sectorRadarHtml}<div class="sector-bosses">${sectorBossHtml}</div></div>`;
  }

  // -- TOP 10 POWER RANKINGS (normalized like boss detail page) --
  const norm = (raw, max) => ((parseFloat(raw) || 0) / max) * 100;
  const minionScores = sectors
    .filter((r) => r["Boss"] && r["Minion"] && r["Status"] === "Engaged")
    .map((r) => {
      const nInt = norm(r["INTELLIGENCE"], intel.totalPossible);
      const nSta = norm(r["STAMINA"], stamina.totalPossible);
      const nTmp = norm(r["TEMPO"], tempo.totalPossible);
      const nRep = norm(r["REPUTATION"], reputation.totalPossible);
      return {
        minion: r["Minion"],
        boss: r["Boss"],
        sector: r["Sector"],
        status: r["Status"],
        nInt, nSta, nTmp, nRep,
        total: nInt + nSta + nTmp + nRep,
      };
    })
    .sort((a, b) => b.total - a.total)
    .slice(0, 10);

  const maxScore = minionScores[0]?.total || 1;
  let rankHtml = "";
  const rankColors = [
    "#ffd700", "#f0c800", "#e0bc00", "#d0b000", "#b8a400",
    "#a09800", "#88c8d0", "#70bcc8", "#58b0c0", "#40a4b8",
  ];
  minionScores.forEach((m, i) => {
    const pct = (m.total / maxScore) * 100;
    const color = rankColors[i] || "#40a4b8";
    const statusIcon = m.status === "Enslaved" ? "&#x2713;" : m.status === "Engaged" ? "&#x25CB;" : "&#x2717;";
    const qKey = m.boss + "|" + m.minion;
    const onQuest = activeQuestKeys.has(qKey);
    const questCell = onQuest
      ? `<span class="rank-quest-active" title="Already on quest board">&#x2605;</span>`
      : `<form class="rank-quest-form" method="POST" action="/quest/start">
          <input type="hidden" name="boss" value="${escHtml(m.boss)}">
          <input type="hidden" name="minion" value="${escHtml(m.minion)}">
          <input type="hidden" name="sector" value="${escHtml(m.sector)}">
          <button type="submit" class="rank-quest-btn" title="Start quest for ${escHtml(m.minion)}">+</button>
        </form>`;
    rankHtml += `
      <div class="rank-entry${onQuest ? " on-quest" : ""}">
        <span class="rank-badge" style="color:${color}">#${i + 1}</span>
        <span class="rank-boss">${m.boss}</span>
        <div class="rank-bar-wrap">
          <div class="rank-bar" style="width:${pct}%; background:${color};"></div>
          <span class="rank-info">${m.minion}</span>
        </div>
        <span class="rank-status">${statusIcon}</span>
        <span class="rank-stat" style="color:#00f2ff">${m.nInt.toFixed(1)}</span>
        <span class="rank-stat" style="color:#00ff9d">${m.nSta.toFixed(1)}</span>
        <span class="rank-stat" style="color:#ff00ff">${m.nTmp.toFixed(1)}</span>
        <span class="rank-stat" style="color:#ff8800">${m.nRep.toFixed(1)}</span>
        <span class="rank-pts">${m.total.toFixed(1)}</span>
        ${questCell}
      </div>`;
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
    .split("[[INTEL_BAR]]").join(tierBar(intel.value, intel.level, "Status pts"))
    .split("[[STAMINA_RANK]]").join(stamina.level)
    .split("[[STAMINA_REM]]").join(stamina.remaining.toFixed(1))
    .split("[[STAMINA_BAR]]").join(tierBar(stamina.value, stamina.level, "Status pts"))
    .split("[[TEMPO_RANK]]").join(tempo.level)
    .split("[[TEMPO_REM]]").join(tempo.remaining.toFixed(1))
    .split("[[TEMPO_BAR]]").join(tierBar(tempo.value, tempo.level, "Status pts"))
    .split("[[REP_RANK]]").join(reputation.level)
    .split("[[REPUTATION_REM]]").join(reputation.remaining.toFixed(1))
    .split("[[REPUTATION_BAR]]").join(tierBar(reputation.value, reputation.level, "Status pts"))
    .split("[[CONF_RANK]]").join(currentConfRank)
    .split("[[CONF_NEXT]]").join(nextConfSub)
    .split("[[CONF_REM]]").join(confidence.remaining.toFixed(1))
    .split("[[CONF_BAR]]").join(tierBar(confidence.value, currentConfRank, "Confidence pts"))
    .split("[[MAIN_RADAR]]").join(mainRadar)
    .split("[[TIER_LIST]]").join(tierListHtml)
    .split("[[BOSS_LIST]]").join(bossHtml)
    .split("[[POWER_RANKINGS]]").join(rankHtml);
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
        gap: 12px;
        padding: 10px 0;
    }

    .mc-sprite { filter: drop-shadow(0 0 4px #00f2ff); flex-shrink: 0; }

    /* Confidence row: bar + radar side by side */
    .confidence-row {
        display: flex;
        align-items: center;
        gap: 25px;
        margin-bottom: 25px;
        border-bottom: 1px dashed #ffea00;
        padding-bottom: 15px;
    }
    .confidence-bar-section { flex: 3; min-width: 0; }
    .radar-main { flex-shrink: 0; }

    .stat-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 20px;
        margin-bottom: 40px;
    }
    .stat-label { font-size: 0.85em; font-weight: bold; margin-bottom: 5px; }
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
        color: #888;
        font-size: 0.85em;
        width: 90px;
        flex-shrink: 0;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .rank-status {
        width: 18px;
        text-align: center;
        flex-shrink: 0;
    }
    .rank-stat {
        width: 32px;
        text-align: right;
        font-size: 0.85em;
        flex-shrink: 0;
    }
    .rank-pts {
        width: 35px;
        text-align: right;
        color: #ffea00;
        font-weight: bold;
        flex-shrink: 0;
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
    .rank-quest-form { flex-shrink: 0; margin: 0; padding: 0; }
    .rank-quest-btn {
        background: none;
        border: 1px solid #ff6600;
        color: #ff6600;
        width: 22px;
        height: 22px;
        font-size: 1em;
        font-weight: bold;
        cursor: pointer;
        font-family: 'Courier New', monospace;
        transition: all 0.2s;
        padding: 0;
        line-height: 20px;
    }
    .rank-quest-btn:hover { background: #ff6600; color: #0a0b10; box-shadow: 0 0 8px rgba(255, 102, 0, 0.5); }
    .rank-quest-active {
        width: 22px;
        text-align: center;
        color: #ffea00;
        font-size: 1em;
        flex-shrink: 0;
        text-shadow: 0 0 6px #ffea00;
    }
    .rank-entry.on-quest .rank-bar-wrap { border-color: rgba(255, 234, 0, 0.4); }

    @keyframes pulse { 0% { opacity: 1; } 50% { opacity: 0.7; } 100% { opacity: 1; } }
    .glitch { font-size: 0.7em; color: #555; margin-top: 20px; text-align: center; line-height: 1.6; }

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
    }
    </style>
</head>
<body>
    <div class="hud-container">
        <h1 class="hud-title"><svg class="mc-sprite" width="36" height="36" viewBox="0 0 8 8" shape-rendering="crispEdges"><rect x="0" y="0" width="8" height="8" fill="#5b3a1a"/><rect x="1" y="0" width="6" height="1" fill="#3b2210"/><rect x="0" y="1" width="1" height="2" fill="#3b2210"/><rect x="7" y="1" width="1" height="2" fill="#3b2210"/><rect x="1" y="1" width="6" height="2" fill="#c69c6d"/><rect x="0" y="3" width="8" height="1" fill="#c69c6d"/><rect x="0" y="4" width="1" height="1" fill="#c69c6d"/><rect x="7" y="4" width="1" height="1" fill="#c69c6d"/><rect x="1" y="4" width="2" height="1" fill="#fff"/><rect x="5" y="4" width="2" height="1" fill="#fff"/><rect x="2" y="4" width="1" height="1" fill="#1a0a2e"/><rect x="5" y="4" width="1" height="1" fill="#1a0a2e"/><rect x="3" y="4" width="2" height="1" fill="#c69c6d"/><rect x="0" y="5" width="8" height="1" fill="#c69c6d"/><rect x="3" y="5" width="2" height="1" fill="#a0724a"/><rect x="0" y="6" width="1" height="1" fill="#c69c6d"/><rect x="7" y="6" width="1" height="1" fill="#c69c6d"/><rect x="1" y="6" width="6" height="1" fill="#b5825a"/><rect x="2" y="6" width="4" height="1" fill="#c69c6d"/><rect x="0" y="7" width="8" height="1" fill="#c69c6d"/></svg><span>Henry's Sovereign HUD</span><svg class="mc-sprite" width="36" height="36" viewBox="0 0 8 8" shape-rendering="crispEdges"><rect x="0" y="0" width="8" height="8" fill="#0a0"/><rect x="1" y="1" width="2" height="2" fill="#000"/><rect x="5" y="1" width="2" height="2" fill="#000"/><rect x="3" y="3" width="2" height="2" fill="#000"/><rect x="2" y="4" width="1" height="1" fill="#000"/><rect x="5" y="4" width="1" height="1" fill="#000"/><rect x="2" y="5" width="4" height="1" fill="#000"/><rect x="0" y="0" width="8" height="1" fill="#060"/><rect x="0" y="0" width="1" height="8" fill="#060"/><rect x="7" y="0" width="1" height="8" fill="#060"/><rect x="0" y="7" width="8" height="1" fill="#060"/><rect x="1" y="1" width="2" height="2" fill="#111"/><rect x="5" y="1" width="2" height="2" fill="#111"/><rect x="3" y="3" width="2" height="2" fill="#111"/><rect x="2" y="4" width="1" height="1" fill="#111"/><rect x="5" y="4" width="1" height="1" fill="#111"/><rect x="2" y="5" width="4" height="1" fill="#111"/></svg></h1>

        <div class="confidence-row">
            <div class="radar-main">[[MAIN_RADAR]]</div>
            <div class="confidence-bar-section">
                <div class="stat-label" style="color: #ffea00; font-size: 1.1em;">
                    CONFIDENCE: [[CONF_RANK]] | +[[CONF_REM]] for [[CONF_NEXT]]
                </div>
                <div class="stat-bar">
                    <div class="fill confidence-fill" style="width: [[CONF_BAR]]%;"></div>
                </div>
            </div>
            <div class="levels-right">
                [[TIER_LIST]]
            </div>
        </div>

        <div class="stat-grid">
            <div class="stat-container">
                <div class="stat-label" style="color: #00f2ff;">INTEL: [[INTEL_RANK]] (+[[INTEL_REM]])</div>
                <div class="stat-bar"><div class="fill intel-fill" style="width: [[INTEL_BAR]]%;"></div></div>
            </div>
            <div class="stat-container">
                <div class="stat-label" style="color: #00ff9d;">STAMINA: [[STAMINA_RANK]] (+[[STAMINA_REM]])</div>
                <div class="stat-bar"><div class="fill stamina-fill" style="width: [[STAMINA_BAR]]%;"></div></div>
            </div>
            <div class="stat-container">
                <div class="stat-label" style="color: #ff00ff;">TEMPO: [[TEMPO_RANK]] (+[[TEMPO_REM]])</div>
                <div class="stat-bar"><div class="fill tempo-fill" style="width: [[TEMPO_BAR]]%;"></div></div>
            </div>
            <div class="stat-container">
                <div class="stat-label" style="color: #ff8800;">REPUTATION: [[REP_RANK]] (+[[REPUTATION_REM]])</div>
                <div class="stat-bar"><div class="fill rep-fill" style="width: [[REPUTATION_BAR]]%;"></div></div>
            </div>
        </div>

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
            <h1>Power Rankings [[QUEST_BADGE]]</h1>
            <div class="rank-entry rank-header">
                <span class="rank-badge"></span>
                <span class="rank-boss" style="color:#aaa;">BOSS</span>
                <div class="rank-bar-wrap" style="background:none;border:none;"></div>
                <span class="rank-status"></span>
                <span class="rank-stat" style="color:#00f2ff">INT</span>
                <span class="rank-stat" style="color:#00ff9d">STA</span>
                <span class="rank-stat" style="color:#ff00ff">TMP</span>
                <span class="rank-stat" style="color:#ff8800">REP</span>
                <span class="rank-pts">TOT</span>
                <span style="width:22px;flex-shrink:0;"></span>
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
      questBtn = `<form method="POST" action="/quest/start" style="margin:0;">
           <input type="hidden" name="boss" value="${escHtml(bossName)}">
           <input type="hidden" name="minion" value="${escHtml(m["Minion"])}">
           <input type="hidden" name="sector" value="${escHtml(sector)}">
           <button type="submit" class="quest-add-btn" title="Start quest">+</button>
         </form>`;
    } else {
      questBtn = `<span style="opacity:0.2;">-</span>`;
    }
    rows += `
      <tr>
        <td>${questBtn}</td>
        <td>${m["Minion"]}</td>
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
        <div style="text-align:center;margin-top:15px;font-size:0.75em;color:#ff6600;">CLICK + TO ADD AN ENGAGED MINION TO YOUR QUEST BOARD</div>
    </div>
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
    Status: "Enslaved",
    "Impact(1-3)": data.impact,
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
        color: #aaa;
        border-left: 2px solid #ff00ff;
        padding-left: 10px;
        margin-top: 8px;
        text-transform: none;
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
    .quest-submit-btn:hover { background: #00ff9d; color: #0a0b10; }
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

    const defRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Definitions",
    });
    const definitions = parseTable(defRes.data.values);

    const { proofType, suggestion } = generateProofSuggestion(sector, definitions);
    const questId = generateQuestId();

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Quests!A:I",
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[questId, boss, minion, sector, "Active", proofType, "", suggestion, ""]],
      },
    });

    res.redirect("/quests");
  } catch (err) {
    console.error("Quest start error:", err);
    res.status(500).send(`<pre style="color:red">Error starting quest: ${err.message}</pre>`);
  }
});

app.post("/quest/submit", async (req, res) => {
  try {
    const { questId, proofLink } = req.body;
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
    const proofLinkCol = headerRow.indexOf("Proof Link");
    const dateCol = headerRow.indexOf("Date Completed");

    let targetRowIdx = -1;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][idCol] === questId) {
        targetRowIdx = i + 1; // 1-based for Sheets API
        break;
      }
    }

    if (targetRowIdx === -1) return res.status(404).send("Quest not found");

    const updates = [];
    updates.push({
      range: "Quests!" + String.fromCharCode(65 + statusCol) + targetRowIdx,
      values: [["Submitted"]],
    });
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

    const idCol = rows[0].indexOf("Quest ID");
    let targetRowIdx = -1;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][idCol] === questId) {
        targetRowIdx = i; // 0-based data index
        break;
      }
    }
    if (targetRowIdx === -1) return res.status(404).send("Quest not found");

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
    const quests = await fetchQuestsData(sheets);

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
      const submitForm = isActive
        ? `<form class="quest-submit-form" method="POST" action="/quest/submit">
             <input type="hidden" name="questId" value="${q["Quest ID"]}">
             <input type="text" name="proofLink" placeholder="PASTE PROOF LINK..." class="proof-input">
             <button type="submit" class="quest-submit-btn">SUBMIT</button>
           </form>`
        : `<span class="quest-proof-link">${q["Proof Link"] || "---"}</span>`;

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
              <span class="quest-proof-type">[${q["Proof Type"]}]</span>
              ${q["Suggested By AI"]}
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
    { id: "catalog", title: "OBJECTIVE CATALOG", desc: "Browse subjects, assign objectives to the student, or populate the catalog with AI.", href: "/admin/catalog", active: true },
    { id: "quests", title: "QUEST APPROVAL", desc: "Review and approve completed quests submitted by Henry.", href: "#", active: false },
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

// ===========================================================================
// CATALOG ROUTES — Parent-facing UI for browsing & managing objectives
// ===========================================================================

// ---------------------------------------------------------------------------
// Catalog helpers
// ---------------------------------------------------------------------------
async function fetchCatalogData(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: CATALOG_SPREADSHEET_ID,
    range: "Catalog",
  });
  const values = res.data.values || [];
  return { headers: values[0] || [], rows: parseTable(values) };
}

async function fetchStudentSectors(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sectors",
  });
  const values = res.data.values || [];
  return { headers: values[0] || [], rows: parseTable(values) };
}

function buildSectorsRowFromCatalog(studentHeaders, item, status) {
  const statNames = ["INTELLIGENCE", "STAMINA", "TEMPO", "REPUTATION"];
  const statFormula = (stat) =>
    `=INDEX($A:$Z,ROW(),MATCH("Impact(1-3)",$1:$1,0))*INDEX(Definitions!$A:$Z,MATCH(INDEX($A:$Z,ROW(),MATCH("Sector",$1:$1,0)),INDEX(Definitions!$A:$Z,,MATCH("Sector",Definitions!$1:$1,0)),0),MATCH("${stat}",Definitions!$1:$1,0))`;
  const validStatuses = ["Locked", "Engaged", "Enslaved"];
  const finalStatus = validStatuses.includes(status) ? status : "Locked";
  const valueMap = {
    Sector: item["Sector"],
    Boss: item["Boss"],
    Minion: item["Minion"],
    Status: finalStatus,
    "Impact(1-3)": item["Impact(1-3)"],
    "Locked for what?": item["Locked for what?"] || "",
  };
  for (const stat of statNames) valueMap[stat] = statFormula(stat);
  return studentHeaders.map((h) => valueMap[h] ?? "");
}

// ---------------------------------------------------------------------------
// Expand catalog table filter + banding to include all rows
// ---------------------------------------------------------------------------
async function expandCatalogTable(sheets) {
  // Get current data size
  const dataRes = await sheets.spreadsheets.values.get({
    spreadsheetId: CATALOG_SPREADSHEET_ID,
    range: "Catalog",
  });
  const totalRows = (dataRes.data.values || []).length;
  const totalCols = (dataRes.data.values || [[]])[0].length;

  // Get sheet metadata
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: CATALOG_SPREADSHEET_ID,
    fields: "sheets.properties,sheets.basicFilter,sheets.bandedRanges",
  });
  const sheet = meta.data.sheets[0];
  const sheetId = sheet.properties.sheetId;

  const requests = [];

  // Clear and re-set filter to cover all rows
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

  // Update banding to cover all rows
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

  if (requests.length > 0) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: CATALOG_SPREADSHEET_ID,
      requestBody: { requests },
    });
  }
}

// ---------------------------------------------------------------------------
// Shared CSS for catalog pages
// ---------------------------------------------------------------------------
const CATALOG_CSS = `
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
        max-width: 1100px;
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
    .subtitle {
        text-align: center;
        color: #888;
        font-size: 0.8em;
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
    .nav-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 15px;
    }
    .hud-link {
        display: inline-block;
        color: #ffea00;
        text-decoration: none;
        border: 1px solid #ffea00;
        padding: 6px 15px;
        font-size: 0.8em;
        transition: all 0.2s;
    }
    .hud-link:hover { background: #ffea00; color: #0a0b10; }
    .btn {
        display: inline-block;
        background: #0a0b10;
        border: 2px solid #ffea00;
        color: #ffea00;
        padding: 10px 20px;
        font-family: 'Courier New', monospace;
        font-size: 0.85em;
        text-transform: uppercase;
        letter-spacing: 2px;
        cursor: pointer;
        transition: all 0.2s;
        text-decoration: none;
    }
    .btn:hover { background: #ffea00; color: #0a0b10; }
    .btn:disabled { opacity: 0.3; cursor: not-allowed; }
    .btn:disabled:hover { background: #0a0b10; color: #ffea00; }
    .btn-cyan { border-color: #00f2ff; color: #00f2ff; }
    .btn-cyan:hover { background: #00f2ff; color: #0a0b10; }
    .btn-magenta { border-color: #ff00ff; color: #ff00ff; }
    .btn-magenta:hover { background: #ff00ff; color: #0a0b10; }
    .success-msg {
        background: rgba(0, 255, 157, 0.1);
        border: 1px solid #00ff9d;
        color: #00ff9d;
        padding: 10px 15px;
        margin-bottom: 20px;
        text-align: center;
        letter-spacing: 1px;
        font-size: 0.85em;
    }
    @media (max-width: 600px) {
        body { padding: 10px; }
        .hud-container { padding: 15px; }
    }
`;

// ---------------------------------------------------------------------------
// GET /admin/catalog — Subject picker page
// ---------------------------------------------------------------------------
app.get("/admin/catalog", async (req, res) => {
  try {
    const sheets = await getSheets();
    const [catalog, student] = await Promise.all([
      fetchCatalogData(sheets),
      fetchStudentSectors(sheets),
    ]);

    // Build student key set for "in sheet" detection
    const studentKeys = new Set();
    for (const row of student.rows) {
      if (row["Sector"] && row["Minion"]) {
        studentKeys.add(`${row["Sector"]}|${row["Boss"]}|${row["Minion"]}`);
      }
    }

    // Group catalog by Subject
    const subjectMap = {};
    for (const row of catalog.rows) {
      const subject = row["Subject"] || row["Sector"] || "General";
      const sector = row["Sector"] || "";
      if (!subjectMap[subject]) subjectMap[subject] = { sector, total: 0, inSheet: 0 };
      subjectMap[subject].total++;
      const key = `${row["Sector"]}|${row["Boss"]}|${row["Minion"]}`;
      if (studentKeys.has(key)) subjectMap[subject].inSheet++;
    }

    const sortedSubjects = Object.keys(subjectMap).sort();

    const cards = sortedSubjects.map((subject) => {
      const info = subjectMap[subject];
      const pct = info.total > 0 ? Math.round((info.inSheet / info.total) * 100) : 0;
      const encodedSubject = encodeURIComponent(subject);
      return `<label class="subject-card">
        <input type="checkbox" name="subjects" value="${escHtml(subject)}" class="subject-chk">
        <div class="subject-name">${escHtml(subject)}</div>
        <div class="subject-stats">${info.inSheet} / ${info.total} assigned</div>
        <div class="subject-bar"><div class="subject-fill" style="width:${pct}%"></div></div>
      </label>`;
    }).join("\n");

    res.send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Objective Catalog - Sovereign HUD</title>
    <style>
    ${CATALOG_CSS}
    .subject-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
        gap: 15px;
        margin-bottom: 30px;
    }
    .subject-card {
        border: 2px solid #ff00ff;
        border-radius: 6px;
        padding: 18px;
        color: inherit;
        transition: all 0.3s;
        display: block;
        cursor: pointer;
    }
    .subject-card:hover {
        background: rgba(255, 0, 255, 0.08);
        box-shadow: 0 0 20px rgba(255, 0, 255, 0.4);
    }
    .subject-card.selected {
        border-color: #ffea00;
        background: rgba(255, 234, 0, 0.08);
        box-shadow: 0 0 15px rgba(255, 234, 0, 0.3);
    }
    .subject-card.selected .subject-name { color: #ffea00; }
    .subject-chk { display: none; }
    .subject-name {
        font-size: 1em;
        font-weight: bold;
        color: #ff00ff;
        margin-bottom: 8px;
        letter-spacing: 2px;
    }
    .subject-stats {
        font-size: 0.75em;
        color: #888;
        margin-bottom: 8px;
    }
    .subject-bar {
        background: #1a1d26;
        border: 1px solid #333;
        height: 8px;
        overflow: hidden;
    }
    .subject-fill {
        height: 100%;
        background: linear-gradient(90deg, #ff00ff, #ff88ff);
        box-shadow: 0 0 6px #ff00ff;
        transition: width 0.3s;
    }
    .action-row {
        display: flex;
        gap: 15px;
        flex-wrap: wrap;
        margin-bottom: 25px;
    }
    </style>
</head>
<body>
    <div class="hud-container">
        <div class="nav-row">
            <a class="back-link" href="/admin">&lt; BACK TO ADMIN</a>
            <a class="hud-link" href="/">HUD &gt;</a>
        </div>
        <h1>Objective Catalog</h1>
        <div class="subtitle">SELECT ONE OR MORE SUBJECTS TO BROWSE OBJECTIVES</div>
        <div class="action-row">
            <a href="/admin/catalog/student" class="btn btn-cyan">VIEW CURRENT STUDENT ITEMS</a>
            <a href="/admin/catalog/locked" class="btn btn-cyan">MANAGE LOCKED ITEMS</a>
            <a href="/admin/catalog/generate" class="btn btn-magenta">POPULATE CATALOG WITH AI</a>
            <button class="btn" id="viewSelectedBtn" disabled>VIEW SELECTED SUBJECTS</button>
        </div>
        <div class="subject-grid">
            ${cards}
        </div>
    </div>
    <script>
    (function() {
        const cards = document.querySelectorAll('.subject-card');
        const btn = document.getElementById('viewSelectedBtn');

        cards.forEach(function(card) {
            card.addEventListener('click', function() {
                const chk = card.querySelector('.subject-chk');
                chk.checked = !chk.checked;
                card.classList.toggle('selected', chk.checked);
                const count = document.querySelectorAll('.subject-chk:checked').length;
                btn.disabled = count === 0;
                btn.textContent = count > 1 ? 'VIEW ' + count + ' SUBJECTS' : 'VIEW SELECTED SUBJECT';
            });
        });

        btn.addEventListener('click', function() {
            const selected = [];
            document.querySelectorAll('.subject-chk:checked').forEach(function(c) {
                selected.push(encodeURIComponent(c.value));
            });
            if (selected.length > 0) {
                window.location.href = '/admin/catalog/view?subjects=' + selected.join(',');
            }
        });
    })();
    </script>
</body>
</html>`);
  } catch (err) {
    console.error("Catalog page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// GET /admin/catalog/student — View current student items
// ---------------------------------------------------------------------------
app.get("/admin/catalog/student", async (req, res) => {
  try {
    const sheets = await getSheets();
    const student = await fetchStudentSectors(sheets);

    // Also fetch catalog to get Subject names
    const catalog = await fetchCatalogData(sheets);
    const subjectLookup = {};
    for (const row of catalog.rows) {
      const key = `${row["Sector"]}|${row["Boss"]}`;
      if (row["Subject"]) subjectLookup[key] = row["Subject"];
    }

    // Group by sector
    const bySector = {};
    for (const row of student.rows) {
      const sector = row["Sector"] || "Unknown";
      if (!bySector[sector]) bySector[sector] = [];
      bySector[sector].push(row);
    }

    const statusColor = { Enslaved: "#00ff9d", Engaged: "#ff8800", Locked: "#666" };

    let sectionsHtml = "";
    for (const sector of Object.keys(bySector).sort()) {
      const rows = bySector[sector];
      const subject = subjectLookup[`${sector}|${rows[0]?.Boss}`] || sector;
      const counts = { Enslaved: 0, Engaged: 0, Locked: 0 };
      for (const r of rows) counts[r["Status"]] = (counts[r["Status"]] || 0) + 1;

      let rowsHtml = rows.map((r) => {
        const color = statusColor[r["Status"]] || "#666";
        return `<tr>
          <td style="color:${color}">${escHtml(r["Status"])}</td>
          <td>${escHtml(r["Boss"])}</td>
          <td>${escHtml(r["Minion"])}</td>
          <td style="text-align:center">${escHtml(r["Impact(1-3)"])}</td>
        </tr>`;
      }).join("");

      sectionsHtml += `
        <div class="student-sector">
          <div class="sector-header">
            <span class="sector-title">${escHtml(subject)} <span style="color:#666;font-size:0.7em">(${escHtml(sector)})</span></span>
            <span class="sector-counts">
              <span style="color:#00ff9d">${counts.Enslaved} completed</span> /
              <span style="color:#ff8800">${counts.Engaged} available</span> /
              <span style="color:#666">${counts.Locked} locked</span>
            </span>
          </div>
          <table class="obj-table">
            <thead><tr><th>Status</th><th>Boss</th><th>Minion</th><th>Impact</th></tr></thead>
            <tbody>${rowsHtml}</tbody>
          </table>
        </div>`;
    }

    res.send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Student Items - Sovereign HUD</title>
    <style>
    ${CATALOG_CSS}
    .student-sector {
        margin-bottom: 25px;
        border: 1px solid rgba(0, 242, 255, 0.3);
        border-radius: 6px;
        overflow: hidden;
    }
    .sector-header {
        background: rgba(0, 242, 255, 0.05);
        padding: 12px 15px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 8px;
    }
    .sector-title { font-weight: bold; color: #00f2ff; letter-spacing: 2px; }
    .sector-counts { font-size: 0.75em; color: #888; }
    .obj-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.8em;
    }
    .obj-table th {
        text-align: left;
        padding: 8px 12px;
        border-bottom: 1px solid #333;
        color: #888;
        font-size: 0.85em;
    }
    .obj-table td {
        padding: 6px 12px;
        border-bottom: 1px solid #1a1d26;
    }
    .obj-table tr:hover { background: rgba(255, 234, 0, 0.03); }
    .total-bar {
        text-align: center;
        padding: 15px;
        font-size: 0.9em;
        color: #888;
        letter-spacing: 2px;
        margin-bottom: 20px;
    }
    .total-bar span { color: #ffea00; font-weight: bold; }
    </style>
</head>
<body>
    <div class="hud-container">
        <div class="nav-row">
            <a class="back-link" href="/admin/catalog">&lt; BACK TO CATALOG</a>
            <a class="hud-link" href="/">HUD &gt;</a>
        </div>
        <h1>Current Student Items</h1>
        <div class="total-bar"><span>${student.rows.length}</span> total objectives assigned</div>
        ${sectionsHtml || '<div class="subtitle">NO ITEMS ASSIGNED YET</div>'}
    </div>
</body>
</html>`);
  } catch (err) {
    console.error("Student items page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// GET /admin/catalog/view?subjects=... — Objectives list filtered by subject
// ---------------------------------------------------------------------------
app.get("/admin/catalog/view", async (req, res) => {
  try {
    const subjectsParam = req.query.subjects || "";
    const selectedSubjects = subjectsParam.split(",").map(s => decodeURIComponent(s)).filter(Boolean);
    if (selectedSubjects.length === 0) return res.redirect("/admin/catalog");

    const added = req.query.added ? parseInt(req.query.added, 10) : 0;

    const sheets = await getSheets();
    const [catalog, student] = await Promise.all([
      fetchCatalogData(sheets),
      fetchStudentSectors(sheets),
    ]);

    // Build student key set
    const studentKeys = new Set();
    for (const row of student.rows) {
      if (row["Sector"] && row["Minion"]) {
        studentKeys.add(`${row["Sector"]}|${row["Boss"]}|${row["Minion"]}`);
      }
    }

    // Filter catalog by selected subjects
    const subjectSet = new Set(selectedSubjects);
    const items = catalog.rows.filter((r) => subjectSet.has(r["Subject"] || r["Sector"]));

    const titleText = selectedSubjects.length === 1
      ? escHtml(selectedSubjects[0])
      : selectedSubjects.length + " Subjects";

    // Split into available and assigned
    const available = [];
    const assigned = [];
    for (const item of items) {
      const key = `${item["Sector"]}|${item["Boss"]}|${item["Minion"]}`;
      if (studentKeys.has(key)) {
        assigned.push(item);
      } else {
        available.push(item);
      }
    }

    // Build available table rows (with checkboxes and status picker)
    let availableRows = "";
    available.forEach((item, idx) => {
      const value = JSON.stringify({
        Sector: item["Sector"],
        Boss: item["Boss"],
        Minion: item["Minion"],
        "Impact(1-3)": item["Impact(1-3)"],
        "Locked for what?": item["Locked for what?"] || "",
      }).replace(/"/g, "&quot;");
      const subjectLabel = selectedSubjects.length > 1 ? `<td>${escHtml(item["Subject"] || item["Sector"])}</td>` : "";
      availableRows += `<tr>
        <td class="chk-cell"><input type="checkbox" name="items" value="${value}" class="item-chk" data-idx="${idx}"></td>
        ${subjectLabel}
        <td>${escHtml(item["Boss"])}</td>
        <td>${escHtml(item["Minion"])}</td>
        <td style="text-align:center">${escHtml(item["Impact(1-3)"])}</td>
        <td class="status-cell">
            <select name="status_${idx}" class="status-select" disabled>
                <option value="Engaged" selected>Engaged (available)</option>
                <option value="Locked">Locked (has prerequisite)</option>
            </select>
        </td>
      </tr>`;
    });

    // Build assigned table rows (no checkboxes, just display)
    let assignedRows = "";
    for (const item of assigned) {
      const statusColor = ({ Enslaved: "#00ff9d", Engaged: "#ff8800", Locked: "#666" })[item["Status"]] || "#666";
      const subjectLabel = selectedSubjects.length > 1 ? `<td>${escHtml(item["Subject"] || item["Sector"])}</td>` : "";
      assignedRows += `<tr>
        <td style="text-align:center"><span class="in-sheet-mark">&#x2714;</span></td>
        ${subjectLabel}
        <td>${escHtml(item["Boss"])}</td>
        <td>${escHtml(item["Minion"])}</td>
        <td style="text-align:center;color:${statusColor}">${escHtml(item["Status"])}</td>
        <td style="text-align:center">${escHtml(item["Impact(1-3)"])}</td>
      </tr>`;
    }

    const subjectTh = selectedSubjects.length > 1 ? "<th>Subject</th>" : "";
    const successHtml = added > 0
      ? `<div class="success-msg">${added} objective${added > 1 ? "s" : ""} added to student sheet</div>`
      : "";
    const subjectsQs = selectedSubjects.map(s => encodeURIComponent(s)).join(",");

    res.send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>${titleText} - Objective Catalog</title>
    <style>
    ${CATALOG_CSS}
    .section-label {
        font-size: 0.9em;
        letter-spacing: 2px;
        margin: 25px 0 10px 0;
        padding-bottom: 5px;
        border-bottom: 1px dashed #333;
    }
    .section-label.available { color: #ffea00; }
    .section-label.assigned { color: #00ff9d; }
    .obj-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.8em;
        margin-bottom: 20px;
    }
    .obj-table th {
        text-align: left;
        padding: 10px 12px;
        border-bottom: 2px solid #ffea00;
        color: #ffea00;
        font-size: 0.85em;
        letter-spacing: 1px;
    }
    .obj-table td {
        padding: 8px 12px;
        border-bottom: 1px solid #1a1d26;
    }
    .obj-table tr:hover { background: rgba(255, 234, 0, 0.03); }
    .in-sheet-mark { color: #00ff9d; font-size: 1.1em; }
    .chk-cell { width: 35px; text-align: center; }
    .item-chk { accent-color: #ffea00; width: 16px; height: 16px; cursor: pointer; }
    .action-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 15px;
        flex-wrap: wrap;
        gap: 10px;
    }
    .select-count { font-size: 0.8em; color: #888; letter-spacing: 1px; }
    .assigned-table th { border-bottom-color: #00ff9d; color: #00ff9d; }
    .empty-note { color: #555; font-size: 0.8em; padding: 15px 0; letter-spacing: 1px; }
    .status-select {
        background: #1a1d26;
        border: 1px solid #333;
        color: #888;
        padding: 3px 6px;
        font-family: 'Courier New', monospace;
        font-size: 0.9em;
        text-transform: uppercase;
    }
    .status-select:not(:disabled) { color: #ffea00; border-color: #ffea00; }
    .status-cell { text-align: center; }
    </style>
</head>
<body>
    <div class="hud-container">
        <div class="nav-row">
            <a class="back-link" href="/admin/catalog">&lt; BACK TO SUBJECTS</a>
            <a class="hud-link" href="/">HUD &gt;</a>
        </div>
        <h1>${titleText}</h1>
        <div class="subtitle">${available.length} AVAILABLE &mdash; ${assigned.length} ALREADY ASSIGNED</div>
        ${successHtml}

        ${available.length > 0 ? `
        <div class="section-label available">AVAILABLE TO ADD (${available.length})</div>
        <form method="POST" action="/admin/catalog/add" id="addForm">
            <input type="hidden" name="returnSubjects" value="${escHtml(subjectsQs)}">
            <div class="action-bar">
                <div>
                    <button type="button" class="btn" id="toggleAll" style="font-size:0.75em;padding:6px 12px">SELECT ALL</button>
                    <span class="select-count" id="selectCount">0 selected</span>
                </div>
                <button type="submit" class="btn" id="submitBtn" disabled>ADD SELECTED TO STUDENT SHEET</button>
            </div>
            <table class="obj-table">
                <thead><tr><th></th>${subjectTh}<th>Boss (Topic)</th><th>Minion (Objective)</th><th>Impact</th><th>Add As</th></tr></thead>
                <tbody>${availableRows}</tbody>
            </table>
        </form>
        ` : '<div class="empty-note">ALL CATALOG OBJECTIVES FOR THIS SUBJECT ARE ALREADY IN THE STUDENT SHEET</div>'}

        ${assigned.length > 0 ? `
        <div class="section-label assigned">ALREADY IN STUDENT SHEET (${assigned.length})</div>
        <table class="obj-table assigned-table">
            <thead><tr><th></th>${subjectTh}<th>Boss (Topic)</th><th>Minion (Objective)</th><th>Status</th><th>Impact</th></tr></thead>
            <tbody>${assignedRows}</tbody>
        </table>
        ` : ''}
    </div>
    <script>
    (function() {
        const checks = document.querySelectorAll('.item-chk');
        if (checks.length === 0) return;
        const countEl = document.getElementById('selectCount');
        const submitBtn = document.getElementById('submitBtn');
        const toggleBtn = document.getElementById('toggleAll');
        let allSelected = false;

        function updateCount() {
            const n = document.querySelectorAll('.item-chk:checked').length;
            countEl.textContent = n + ' selected';
            submitBtn.disabled = n === 0;
        }

        function toggleStatus(chk) {
            const row = chk.closest('tr');
            const sel = row.querySelector('.status-select');
            if (sel) sel.disabled = !chk.checked;
        }

        checks.forEach(function(c) {
            c.addEventListener('change', function() {
                toggleStatus(c);
                updateCount();
            });
        });

        toggleBtn.addEventListener('click', function() {
            allSelected = !allSelected;
            checks.forEach(function(c) {
                c.checked = allSelected;
                toggleStatus(c);
            });
            toggleBtn.textContent = allSelected ? 'DESELECT ALL' : 'SELECT ALL';
            updateCount();
        });
    })();
    </script>
</body>
</html>`);
  } catch (err) {
    console.error("Catalog view page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// GET /admin/catalog/locked — Manage locked items
// ---------------------------------------------------------------------------
app.get("/admin/catalog/locked", async (req, res) => {
  try {
    const updated = req.query.updated ? parseInt(req.query.updated, 10) : 0;
    const sheets = await getSheets();
    const student = await fetchStudentSectors(sheets);
    const catalog = await fetchCatalogData(sheets);

    // Subject lookup from catalog
    const subjectLookup = {};
    for (const row of catalog.rows) {
      const key = `${row["Sector"]}|${row["Boss"]}`;
      if (row["Subject"]) subjectLookup[key] = row["Subject"];
    }

    // Get all locked items from student sheet
    const lockedItems = [];
    const allMinions = []; // for prerequisite picker
    const headers = student.headers;
    const sectorIdx = headers.indexOf("Sector");
    const bossIdx = headers.indexOf("Boss");
    const minionIdx = headers.indexOf("Minion");
    const statusIdx = headers.indexOf("Status");
    const lockedForIdx = headers.indexOf("Locked for what?");

    const rawRows = student.rows;
    for (let i = 0; i < rawRows.length; i++) {
      const row = rawRows[i];
      if (row["Boss"] && row["Minion"]) {
        allMinions.push({
          sector: row["Sector"],
          boss: row["Boss"],
          minion: row["Minion"],
          status: row["Status"],
        });
      }
      if (row["Status"] === "Locked") {
        lockedItems.push({
          rowNum: i + 2, // 1-indexed, +1 for header
          sector: row["Sector"],
          boss: row["Boss"],
          minion: row["Minion"],
          lockedFor: row["Locked for what?"] || "",
          subject: subjectLookup[`${row["Sector"]}|${row["Boss"]}`] || row["Sector"],
        });
      }
    }

    // Build prerequisite options (all non-locked minions)
    const prereqOptions = allMinions
      .filter(m => m.status !== "Locked")
      .map(m => `<option value="${escHtml(m.minion)}">${escHtml(m.boss)} - ${escHtml(m.minion)} (${escHtml(m.status)})</option>`)
      .join("");

    // Build locked items table
    let tableRows = "";
    for (const item of lockedItems) {
      tableRows += `<tr>
        <td>
            <input type="checkbox" name="rows" value="${item.rowNum}" class="locked-chk" data-row="${item.rowNum}">
        </td>
        <td>${escHtml(item.subject)}</td>
        <td>${escHtml(item.boss)}</td>
        <td>${escHtml(item.minion)}</td>
        <td>
            <select name="prereq_${item.rowNum}" class="prereq-select">
                <option value="${escHtml(item.lockedFor)}">${item.lockedFor ? escHtml(item.lockedFor) : '-- none --'}</option>
                ${prereqOptions}
            </select>
        </td>
      </tr>`;
    }

    const successHtml = updated > 0
      ? `<div class="success-msg">${updated} item${updated > 1 ? "s" : ""} updated</div>`
      : "";

    res.send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Locked Items - Sovereign HUD</title>
    <style>
    ${CATALOG_CSS}
    .obj-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.8em;
        margin-bottom: 20px;
    }
    .obj-table th {
        text-align: left;
        padding: 10px 12px;
        border-bottom: 2px solid #ffea00;
        color: #ffea00;
        font-size: 0.85em;
        letter-spacing: 1px;
    }
    .obj-table td {
        padding: 8px 12px;
        border-bottom: 1px solid #1a1d26;
    }
    .obj-table tr:hover { background: rgba(255, 234, 0, 0.03); }
    .locked-chk { accent-color: #ffea00; width: 16px; height: 16px; cursor: pointer; }
    .prereq-select {
        background: #1a1d26;
        border: 1px solid #333;
        color: #888;
        padding: 4px 6px;
        font-family: 'Courier New', monospace;
        font-size: 0.9em;
        max-width: 280px;
        text-transform: none;
    }
    .action-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 15px;
        flex-wrap: wrap;
        gap: 10px;
    }
    .select-count { font-size: 0.8em; color: #888; letter-spacing: 1px; }
    .action-buttons { display: flex; gap: 10px; flex-wrap: wrap; }
    </style>
</head>
<body>
    <div class="hud-container">
        <div class="nav-row">
            <a class="back-link" href="/admin/catalog">&lt; BACK TO CATALOG</a>
            <a class="hud-link" href="/">HUD &gt;</a>
        </div>
        <h1>Locked Items</h1>
        <div class="subtitle">${lockedItems.length} LOCKED OBJECTIVES &mdash; SELECT TO UNLOCK OR SET PREREQUISITES</div>
        ${successHtml}
        ${lockedItems.length > 0 ? `
        <form method="POST" action="/admin/catalog/locked/update" id="lockedForm">
            <div class="action-bar">
                <div>
                    <button type="button" class="btn" id="toggleAll" style="font-size:0.75em;padding:6px 12px">SELECT ALL</button>
                    <span class="select-count" id="selectCount">0 selected</span>
                </div>
                <div class="action-buttons">
                    <button type="submit" name="action" value="unlock" class="btn" id="unlockBtn" disabled>UNLOCK SELECTED (SET ENGAGED)</button>
                    <button type="submit" name="action" value="prereq" class="btn btn-cyan" id="prereqBtn" disabled>UPDATE PREREQUISITES</button>
                </div>
            </div>
            <table class="obj-table">
                <thead>
                    <tr><th></th><th>Subject</th><th>Boss</th><th>Minion</th><th>Prerequisite</th></tr>
                </thead>
                <tbody>${tableRows}</tbody>
            </table>
        </form>
        ` : '<div class="subtitle">NO LOCKED ITEMS</div>'}
    </div>
    <script>
    (function() {
        const checks = document.querySelectorAll('.locked-chk');
        if (checks.length === 0) return;
        const countEl = document.getElementById('selectCount');
        const unlockBtn = document.getElementById('unlockBtn');
        const prereqBtn = document.getElementById('prereqBtn');
        const toggleBtn = document.getElementById('toggleAll');
        let allSelected = false;

        function updateCount() {
            const n = document.querySelectorAll('.locked-chk:checked').length;
            countEl.textContent = n + ' selected';
            unlockBtn.disabled = n === 0;
            prereqBtn.disabled = n === 0;
        }

        checks.forEach(c => c.addEventListener('change', updateCount));

        toggleBtn.addEventListener('click', function() {
            allSelected = !allSelected;
            checks.forEach(c => { c.checked = allSelected; });
            toggleBtn.textContent = allSelected ? 'DESELECT ALL' : 'SELECT ALL';
            updateCount();
        });
    })();
    </script>
</body>
</html>`);
  } catch (err) {
    console.error("Locked items page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// POST /admin/catalog/locked/update — Unlock items or update prerequisites
// ---------------------------------------------------------------------------
app.post("/admin/catalog/locked/update", async (req, res) => {
  try {
    let rows = req.body.rows;
    const action = req.body.action;
    if (!rows) return res.redirect("/admin/catalog/locked");
    if (!Array.isArray(rows)) rows = [rows];

    const sheets = await getSheets();

    // Get student sheet headers for column positions
    const student = await fetchStudentSectors(sheets);
    const headers = student.headers;
    const statusCol = headers.indexOf("Status");
    const lockedForCol = headers.indexOf("Locked for what?");

    const updates = [];

    for (const rowNum of rows) {
      const r = parseInt(rowNum, 10);
      if (isNaN(r)) continue;

      if (action === "unlock" && statusCol >= 0) {
        // Change status from Locked to Engaged
        const col = String.fromCharCode(65 + statusCol);
        updates.push({ range: `Sectors!${col}${r}`, values: [["Engaged"]] });
      }

      if (action === "prereq" && lockedForCol >= 0) {
        // Update prerequisite field
        const prereqKey = `prereq_${r}`;
        const prereqValue = req.body[prereqKey] || "";
        const col = String.fromCharCode(65 + lockedForCol);
        updates.push({ range: `Sectors!${col}${r}`, values: [[prereqValue]] });
      }
    }

    if (updates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { valueInputOption: "RAW", data: updates },
      });
    }

    res.redirect(`/admin/catalog/locked?updated=${rows.length}`);
  } catch (err) {
    console.error("Locked update error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// POST /admin/catalog/add — Add selected catalog items to student sheet
// ---------------------------------------------------------------------------
app.post("/admin/catalog/add", async (req, res) => {
  try {
    let items = req.body.items;
    if (!items) return res.redirect("/admin/catalog");
    if (!Array.isArray(items)) items = [items];

    // Parse JSON strings from form and pair with statuses
    const parsed = items.map((i, idx) => {
      const item = typeof i === "string" ? JSON.parse(i) : i;
      // Find the matching status field — the form sends status_N where N is the data-idx
      // Since only checked items are submitted, we need to find the right status
      return item;
    });

    const sheets = await getSheets();
    const student = await fetchStudentSectors(sheets);
    const studentHeaders = student.headers;

    // Extract all status_N fields from the form body
    const statusMap = {};
    for (const key in req.body) {
      if (key.startsWith("status_")) {
        statusMap[key] = req.body[key];
      }
    }

    // Build rows for student sheet with per-item status
    const newRows = parsed.map((item, i) => {
      // Try to find matching status by iterating status fields in order
      const statusKeys = Object.keys(statusMap).sort((a, b) => parseInt(a.split("_")[1]) - parseInt(b.split("_")[1]));
      const status = statusKeys[i] ? statusMap[statusKeys[i]] : "Engaged";
      return buildSectorsRowFromCatalog(studentHeaders, item, status);
    });

    // Append to student Sectors sheet
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Sectors!A:Z",
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: { values: newRows },
    });

    // Update catalog "In Henry's Sheet" column
    const catalogRes = await sheets.spreadsheets.values.get({
      spreadsheetId: CATALOG_SPREADSHEET_ID,
      range: "Catalog",
    });
    const catValues = catalogRes.data.values || [];
    const catHeaders = catValues[0] || [];
    const inSheetCol = catHeaders.indexOf("In Henry's Sheet");
    const catSectorCol = catHeaders.indexOf("Sector");
    const catBossCol = catHeaders.indexOf("Boss");
    const catMinionCol = catHeaders.indexOf("Minion");

    if (inSheetCol >= 0) {
      const addedKeys = new Set(parsed.map((p) => `${p.Sector}|${p.Boss}|${p.Minion}`));
      const updates = [];
      for (let i = 1; i < catValues.length; i++) {
        const row = catValues[i];
        const key = `${row[catSectorCol] || ""}|${row[catBossCol] || ""}|${row[catMinionCol] || ""}`;
        if (addedKeys.has(key)) {
          const col = String.fromCharCode(65 + inSheetCol);
          updates.push({ range: `Catalog!${col}${i + 1}`, values: [["Yes"]] });
        }
      }
      if (updates.length > 0) {
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: CATALOG_SPREADSHEET_ID,
          requestBody: { valueInputOption: "RAW", data: updates },
        });
      }
    }

    // Redirect back with success count
    const returnSubjects = req.body.returnSubjects || "";
    res.redirect(`/admin/catalog/view?subjects=${returnSubjects}&added=${parsed.length}`);
  } catch (err) {
    console.error("Catalog add error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// GET /admin/catalog/generate — AI objective generation page
// ---------------------------------------------------------------------------
app.get("/admin/catalog/generate", async (req, res) => {
  try {
    const sheets = await getSheets();
    const catalog = await fetchCatalogData(sheets);

    // Get unique sectors and subjects
    const sectorSubjects = {};
    for (const row of catalog.rows) {
      const sector = row["Sector"];
      const subject = row["Subject"] || sector;
      if (sector && !sectorSubjects[sector]) sectorSubjects[sector] = subject;
    }

    const options = Object.entries(sectorSubjects)
      .sort((a, b) => a[1].localeCompare(b[1]))
      .map(([sector, subject]) =>
        `<option value="${escHtml(sector)}">${escHtml(subject)} (${escHtml(sector)})</option>`
      ).join("");

    res.send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Generate Objectives - Sovereign HUD</title>
    <style>
    ${CATALOG_CSS}
    .gen-form {
        max-width: 600px;
        margin: 0 auto 30px auto;
    }
    .form-group {
        margin-bottom: 20px;
    }
    .form-group label {
        display: block;
        margin-bottom: 6px;
        font-size: 0.8em;
        letter-spacing: 2px;
        color: #00f2ff;
    }
    .form-group select, .form-group textarea {
        width: 100%;
        background: #1a1d26;
        border: 1px solid #333;
        color: #ffea00;
        padding: 10px;
        font-family: 'Courier New', monospace;
        font-size: 0.9em;
    }
    .form-group select:focus, .form-group textarea:focus {
        outline: none;
        border-color: #ffea00;
    }
    .spinner {
        display: none;
        text-align: center;
        padding: 30px;
        color: #ff00ff;
        font-size: 0.9em;
        letter-spacing: 2px;
    }
    .spinner.active { display: block; }
    .results { display: none; }
    .results.active { display: block; }
    .result-card {
        border: 1px solid #333;
        border-radius: 4px;
        padding: 12px;
        margin-bottom: 10px;
        display: flex;
        align-items: flex-start;
        gap: 12px;
    }
    .result-card:hover { border-color: #ff00ff; }
    .result-chk { accent-color: #ff00ff; width: 16px; height: 16px; margin-top: 3px; cursor: pointer; }
    .result-info { flex: 1; }
    .result-boss { font-size: 0.8em; color: #ff00ff; margin-bottom: 3px; }
    .result-minion { font-size: 0.9em; color: #ffea00; }
    .result-meta { font-size: 0.7em; color: #666; margin-top: 4px; text-transform: none; }
    .result-actions {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-top: 20px;
        flex-wrap: wrap;
        gap: 10px;
    }
    .result-count { font-size: 0.8em; color: #888; }
    </style>
</head>
<body>
    <div class="hud-container">
        <div class="nav-row">
            <a class="back-link" href="/admin/catalog">&lt; BACK TO CATALOG</a>
            <a class="hud-link" href="/">HUD &gt;</a>
        </div>
        <h1>Populate Catalog with AI</h1>
        <div class="subtitle">ADD AI-GENERATED HIGH-SCHOOL OBJECTIVES TO THE MASTER CATALOG FOR REVIEW</div>

        <div class="gen-form" id="genForm">
            <div class="form-group">
                <label>Subject / Sector</label>
                <select id="sectorSelect">
                    <option value="">-- Select a subject --</option>
                    ${options}
                </select>
            </div>
            <div class="form-group">
                <label>How many objectives?</label>
                <select id="countSelect">
                    <option value="5">5</option>
                    <option value="10" selected>10</option>
                    <option value="20">20</option>
                </select>
            </div>
            <div class="form-group">
                <label>Focus area or context (optional)</label>
                <textarea id="contextInput" rows="3" placeholder="e.g., focus on trigonometry, or skip basic topics..." style="text-transform:none"></textarea>
            </div>
            <button class="btn btn-magenta" id="generateBtn">GENERATE OBJECTIVES</button>
        </div>

        <div class="spinner" id="spinner">ANALYZING EXISTING OBJECTIVES AND GENERATING NEW ONES...</div>

        <div class="results" id="results">
            <div id="resultCards"></div>
            <div class="result-actions">
                <div>
                    <button class="btn" id="selectAllResults" style="font-size:0.75em;padding:6px 12px">SELECT ALL</button>
                    <span class="result-count" id="resultCount">0 selected</span>
                </div>
                <button class="btn btn-magenta" id="approveBtn" disabled>ADD SELECTED TO CATALOG</button>
            </div>
        </div>
    </div>
    <script>
    (function() {
        const genBtn = document.getElementById('generateBtn');
        const spinner = document.getElementById('spinner');
        const resultsDiv = document.getElementById('results');
        const cardsDiv = document.getElementById('resultCards');
        const approveBtn = document.getElementById('approveBtn');
        const resultCount = document.getElementById('resultCount');
        const selectAllBtn = document.getElementById('selectAllResults');
        let generatedItems = [];
        let allSelected = false;

        function updateResultCount() {
            const n = document.querySelectorAll('.result-chk:checked').length;
            resultCount.textContent = n + ' selected';
            approveBtn.disabled = n === 0;
        }

        genBtn.addEventListener('click', async function() {
            const sector = document.getElementById('sectorSelect').value;
            if (!sector) { alert('Please select a subject'); return; }
            const count = document.getElementById('countSelect').value;
            const context = document.getElementById('contextInput').value;

            document.getElementById('genForm').style.display = 'none';
            spinner.className = 'spinner active';
            resultsDiv.className = 'results';

            try {
                const resp = await fetch('/admin/catalog/generate/run', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ sector, count: parseInt(count), context }),
                });
                const data = await resp.json();
                if (data.error) throw new Error(data.error);

                generatedItems = data.items;
                cardsDiv.innerHTML = generatedItems.map(function(item, i) {
                    return '<div class="result-card">' +
                        '<input type="checkbox" class="result-chk" data-idx="' + i + '" checked>' +
                        '<div class="result-info">' +
                            '<div class="result-boss">' + item.boss + ' &mdash; Impact: ' + item.impact + '</div>' +
                            '<div class="result-minion">' + item.minion + '</div>' +
                            '<div class="result-meta">Proof: ' + (item.suggestedProofMethod || 'N/A') + ' | ' + (item.reasoning || '') + '</div>' +
                        '</div>' +
                    '</div>';
                }).join('');

                allSelected = true;
                selectAllBtn.textContent = 'DESELECT ALL';
                document.querySelectorAll('.result-chk').forEach(function(c) {
                    c.addEventListener('change', updateResultCount);
                });
                updateResultCount();

                spinner.className = 'spinner';
                resultsDiv.className = 'results active';
            } catch(e) {
                spinner.className = 'spinner';
                document.getElementById('genForm').style.display = 'block';
                alert('Generation failed: ' + e.message);
            }
        });

        selectAllBtn.addEventListener('click', function() {
            allSelected = !allSelected;
            document.querySelectorAll('.result-chk').forEach(function(c) { c.checked = allSelected; });
            selectAllBtn.textContent = allSelected ? 'DESELECT ALL' : 'SELECT ALL';
            updateResultCount();
        });

        approveBtn.addEventListener('click', async function() {
            const selected = [];
            document.querySelectorAll('.result-chk:checked').forEach(function(c) {
                selected.push(generatedItems[parseInt(c.dataset.idx)]);
            });
            if (selected.length === 0) return;

            approveBtn.disabled = true;
            approveBtn.textContent = 'WRITING TO CATALOG...';

            try {
                const sector = document.getElementById('sectorSelect').value;
                const resp = await fetch('/admin/catalog/generate/approve', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ sector, items: selected }),
                });
                const data = await resp.json();
                if (data.error) throw new Error(data.error);
                window.location.href = '/admin/catalog/view?subjects=' + encodeURIComponent(document.getElementById('sectorSelect').selectedOptions[0].text.split(' (')[0]);
            } catch(e) {
                alert('Failed to save: ' + e.message);
                approveBtn.disabled = false;
                approveBtn.textContent = 'ADD SELECTED TO CATALOG';
            }
        });
    })();
    </script>
</body>
</html>`);
  } catch (err) {
    console.error("Generate page error:", err);
    res.status(500).send(`<pre style="color:red">Error: ${err.message}</pre>`);
  }
});

// ---------------------------------------------------------------------------
// POST /admin/catalog/generate/run — Execute AI generation
// ---------------------------------------------------------------------------
app.post("/admin/catalog/generate/run", async (req, res) => {
  try {
    const { sector, count, context } = req.body;
    if (!sector || !count) return res.status(400).json({ error: "Sector and count required" });
    if (!ANTHROPIC_API_KEY) return res.status(500).json({ error: "ANTHROPIC_API_KEY not configured" });

    const sheets = await getSheets();
    const catalog = await fetchCatalogData(sheets);

    // Get existing items for this sector
    const existing = catalog.rows.filter((r) => r["Sector"] === sector);
    const subject = existing[0]?.Subject || sector;
    const bossList = {};
    for (const item of existing) {
      const boss = item["Boss"] || "Unknown";
      if (!bossList[boss]) bossList[boss] = [];
      bossList[boss].push(item["Minion"]);
    }

    const existingDesc = Object.entries(bossList)
      .map(([boss, minions]) => `  Boss "${boss}": ${minions.join(", ")}`)
      .join("\n");

    const client = new Anthropic({ apiKey: ANTHROPIC_API_KEY });

    const response = await client.messages.create({
      model: "claude-sonnet-4-20250514",
      max_tokens: 4096,
      system: `You are a curriculum designer for a gamified homeschool tracker.
You create learning objectives ("Minions") organized under topics ("Bosses") within subject areas ("Sectors").

The student is in high school (grades 9-12). Objectives must be grade-appropriate.

SECTOR: ${sector}
SUBJECT: ${subject}

EXISTING BOSSES AND MINIONS IN THIS SECTOR:
${existingDesc || "  (none yet)"}

RULES:
1. Generate exactly ${count} new objectives that DO NOT duplicate existing ones.
2. Each objective needs: boss (topic name), minion (specific skill/concept, 2-6 words), impact (1-3 scale: 1=basic, 2=intermediate, 3=advanced), suggestedProofMethod (one of: Document, Spreadsheet, Presentation, Video, Project).
3. Prefer assigning to existing Bosses when the topic fits. Only create new Bosses if the objective doesn't fit existing ones.
4. New Boss names should be creative/thematic (like "The Algebraist", "The Physicist") — not plain subject names.
5. Objectives should be concrete and assessable, not vague.
6. Vary the impact levels across the set.
7. Return ONLY a valid JSON array, no other text.`,
      messages: [{
        role: "user",
        content: `Generate ${count} new high-school-level learning objectives for the ${subject} sector.${context ? " Additional context: " + context : ""}

Return a JSON array:
[{"boss":"Topic Name","minion":"Objective Name","impact":2,"suggestedProofMethod":"Document","reasoning":"Brief explanation"}]`,
      }],
    });

    const text = response.content[0].text;
    const jsonMatch = text.match(/\[[\s\S]*\]/);
    if (!jsonMatch) return res.status(500).json({ error: "AI did not return valid JSON" });

    const items = JSON.parse(jsonMatch[0]);
    res.json({ items, usage: { input: response.usage.input_tokens, output: response.usage.output_tokens } });
  } catch (err) {
    console.error("AI generation error:", err);
    res.status(500).json({ error: err.message });
  }
});

// ---------------------------------------------------------------------------
// POST /admin/catalog/generate/approve — Write AI suggestions to catalog
// ---------------------------------------------------------------------------
app.post("/admin/catalog/generate/approve", async (req, res) => {
  try {
    const { sector, items } = req.body;
    if (!items || !items.length) return res.status(400).json({ error: "No items to approve" });

    const sheets = await getSheets();

    // Get catalog headers to match column order
    const catalog = await fetchCatalogData(sheets);
    const headers = catalog.headers;

    // Get subject name from existing catalog rows
    const existingInSector = catalog.rows.find((r) => r["Sector"] === sector);
    const subject = existingInSector?.Subject || sector;

    const newRows = items.map((item) => {
      const valueMap = {
        Sector: sector,
        Subject: subject,
        Boss: item.boss,
        Minion: item.minion,
        Status: "Locked",
        "Impact(1-3)": String(item.impact || 2),
        "Locked for what?": "",
        "Suggested Proof Method": item.suggestedProofMethod || "",
        "Grade Level": "High School",
        "In Henry's Sheet": "No",
        Source: "AI Generated",
      };
      return headers.map((h) => valueMap[h] ?? "");
    });

    await sheets.spreadsheets.values.append({
      spreadsheetId: CATALOG_SPREADSHEET_ID,
      range: "Catalog!A:Z",
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: { values: newRows },
    });

    // Expand the table filter and banding to include new rows
    try {
      await expandCatalogTable(sheets);
      console.log("Catalog table filter expanded successfully.");
    } catch (expandErr) {
      console.error("Warning: failed to expand catalog table filter:", expandErr.message);
    }

    res.json({ success: true, count: newRows.length });
  } catch (err) {
    console.error("Catalog approve error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Sovereign HUD online at http://localhost:${PORT}`);
});

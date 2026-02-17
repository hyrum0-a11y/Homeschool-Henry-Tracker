/**
 * 1. THE WEB APP "FRONT DOOR"
 */
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Command_Center");

  if (!dataSheet) {
    return HtmlService.createHtmlOutput("Error: Sheet 'Command_Center' missing.");
  }

  const rawHtml = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Sovereign HUD</title>
    <style>
    /* 1. GLOBAL SYSTEM STYLES */
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
        max-width: 900px; 
        margin: auto; 
    }
    h1 { 
        text-shadow: 2px 2px #ff00ff; 
        border-bottom: 1px solid #00f2ff; 
        letter-spacing: 2px; 
        text-align: center; 
        margin-bottom: 15px; 
    }

    /* 2. CORE STATS & PROGRESSION */
    #header-section { margin-bottom: 25px; }

    .confidence-section { 
        margin-bottom: 25px; 
        border-bottom: 1px dashed #ffea00; 
        padding-bottom: 15px; 
    }

    .core-layout {
        display: flex;
        flex-direction: row;
        justify-content: space-between;
        gap: 30px;
        margin-bottom: 40px;
    }

    .stats-left { flex: 7; }
    
    .stat-grid { 
        display: grid; 
        grid-template-columns: 1fr 1fr; 
        gap: 20px; 
    }

    /* Stat Bars */
    .stat-label { font-size: 0.85em; font-weight: bold; margin-bottom: 5px; }
    .stat-bar { background: #1a1d26; border: 1px solid #333; height: 25px; position: relative; overflow: hidden; } /* Added overflow: hidden */
    .fill { height: 100%; transition: width 0.5s ease-in-out; }

    /* Bar Gradients */
    .intel-fill      { background: linear-gradient(90deg, #00f2ff, #0077ff); box-shadow: 0 0 15px #00f2ff; }
    .stamina-fill    { background: linear-gradient(90deg, #00ff9d, #008844); box-shadow: 0 0 10px #00ff9d; }
    .tempo-fill      { background: linear-gradient(90deg, #ff00ff, #880088); box-shadow: 0 0 10px #ff00ff; }
    .rep-fill        { background: linear-gradient(90deg, #ff8800, #ff4400); box-shadow: 0 0 10px #ff8800; }
    .confidence-fill { background: linear-gradient(90deg, #ffea00, #ffaa00); box-shadow: 0 0 10px #ffea00; }

    /* 3. METAL TIER LIST (RIGHT SIDE) */
    .levels-right {
        flex: 3;
        border-left: 1px solid rgba(0, 242, 255, 0.2);
        padding-left: 20px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }

    .tier-list { list-style: none; padding: 0; margin: 0; }
    
    .tier-item {
        font-size: 0.8em;
        margin: 10px 0;
        padding-left: 0;
        border-left: none; 
        transition: all 0.3s ease;
        letter-spacing: 1px;
    }

    /* Metal Color Logic */
    .rank-platinum { color: #e5e4e2; opacity: 0.4; }
    .rank-gold     { color: #ffd700; opacity: 0.4; }
    .rank-silver   { color: #c0c0c0; opacity: 0.4; }
    .rank-copper   { color: #b87333; opacity: 0.4; }
    .rank-bronze   { color: #cd7f32; opacity: 0.4; }

    /* Highlighting Active Rank (Option: Glow + Slide) */
    .tier-item.active { 
        opacity: 1; 
        font-weight: bold;
        text-shadow: 0 0 12px currentColor;
        transform: translateX(10px);
    }

    .tier-item.next { opacity: 0.7; }

    /* 4. THE FRONTIER & SECTOR MAP */
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
        min-width: 200px;
        display: flex;
        flex-wrap: wrap;
        justify-content: center;
        gap: 20px;
        flex: 1 1 auto;
        max-width: 400px;
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

    .boss-orb-container {
        display: flex;
        flex-direction: column;
        align-items: center;
        width: 100px; /* Slightly wider for the (X/X) labels */
        position: relative;
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

    /* 5. BOSS KEY / LEGEND */
    .boss-key {
        display: flex;
        justify-content: center;
        gap: 25px;
        margin-bottom: 20px;
        font-size: 0.75em;
    }
    .key-item {
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .key-swatch {
        display: inline-block;
        width: 12px;
        height: 12px;
        border-radius: 50%;
    }

    /* 6. UTILITY & ANIMATION */
    @keyframes pulse { 0% { opacity: 1; } 50% { opacity: 0.7; } 100% { opacity: 1; } }
    .glitch { font-size: 0.7em; color: #555; margin-top: 20px; text-align: center; line-height: 1.6; }


    /* 6. MOBILE RESPONSIVE */
    @media (max-width: 600px) {
        body { padding: 10px; }
        .hud-container { padding: 15px; }
        .core-layout {
            flex-direction: column;
            gap: 15px;
        }
        .levels-right {
            border-left: none;
            border-top: 1px solid rgba(0, 242, 255, 0.2);
            padding-left: 0;
            padding-top: 15px;
        }
        .stat-grid {
            grid-template-columns: 1fr;
        }
        .sector-zone {
            min-width: unset;
            max-width: 100%;
        }
    }

</style>
</head>
<body>
    <div class="hud-container">
        <h1>Henry's Sovereign HUD</h1>
        <div class="confidence-section">
            <div class="stat-label" style="color: #ffea00; font-size: 1.1em;">
                CONFIDENCE: [[CONF_RANK]], +[[CONF_REM]] FOR [[CONF_NEXT]]
            </div>
            <div class="stat-bar">
                <div class="fill confidence-fill" style="width: [[CONF_BAR]]%;"></div>
            </div>
        </div>

        <div class="core-layout">
            <div class="stats-left">
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
            </div>

            <div class="levels-right">
                [[TIER_LIST]]
            </div>
        </div>

        <div class="sector-map">
            <h1>Bosses & Minions </h1>
            <div class="boss-key">
                <span class="key-item"><span class="key-swatch" style="background:#00ff9d;"></span> Enslaved</span>
                <span class="key-item"><span class="key-swatch" style="background:#ff6600;"></span> Engaged</span>
                <span class="key-item"><span class="key-swatch" style="background:#2a2d36; border: 1px solid #555;"></span> Locked</span>
            </div>
            <div id="bosses">[[BOSS_LIST]]</div>
        </div>

        <div class="glitch">
            <span style="color: #ff00ff;">SYSTEM: ONLINE</span><br>
            > DATA_STREAM_SYNCED...<br>
            > NO_THREATS_DETECTED_IN_CORE...
        </div>
    </div>

</body>
</html>`;
  const processedHtml = processAllData(rawHtml, dataSheet);

  return HtmlService.createHtmlOutput(processedHtml)
    .setTitle("Sovereign HUD")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 2. THE DATA ENGINE
 */
function processAllData(html, dataSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const defSheet = ss.getSheetByName("Definitions");
  const sectorSheet = ss.getSheetByName("Sectors"); // NEW DATA SOURCE
  let bossHtml = ""; 

  // 1. Pull the Definitions Table
  const defData = defSheet.getRange("E2:I6").getValues(); 

  // 2. Pull Core Stats
  const intel = dataSheet.getRange("B2").getValue() || 0; 
  const stamina = dataSheet.getRange("B3").getValue() || 0;
  const tempo = dataSheet.getRange("B4").getValue() || 0;
  const reputation = dataSheet.getRange("B5").getValue() || 0;
  const confidence = dataSheet.getRange("B6").getValue() || 0;
  const currentConfRank = dataSheet.getRange("C6").getValue() || "BRONZE: INITIATE";

  // 3. Status Rank Lookup Logic
  function getStatusRankInfo(pts) {
    let current = "BRONZE";
    let next = "COPPER";
    for (let i = 0; i < defData.length; i++) {
      if (pts >= defData[i][4]) { 
        current = defData[i][1];
        next = (i + 1 < defData.length) ? defData[i+1][1] : "MAX";
      }
    }
    return { current: current, next: next };
  }

  // 4. Next Confidence Tier Lookup
  let nextConfTier = "MAX LEVEL";
  for (let i = 0; i < defData.length; i++) {
    let tierName = defData[i][1];
    if (currentConfRank.indexOf(tierName) !== -1) {
      if (i + 1 < defData.length) {
        nextConfTier = defData[i+1][1]; 
      }
      break;
    }
  }

  // 5. BUILD THE FRONTIER DATA (ENSLAVEMENT ENGINE)
  // Dynamic header lookup — column order doesn't matter
  const sectorHeaders = sectorSheet.getRange(1, 1, 1, sectorSheet.getLastColumn()).getValues()[0];
  const sCol = sectorHeaders.indexOf("Sector");
  const bCol = sectorHeaders.indexOf("Boss");
  const stCol = sectorHeaders.indexOf("Status");
  const lastDataRow = Math.max(sectorSheet.getLastRow(), 2);
  const sectorData = sectorSheet.getRange(2, 1, lastDataRow - 1, sectorSheet.getLastColumn()).getValues();
  let bossMap = {};

  // Step A: Group Minions by Boss and Count by Status
  for (let i = 0; i < sectorData.length; i++) {
    let sector = sectorData[i][sCol];
    let bossName = sectorData[i][bCol];
    let status = sectorData[i][stCol];

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

  // Step B: Build HTML (Sorted by Size/Threat Level)
  for (let sector in bossMap) {
    let sectorHtml = "";
    
    // 1. Sort bosses by minion count (largest first)
    let sortedBossNames = Object.keys(bossMap[sector]).sort((a, b) => {
      return bossMap[sector][b].total - bossMap[sector][a].total;
    });

    // 2. Loop through the sorted list to generate the HTML
    for (let bossName of sortedBossNames) {
      let stats = bossMap[sector][bossName];
      let fraction = `(${stats.enslaved}/${stats.total})`;

      // Diameter scales with minion count
      let diameter = 35 + (stats.total * 6);
      let safeTotal = stats.total || 1;

      // Pie slices: Enslaved (green), Engaged (orange), Locked (dark gray)
      // 1% dark gaps act as separator lines between slices
      let pEnslaved = (stats.enslaved / safeTotal) * 100;
      let pEngaged = (stats.engaged / safeTotal) * 100;

      let gradient = `conic-gradient(
        #00ff9d 0% ${pEnslaved}%,
        #0a0b10 ${pEnslaved}% ${pEnslaved + 1}%,
        #ff6600 ${pEnslaved + 1}% ${pEnslaved + 1 + pEngaged}%,
        #0a0b10 ${pEnslaved + 1 + pEngaged}% ${pEnslaved + 2 + pEngaged}%,
        #2a2d36 ${pEnslaved + 2 + pEngaged}% 100%
      )`;

      sectorHtml += `
        <div class="boss-orb-container">
          <div class="boss-pie" style="width:${diameter}px; height:${diameter}px; background:${gradient};"></div>
          <div class="boss-label">${bossName} ${fraction}</div>
        </div>`;
    }
    
    // Wrap the sorted bosses into their Sector Territory
    bossHtml += `<div class="sector-zone" data-sector="${sector.toUpperCase()}">${sectorHtml}</div>`;
  }

  // 6. Build Dynamic Tier List
  const tiers = [
    { key: "PLATINUM", rank: "platinum", label: "PLATINUM: Leader" },
    { key: "GOLD", rank: "gold", label: "GOLD: Independent" },
    { key: "SILVER", rank: "silver", label: "SILVER: Reliable" },
    { key: "COPPER", rank: "copper", label: "COPPER: Adept" },
    { key: "BRONZE", rank: "bronze", label: "BRONZE: Initiate" }
  ];
  let activeIdx = tiers.length - 1;
  for (let i = 0; i < tiers.length; i++) {
    if (currentConfRank.toUpperCase().indexOf(tiers[i].key) !== -1) {
      activeIdx = i;
      break;
    }
  }
  let tierListHtml = '<ul class="tier-list">';
  for (let i = 0; i < tiers.length; i++) {
    let cls = "tier-item rank-" + tiers[i].rank;
    if (i === activeIdx) cls += " active";
    else if (i === activeIdx - 1) cls += " next";
    tierListHtml += '<li class="' + cls + '">' + tiers[i].label + '</li>';
  }
  tierListHtml += '</ul>';

  // 7. Final HUD Data Replacements
  const iI = getStatusRankInfo(intel);
  const sI = getStatusRankInfo(stamina);
  const tI = getStatusRankInfo(tempo);
  const rI = getStatusRankInfo(reputation);

  return html.split("[[INTEL_RANK]]").join(dataSheet.getRange("C2").getValue())
             .split("[[INTEL_REM]]").join(parseFloat(dataSheet.getRange("F2").getValue()).toFixed(1))
             .split("[[INTEL_BAR]]").join(Math.min(100, (intel % 20 / 20) * 100).toFixed(1))
             
             .split("[[STAMINA_RANK]]").join(dataSheet.getRange("C3").getValue())
             .split("[[STAMINA_REM]]").join(parseFloat(dataSheet.getRange("F3").getValue()).toFixed(1))
             .split("[[STAMINA_BAR]]").join(Math.min(100, (stamina % 20 / 20) * 100).toFixed(1))

             .split("[[TEMPO_RANK]]").join(dataSheet.getRange("C4").getValue())
             .split("[[TEMPO_REM]]").join(parseFloat(dataSheet.getRange("F4").getValue()).toFixed(1))
             .split("[[TEMPO_BAR]]").join(Math.min(100, (tempo % 20 / 20) * 100).toFixed(1))

             .split("[[REP_RANK]]").join(dataSheet.getRange("C5").getValue())
             .split("[[REPUTATION_REM]]").join(parseFloat(dataSheet.getRange("F5").getValue()).toFixed(1))
             .split("[[REPUTATION_BAR]]").join(Math.min(100, (reputation % 20 / 20) * 100).toFixed(1))

             .split("[[CONF_RANK]]").join(currentConfRank)
             .split("[[CONF_NEXT]]").join(nextConfTier)
             .split("[[CONF_REM]]").join(parseFloat(dataSheet.getRange("F6").getValue()).toFixed(1))
             .split("[[CONF_BAR]]").join(Math.min(100, (confidence % 80 / 80) * 100).toFixed(1))
             .split("[[TIER_LIST]]").join(tierListHtml)
             .split("[[BOSS_LIST]]").join(bossHtml);
}

/**
 * 3. THE TOOLBAR MENU
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 SYSTEM')
    .addItem('Get HUD Link', 'showUrl')
    .addToUi();
}

function showUrl() {
  const url = ScriptApp.getService().getUrl();
  const htmlOutput = HtmlService.createHtmlOutput(
    '<div style="padding:20px; font-family:monospace; background:#0a0b10; color:#00f2ff; text-align:center;">' +
    '<h3 style="color:#ff00ff;">SYSTEM SYNC READY</h3>' +
    '<button onclick="launch()" style="background:#00f2ff; color:#000; padding:15px 30px; border:none; font-weight:bold; cursor:pointer; border-radius:5px; box-shadow: 0 0 10px #00f2ff;">' +
    'BOOT HUD</button>' +
    '<script>function launch() { window.open("' + url + '", "_blank"); google.script.host.close(); }</script>' +
    '</div>'
  ).setWidth(400).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Terminal Access');
}
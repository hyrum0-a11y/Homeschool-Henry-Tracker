# Homeschool Henry Tracker (Sovereign HUD)

A gamified homeschool progress tracker that turns learning objectives into a video-game-style HUD with stats, quests, and sector maps. Built with Node.js/Express, Google Sheets as the database, and Claude AI for content generation.

## Features

### Student HUD Dashboard
- Animated shimmer title with Minecraft pixel art sprites
- Radar charts, confidence bar, and 4 individual stat bars (Intel, Stamina, Tempo, Reputation)
- Stat bars show current rank with `+N for [next level]` format
- **Game Mode Display** — Creative Mode / Survival Mode banner with Minecraft-style pixel hearts that glow proportionally to boss conquest progress
- **Ring of Guardians** — Visual orbs showing survival-required bosses with conic gradient fill indicating progress
- **Boss Conquest Rankings** — Bosses with remaining minions, sorted by completion percentage, linking to boss detail pages
- Sector map with clickable boss orbs and pie-chart progress indicators

### Quest Board
- Students take on quests to prove mastery of learning objectives
- Task details pulled from Sectors sheet (not generic AI suggestions)
- Artifact type dropdown and proof link submission
- Submit button disabled until proof text is entered
- Multi-select quest adding from boss and sector pages (batch checkbox selection)
- Quest status synced to Sectors sheet "Quest Status" column on every transition

### Quest Approval (Admin)
- Full quest lifecycle: Active → Submitted → Approved / Rejected
- Approve auto-enslaves the minion (Status → "Enslaved" on Sectors sheet)
- Reject keeps quest on board for student feedback; can be reopened
- Sync button to backfill existing quest statuses to Sectors sheet

### Survival Mode System
- Per-boss "Survival Mode Required" flag (Minecraft-inspired Creative vs Survival concept)
- Manage toggles from admin with confirmation dialogs in both directions
- HUD shows mode banner, proportionally glowing hearts, and Ring of Guardians orbs
- Hearts and orbs sorted by conquest rate (highest first, then alphabetical)

### Objective Catalog (removed from web UI)
- Catalog routes have been removed from the web interface
- The Master Catalog Google Sheet is retained for reference and future use
- Student items are managed directly in the Sectors Google Sheet

### Other Features
- **AI Photo Import** — Upload lesson photos and Claude AI classifies them into the tracker
- **AI Catalog Population** — Generate high-school-level learning objectives with AI
- **Locked Items Management** — Set prerequisites and unlock objectives when ready
- **Army Page** — View all completed (enslaved) minions grouped by sector

## Architecture

- **server.js** — Single-file Express server (~5000+ lines) serving inline HTML with a dark cyberpunk/Minecraft hybrid theme
- **Google Sheets** — Two spreadsheets:
  - **Student Sheet** — Per-student data (Sectors, Definitions, Command_Center, Quests tabs)
  - **Master Catalog** — Shared catalog of all learning objectives with subject mappings
- **Claude AI** — Used for photo classification and objective generation (via Anthropic API)

### Key Patterns
- Dynamic column lookup (never hardcoded column positions)
- Row builder pattern: `valueMap` + `headers.map((h) => valueMap[h] ?? "")`
- Composite key matching: `Sector|Boss|Minion` for cross-sheet identification
- Template placeholder system: `html.split("[[PLACEHOLDER]]").join(value)`
- Catalog data fetched with explicit range `Catalog!A:Z` (not bare sheet name) to ensure all rows are returned regardless of Google Sheets table/filter boundaries

## Status Definitions

| Status | Meaning |
|--------|---------|
| **Locked** | Has a prerequisite that must be completed first |
| **Engaged** | Available for the student to work on |
| **Enslaved** | Completed/mastered by the student |

### Quest Status (synced to Sectors sheet)

| Quest Status | Meaning |
|--------------|---------|
| **Active** | Quest started, student working on it |
| **Submitted** | Student submitted proof, awaiting teacher review |
| **Approved** | Teacher approved; minion auto-enslaved |
| **Rejected** | Teacher rejected; student can re-submit |

## Setup

### Prerequisites
- Node.js
- Google Cloud project with Sheets API and Drive API enabled
- Google service account with credentials.json
- Anthropic API key (for AI features)

### Environment Variables (.env)
```
SPREADSHEET_ID=<student spreadsheet id>
CATALOG_SPREADSHEET_ID=<master catalog spreadsheet id>
GOOGLE_CREDENTIALS_PATH=./credentials.json
PORT=3000
ANTHROPIC_API_KEY=<your key>
```

### Install & Run
```bash
npm install
npm start
```

Then open http://localhost:3000

### One-Time Setup Scripts
- `setup-drive.js` — Creates Google Drive folder structure and Master Catalog spreadsheet
- `setup-catalog-table.js` — Adds Subject column and formats catalog as a Google Sheets table
- `fix-catalog-filter.js` — Re-expands catalog table filter/banding after adding rows manually
- `fix-formulas.js` — Updates stat formulas to header-resilient INDEX/MATCH format
- `fix-task-validation.js` — Removes data validation dropdowns from the Task column
- `backfill-subject.js` — Backfills Subject column for existing Sectors rows from catalog
- `backfill-task.js` — Backfills Task column for existing Sectors rows from catalog

### Sectors Sheet Setup
Ensure the Sectors sheet has these columns in the header row:
- Sector, Subject, Boss, Minion, Task, Status, Impact(1-3), Locked for what?, Survival Mode Required, Quest Status, INTELLIGENCE, STAMINA, TEMPO, REPUTATION

## Routes

| Route | Description |
|-------|-------------|
| `/` | Main HUD dashboard |
| `/boss/:bossName` | Boss detail page with minion table and multi-select quest adding |
| `/sector/:sectorName` | Sector overview with all bosses and multi-select quest adding |
| `/quests` | Quest board with proof submission |
| `/army` | All enslaved minions |
| `/admin` | Parent admin console |
| `/admin/quests` | Quest approval (approve, reject, reopen, sync) |
| `/admin/import` | AI photo import |
| ~~`/admin/catalog`~~ | *(Removed)* Catalog routes removed from web UI |

## Sector-to-Subject Mapping

Sectors use gamified names internally. The catalog maps them to parent-friendly subject names at the Boss level:

| Sector | Example Subjects |
|--------|-----------------|
| COMMUNICATION | English, Spanish, Music, Visual Arts, Creative Writing |
| CORPOREAL | Chemistry, Physics, Geography |
| HUMANITY | Psychology, Government, Sociology |
| LOGIC | Algebra, Calculus, Trigonometry, Critical Thinking |
| SYSTEMS | Engineering, Architecture, Financial Literacy |
| VITALITY | Health & Fitness, Personal Care |

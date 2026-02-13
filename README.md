# Homeschool Henry Tracker (Sovereign HUD)

A gamified homeschool progress tracker that turns learning objectives into a video-game-style HUD with stats, quests, and sector maps. Built with Node.js/Express, Google Sheets as the database, and Claude AI for content generation.

## Features

- **Student HUD Dashboard** — Radar charts, stat bars, confidence ranking, sector map with boss orbs
- **Quest Board** — Students can take on quests to prove mastery of learning objectives
- **Objective Catalog** — Browse subjects, assign objectives to students, manage locked/engaged/enslaved statuses
- **AI Photo Import** — Upload lesson photos and Claude AI classifies them into the tracker
- **AI Catalog Population** — Generate high-school-level learning objectives with AI to populate the master catalog
- **Locked Items Management** — Set prerequisites and unlock objectives when students are ready
- **Army Page** — View all completed (enslaved) minions grouped by sector

## Architecture

- **server.js** — Single-file Express server serving inline HTML with a dark cyberpunk theme
- **Google Sheets** — Two spreadsheets:
  - **Student Sheet** — Per-student data (Sectors, Definitions, Command_Center, Quests tabs)
  - **Master Catalog** — Shared catalog of all learning objectives with subject mappings
- **Claude AI** — Used for photo classification and objective generation (via Anthropic API)

## Status Definitions

| Status | Meaning |
|--------|---------|
| **Engaged** | Available for the student to work on |
| **Locked** | Has a prerequisite that must be completed first |
| **Enslaved** | Completed/mastered by the student |

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
- `setup-drive.js` — Creates Google Drive folder structure
- `setup-catalog-table.js` — Adds Subject column and formats catalog as a Google Sheets table
- `fix-catalog-filter.js` — Re-expands catalog table filter after adding rows

## Routes

| Route | Description |
|-------|-------------|
| `/` | Main HUD dashboard |
| `/boss/:bossName` | Boss detail page with minion table |
| `/quests` | Quest board |
| `/army` | All enslaved minions |
| `/admin` | Parent admin console |
| `/admin/import` | AI photo import |
| `/admin/catalog` | Subject picker (objective catalog) |
| `/admin/catalog/view` | Browse objectives by subject |
| `/admin/catalog/student` | View current student items |
| `/admin/catalog/locked` | Manage locked items and prerequisites |
| `/admin/catalog/generate` | AI objective generation |

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

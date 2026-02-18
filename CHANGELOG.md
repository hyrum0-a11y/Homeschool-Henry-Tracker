# Changelog

## [2.6.0] - 2025-02-18

### Added
- **Timeline stat growth chart** on Progress page — SVG bar chart showing cumulative stat growth over time. Automatically selects daily/weekly/monthly view based on date range. Falls back to simple progress bars when no date data exists

### Changed
- **Quest board** — Abandoned quests no longer appear on the board (kept in Quests sheet for audit trail only)
- **Manual Entry form** — All fields (Sector, Subject, Boss, Minion) are now dropdown selects populated from existing data, with a "+ Add new..." option that reveals a text input. Impact labels updated to "Low/Medium/Highest impact to score"

## [2.5.0] - 2025-02-18

### Added
- **Quest date tracking** — New `Date Quest Added` and `Date Quest Completed` columns in Sectors sheet are automatically populated: date added when quest starts, date completed when parent approves, both cleared when quest is abandoned

### Changed
- **`updateSectorsQuestStatus` helper** — Now handles multi-column date tracking with support for columns beyond Z (AA, AB, etc.)
- **`buildSectorsRow`** — Includes empty `Date Quest Added` / `Date Quest Completed` columns for new rows

## [2.4.0] - 2025-02-18

### Added
- **Progress Report page** (`/progress`) — New accomplishments dashboard showing summary cards, stat level bars, sector conquest donuts, boss conquest chart, survival mode progress, and a timeline of completed quests. Linked from HUD via center bottom FAB button
- **Manual Entry page** (`/admin/manual`) — Parents can add new objectives (minions) directly from the admin panel without opening Google Sheets. Supports sector/boss/subject autocomplete from existing data, task description, impact level, and optional "add to quest board" toggle
- **Guardian badge on Henry's Army** — Minions belonging to survival-mode bosses are highlighted with a red shield icon, red text, and subtle red background tint
- **Audit trail for quest abandonment** — Instead of deleting abandoned quests, they are now marked "Abandoned" with the parent's name and date recorded in the Quests sheet

### Changed
- **Admin panel** — "Manual Entry" and "Progress Reports" cards now active and linked
- **Quest board** — "Abandoned" status added with dimmed card styling

## [2.3.0] - 2025-02-18

### Added
- **Guardians overview page** (`/guardians`) — Clicking the "Ring of Guardians" title opens a full page showing all survival-mode bosses with their minion tables, completion percentages, and quest board checkboxes
- **Survival badge on boss/sector pages** — Survival-mode bosses now display a red shield icon and "GUARDIAN" tag on sector detail pages; boss detail pages show a "SURVIVAL MODE GUARDIAN" banner
- **Parent permission for quest deletion** — Abandoning a quest now requires two confirmations: a standard "Are you sure?" prompt plus a parent name entry to discourage unauthorized deletions

### Changed
- **Ring of Guardians title** — Now a clickable link (yellow hover glow) that navigates to the guardians overview page

## [2.2.0] - 2025-02-18

### Added
- **Clickable Ring of Guardians** — Guardian orbs now have yellow hover glow, pointer cursor, and label highlight for better discoverability as links to boss pages
- **Boss Conquest top-10 limit** — Boss Conquest Rankings on the HUD now display only the top 10 incomplete bosses

### Removed
- **Catalog admin card** — Removed the greyed-out "Objective Catalog" card from the admin page entirely

## [2.1.0] - 2025-02-18

### Added
- **Full tier name in stat labels** — "for Silver I" instead of just "for I"; progress info styled as subtle inline text for stat bars, second line for confidence bar

### Removed
- **Objective Catalog web UI** — All catalog routes removed (`/admin/catalog`, `/admin/catalog/view`, `/admin/catalog/student`, `/admin/catalog/locked`, and related POST routes). Master Catalog Google Sheet retained for reference. Student items now managed directly in the Sectors sheet.
- ~1,580 lines of catalog helpers, CSS, and route handlers removed from server.js

### Fixed
- **Catalog data range** — Changed all `range: "Catalog"` to `range: "Catalog!A:Z"` across helper scripts to ensure new rows are always returned regardless of Google Sheets table/filter boundaries
- **Stat label next-level display** — `nextSubRank()` now returns the full tier name instead of stripping the metal prefix

## [2.0.0] - 2025-02-16

### Added
- **Quest Approval Admin Page** (`/admin/quests`) — Teachers can approve, reject, or reopen submitted quests
- **Quest Status synced to Sectors sheet** — New "Quest Status" column tracks Active/Submitted/Approved/Rejected per minion
- **Auto-enslave on approval** — Approving a quest automatically changes the minion's Status to "Enslaved"
- **Sync to Sectors button** — One-click backfill of existing quest statuses to the Sectors sheet
- **Multi-select quest adding** — Checkboxes on boss and sector pages allow batch-adding minions to the quest board
- **Batch quest route** (`POST /quest/start-batch`) — Creates multiple quests in a single Sheets API call
- **Game Mode Display** — Creative Mode / Survival Mode banner with Minecraft-style pixel hearts
- **Proportional heart glow** — Hearts glow from dim to bright based on boss conquest fraction
- **Ring of Guardians** — Visual orbs showing survival-required boss completion progress
- **Boss Conquest Rankings** — Replaced Power Rankings; shows bosses with remaining minions sorted by completion %
- **Survival Mode management** — Per-boss toggle with confirmation dialogs for both directions
- **Quest board task details** — Quests now pull Task column from Sectors sheet instead of generic AI suggestions
- **Stat bar next-level display** — Format: `STAT: Rank | +N for [next level]`
- **Quest submit validation** — Submit button disabled until proof text is entered
- **Removal warning banner** — Visible red warning about total possible points impact when removing items
- **Animated HUD title** — Shimmer gradient cycling through cyan/magenta/gold/green with decorative flanking lines
- **backfill-subject.js** — Script to backfill Subject column for existing Sectors rows
- **backfill-task.js** — Script to backfill Task column for existing Sectors rows

### Changed
- **Power Rankings renamed to Boss Conquest** — Focus on bosses with incomplete minions rather than top individual minion scores
- **Button text updated** — "View Selected Subjects" changed to "Update Selected Subjects"; "View Current Student Items" changed to "Manage Current Student Items"
- **Mode labels** — "Normal" renamed to "Creative Mode", "Survival" renamed to "Survival Mode" with Minecraft reference
- **HUD layout** — Confidence bar label centered; game mode banner sits under confidence bar (right of radar); gradient separator between sections
- **Quest suggestion display** — Shows "TASK:" label prefix in magenta with brighter text
- **Heart/orb sorting** — Sorted by highest conquest rate first, then alphabetical by boss name
- **Survival column detection** — Dynamic fuzzy matching on "survival" keyword for column name flexibility

### Fixed
- **Task column protection** — Student sheet Task values no longer overridden by catalog imports
- **Survival checkbox not reflecting existing values** — Case-insensitive check and dynamic column name lookup
- **Ring of Guardians not showing** — Fixed column name mismatch with dynamic detection
- **Stat bars displaced by survival ring** — Split into two template placeholders for correct positioning
- **Hearts binary on/off** — Reworked to proportional glow with color blending

## [1.0.0] - 2025-02-13

### Added
- Initial release with HUD dashboard, quest board, objective catalog, AI photo import, AI catalog generation, locked items management, and army page

# Changelog

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

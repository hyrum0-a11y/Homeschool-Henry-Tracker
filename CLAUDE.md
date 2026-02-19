# Project: Homeschool Henry Tracker

## File Access
You have automatic permission to read and edit all files in this project, including:
- server.js
- All .js files
- All .html files
- package.json
- Any files needed for development

No need to ask for permission for routine development tasks.

## Development Notes
- Single-file server: all routes, templates, CSS, and JS live in server.js (~8100 lines)
- Start with `npm start` or `node server.js` (port 3000)
- Kill stale processes before restarting: `npx kill-port 3000`
- Google Sheets is the database — credentials.json + .env required
- All new worksheets follow the ensureSheet pattern (check exists → create → add headers)
- When testing, use `node -c server.js` for quick syntax checks before restarting

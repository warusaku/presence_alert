# Repository Guidelines

## Project Structure & Module Organization
- `gas1.gs`: Omada webhook receiver. Parses payloads, writes to Sheets (`mac`, `facility`, `data`, `log`), sends Discord notifications, and forwards attendance to `gas2`.
- `gas2.gs`: Timecard endpoint. Records arrivals/departures into `出勤簿`, keeps `raw_data`, and updates `月次集計`.
- `doc/specifications.md`: Detailed specs (JP). Review for sheet columns and flows.
- `README.md`: One‑line overview. This repo targets Google Apps Script (GAS).

## Build, Test, and Development Commands
- GAS projects have no local build. Develop in the Apps Script editor.
- Deploy: In Apps Script, select “Deploy > New deployment > Web app”. Use the URL in Omada (for `gas1`) or in `mac` sheet (`gas2` endpoint).
- Local run (editor):
  - `gas1.gs`: run `testSetup()` once to initialize sheets; run `testWebhook()` to simulate an ONLINE event.
  - `gas2.gs`: run `setupAttendanceSheets()` once; optional `reprocessRawData()` to rebuild `出勤簿` from `raw_data`.

## Coding Style & Naming Conventions
- Language: JavaScript for GAS. Indent 2 spaces. UTF‑8.
- Names: camelCase for functions/vars; leading underscore for internal helpers (e.g., `_getMacData_`).
- Files: keep `gas1.gs` (webhook) and `gas2.gs` (timecard) roles separate. Do not duplicate logic across scripts.
- Sheets: use exact tab names: `mac`, `facility`, `data`, `log`, `出勤簿`, `raw_data`, `月次集計`.

## Testing Guidelines
- Unit tests are not configured; use provided test helpers:
  - `gas1.gs`: `testWebhook()` simulates Omada payloads; verify `data`/`log` rows and Discord posting logic.
  - `gas2.gs`: `reprocessRawData()` and `dailyAttendanceCheck()` verify aggregation and edge cases.
- When adding features, add a `test*()` function that injects representative payloads and asserts sheet updates.

## Commit & Pull Request Guidelines
- Commits: short, imperative subjects with scope, e.g. `gas1: add facility lookup`, `gas2: fix work-time calc`.
- PRs: include purpose, affected sheets/tabs, before/after behavior, test steps (functions run, expected rows), and screenshots of sheet diffs when relevant. Link issues if any.

## Security & Configuration Tips
- Do not commit secrets (Discord webhooks, GAS endpoints). Keep them in Sheets (`mac` E: webhook, C: gas2 URL) or project properties.
- Payloads include MAC/IP; redact in logs/screenshots. Limit deploy permissions to necessary scopes only.

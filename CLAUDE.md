# Project Rules

## Testing requirement — ALWAYS before saying "ready"

Before telling the user any feature or fix is complete, you MUST:

1. Make sure the Flask app is running:
   ```bash
   lsof -t -i:5000 | xargs kill 2>/dev/null; sleep 1
   cd schedule_comparison && source .env && .venv/bin/python app.py > /tmp/sce_app.log 2>&1 &
   sleep 3
   ```

2. Run the full Playwright test suite from the repo root:
   ```bash
   npx playwright test --reporter=list
   ```

3. Also run the Python unit tests:
   ```bash
   cd schedule_comparison && python -m pytest tests/ -v
   ```

4. All tests must pass before reporting the work as done. If any test fails, fix the issue first.

## Repo layout

- `schedule_comparison/` — Flask web app (Python)
- `tests/e2e/` — Playwright end-to-end tests (TypeScript)
- `playwright.config.ts` — Playwright config (baseURL: localhost:5000, uses system Chromium at `/opt/pw-browsers/chromium-1194/`)

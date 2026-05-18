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

---

# Karpathy Guidelines

Behavioral guidelines to reduce common LLM coding mistakes.

**Tradeoff:** These guidelines bias toward caution over speed. For trivial tasks, use judgment.

## 1. Think Before Coding

**Don't assume. Don't hide confusion. Surface tradeoffs.**

Before implementing:
- State your assumptions explicitly. If uncertain, ask.
- If multiple interpretations exist, present them - don't pick silently.
- If a simpler approach exists, say so. Push back when warranted.
- If something is unclear, stop. Name what's confusing. Ask.

## 2. Simplicity First

**Minimum code that solves the problem. Nothing speculative.**

- No features beyond what was asked.
- No abstractions for single-use code.
- No "flexibility" or "configurability" that wasn't requested.
- No error handling for impossible scenarios.
- If you write 200 lines and it could be 50, rewrite it.

Ask yourself: "Would a senior engineer say this is overcomplicated?" If yes, simplify.

## 3. Surgical Changes

**Touch only what you must. Clean up only your own mess.**

When editing existing code:
- Don't "improve" adjacent code, comments, or formatting.
- Don't refactor things that aren't broken.
- Match existing style, even if you'd do it differently.
- If you notice unrelated dead code, mention it - don't delete it.

When your changes create orphans:
- Remove imports/variables/functions that YOUR changes made unused.
- Don't remove pre-existing dead code unless asked.

The test: Every changed line should trace directly to the user's request.

## 4. Goal-Driven Execution

**Define success criteria. Loop until verified.**

Transform tasks into verifiable goals:
- "Add validation" → "Write tests for invalid inputs, then make them pass"
- "Fix the bug" → "Write a test that reproduces it, then make it pass"
- "Refactor X" → "Ensure tests pass before and after"

For multi-step tasks, state a brief plan:
```
1. [Step] → verify: [check]
2. [Step] → verify: [check]
3. [Step] → verify: [check]
```

Strong success criteria let you loop independently. Weak criteria ("make it work") require constant clarification.

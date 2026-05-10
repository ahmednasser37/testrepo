# Schedule Comparison App

A Flask web application that accepts two Primavera P6 XER files (baseline + updated), compares them, and renders an interactive HTML dashboard with an AI-written narrative via OpenRouter.

## Prerequisites

Python 3.11+ recommended.

```bash
cd schedule_comparison
pip install -r requirements.txt
```

## Run

```bash
python app.py
# Open http://localhost:5000
```

Or with Flask dev server:

```bash
FLASK_APP=app.py FLASK_DEBUG=1 flask run
```

## Environment Variables

| Variable | Default | Description |
|---|---|---|
| `OPENROUTER_API_KEY` | *(unset)* | API key for AI narrative. If unset, a stub message is shown. |
| `OPENROUTER_MODEL` | `deepseek/deepseek-chat-v3-0324:free` | OpenRouter model ID |
| `SCE_CACHE_DIR` | `/tmp/sce_cache` | Directory for AI response and export caches |
| `SCE_UPLOAD_MAX_MB` | `50` | Maximum upload size per file (MB) |
| `PORT` | `5000` | HTTP listen port |

Copy `.env.example` to `.env` and fill in your key, then `source .env` before starting.

## Run Tests

```bash
python -m pytest tests/ -v
```

## Usage

1. Open the app and upload two `.xer` files — label them *Baseline* and *Updated*.
2. Click **Compare** — the dashboard renders immediately.
3. Use the export buttons to download the activity variance table as **CSV** or the full comparison as **JSON**.

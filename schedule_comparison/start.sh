#!/usr/bin/env bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Load env vars if .env exists
if [ -f .env ]; then
  set -o allexport
  source .env
  set +o allexport
fi

PORT="${PORT:-5000}"
WORKERS="${WORKERS:-2}"
TIMEOUT="${TIMEOUT:-120}"
LOG_FILE="${LOG_FILE:-/tmp/sce_app.log}"

mkdir -p "${SCE_CACHE_DIR:-/tmp/sce_cache}"

exec .venv/bin/gunicorn \
  --bind "0.0.0.0:${PORT}" \
  --workers "${WORKERS}" \
  --timeout "${TIMEOUT}" \
  --access-logfile "${LOG_FILE}" \
  --error-logfile "${LOG_FILE}" \
  app:app

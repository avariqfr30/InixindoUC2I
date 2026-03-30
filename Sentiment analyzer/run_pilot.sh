#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

export APP_MODE="${APP_MODE:-demo}"
export HOST="${HOST:-0.0.0.0}"
export PORT="${PORT:-8000}"
export WAITRESS_THREADS="${WAITRESS_THREADS:-8}"
export WAITRESS_CONNECTION_LIMIT="${WAITRESS_CONNECTION_LIMIT:-100}"
export WAITRESS_CHANNEL_TIMEOUT="${WAITRESS_CHANNEL_TIMEOUT:-240}"
export REPORT_JOB_WORKERS="${REPORT_JOB_WORKERS:-3}"
export REPORT_MAX_PENDING_JOBS="${REPORT_MAX_PENDING_JOBS:-24}"
export REPORT_JOB_RETENTION_SECONDS="${REPORT_JOB_RETENTION_SECONDS:-86400}"
export OSINT_CACHE_TTL_SECONDS="${OSINT_CACHE_TTL_SECONDS:-21600}"
export RESEED_DEMO_DATA="${RESEED_DEMO_DATA:-0}"

if [[ "$APP_MODE" == "demo" && "$RESEED_DEMO_DATA" == "1" ]]; then
  python3 seed_demo_data.py
fi

exec python3 -m waitress \
  --host="$HOST" \
  --port="$PORT" \
  --threads="$WAITRESS_THREADS" \
  --connection-limit="$WAITRESS_CONNECTION_LIMIT" \
  --channel-timeout="$WAITRESS_CHANNEL_TIMEOUT" \
  app:app

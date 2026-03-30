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

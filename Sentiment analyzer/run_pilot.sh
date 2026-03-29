#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

export APP_MODE="${APP_MODE:-demo}"
export PILOT_SERVER="${PILOT_SERVER:-waitress}"
export HOST="${HOST:-0.0.0.0}"
export PORT="${PORT:-8000}"
export PILOT_THREADS="${PILOT_THREADS:-8}"
export PILOT_CONNECTION_LIMIT="${PILOT_CONNECTION_LIMIT:-100}"
export PILOT_CHANNEL_TIMEOUT="${PILOT_CHANNEL_TIMEOUT:-240}"
export GUNICORN_WORKERS="${GUNICORN_WORKERS:-1}"
export GUNICORN_THREADS="${GUNICORN_THREADS:-8}"
export GUNICORN_TIMEOUT="${GUNICORN_TIMEOUT:-240}"

if [[ "$PILOT_SERVER" == "waitress" ]]; then
  exec python3 pilot_server.py
fi

if [[ "$PILOT_SERVER" == "gunicorn" ]]; then
  if command -v gunicorn >/dev/null 2>&1; then
    exec gunicorn --config gunicorn.conf.py wsgi:app
  fi

  exec python3 -m gunicorn --config gunicorn.conf.py wsgi:app
fi

exec python3 app.py

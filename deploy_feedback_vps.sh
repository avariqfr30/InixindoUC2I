#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 1 ]]; then
  echo "Usage: $0 /path/to/ssh_key.pem"
  exit 1
fi

KEY_PATH="$1"
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOCAL_APP_DIR="$ROOT_DIR/Sentiment analyzer/"
REMOTE_HOST="${REMOTE_HOST:-ubuntu@18.136.190.197}"
REMOTE_APP_DIR="${REMOTE_APP_DIR:-/opt/apps/inixindo-feedback/current}"
REMOTE_BASE_DIR="${REMOTE_BASE_DIR:-/opt/apps/inixindo-feedback}"
REMOTE_VENV_DIR="${REMOTE_VENV_DIR:-/opt/apps/inixindo-feedback/venv}"
SERVICE_NAME="${SERVICE_NAME:-inixindo-feedback}"
PUBLIC_URL="${PUBLIC_URL:-https://feedback.inworx.id}"
LOCAL_HEALTH_URL="${LOCAL_HEALTH_URL:-http://127.0.0.1:6001/health}"
LOCAL_READY_URL="${LOCAL_READY_URL:-http://127.0.0.1:6001/ready}"
SSH_OPTS=(-i "$KEY_PATH" -o StrictHostKeyChecking=accept-new)

if [[ ! -d "$LOCAL_APP_DIR" ]]; then
  echo "Local app directory not found: $LOCAL_APP_DIR"
  exit 1
fi

ssh "${SSH_OPTS[@]}" "$REMOTE_HOST" "mkdir -p '$REMOTE_APP_DIR' '$REMOTE_BASE_DIR'"

rsync -avz --delete \
  --exclude '.DS_Store' \
  --exclude '__pycache__/' \
  --exclude '.pytest_cache/' \
  --exclude '.mypy_cache/' \
  --exclude 'profiles/*.env' \
  --exclude 'internal_connector.production.json' \
  --exclude 'data/auth.db' \
  --exclude 'data/cx_feedback.db' \
  --exclude 'data/report_jobs.json' \
  --exclude 'data/osint_cache.json' \
  --exclude 'data/generated_reports/' \
  -e "ssh ${SSH_OPTS[*]}" \
  "$LOCAL_APP_DIR" \
  "$REMOTE_HOST:$REMOTE_APP_DIR/"

ssh "${SSH_OPTS[@]}" "$REMOTE_HOST" "
  set -e
  python3 -m venv '$REMOTE_VENV_DIR'
  source '$REMOTE_VENV_DIR/bin/activate'
  cd '$REMOTE_APP_DIR'
  pip install -r requirements.txt >/tmp/${SERVICE_NAME}_pip.log 2>&1 || { cat /tmp/${SERVICE_NAME}_pip.log; exit 1; }
  sudo systemctl restart '$SERVICE_NAME'
  sudo systemctl status '$SERVICE_NAME' --no-pager -l | sed -n '1,20p'
  for _ in \$(seq 1 30); do
    if curl -fsS '$LOCAL_HEALTH_URL' >/tmp/${SERVICE_NAME}_health.json 2>/dev/null; then
      cat /tmp/${SERVICE_NAME}_health.json
      echo
      break
    fi
    sleep 1
  done
  test -f /tmp/${SERVICE_NAME}_health.json
  echo
  for _ in \$(seq 1 30); do
    if curl -fsS '$LOCAL_READY_URL' >/tmp/${SERVICE_NAME}_ready.json 2>/dev/null; then
      cat /tmp/${SERVICE_NAME}_ready.json
      echo
      break
    fi
    sleep 1
  done
  test -f /tmp/${SERVICE_NAME}_ready.json
  echo
"

curl -fsSI "$PUBLIC_URL" | sed -n '1,12p'

echo
echo "Deployment complete for $PUBLIC_URL"

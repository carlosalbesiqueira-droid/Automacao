#!/usr/bin/env bash
set -euo pipefail

export PORT="${PORT:-3210}"
export BOT_FATURAS_API_HOST="${BOT_FATURAS_API_HOST:-127.0.0.1}"
export BOT_FATURAS_API_PORT="${BOT_FATURAS_API_PORT:-8321}"
export BOT_FATURAS_API_BASE="${BOT_FATURAS_API_BASE:-http://127.0.0.1:${BOT_FATURAS_API_PORT}}"
export BOT_FATURAS_STORAGE_DIR="${BOT_FATURAS_STORAGE_DIR:-/data/storage/faturas}"
export BOT_FATURAS_DB_PATH="${BOT_FATURAS_DB_PATH:-/data/storage/faturas/bot_faturas.db}"

mkdir -p "$(dirname "$BOT_FATURAS_DB_PATH")"

python scripts/run_bot_faturas_api.py &
API_PID=$!

node src/server.js &
WEB_PID=$!

cleanup() {
  kill "$API_PID" "$WEB_PID" 2>/dev/null || true
}

trap cleanup INT TERM EXIT

wait -n "$API_PID" "$WEB_PID"
STATUS=$?
cleanup
wait "$API_PID" "$WEB_PID" 2>/dev/null || true
exit "$STATUS"

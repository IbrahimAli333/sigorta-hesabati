#!/usr/bin/env bash
# start.sh — Run the dev server, killing any leftover instance first.
# Usage: ./start.sh
set -euo pipefail

cd "$(dirname "$0")"

# Free port 5050 if something is holding it.
PIDS=$(lsof -ti :5050 2>/dev/null || true)
if [ -n "$PIDS" ]; then
  echo "→ Freeing port 5050 (killing PID $PIDS)..."
  echo "$PIDS" | xargs kill 2>/dev/null || true
  sleep 1
fi

echo "→ Starting Flask dev server on http://127.0.0.1:5050"
exec python3 app.py

#!/usr/bin/env bash
set -euo pipefail

# Run FastAPI backend and Vite frontend together
ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT_DIR"

echo "Starting backend (uvicorn) on http://localhost:8000 …"
( source .venv/bin/activate && uvicorn backend.main:app --reload --port 8000 ) &
BACK_PID=$!

echo "Starting frontend (Vite) on http://localhost:5173 …"
( cd frontend && npm run dev ) &
FRONT_PID=$!

cleanup() {
  echo "\nStopping services…"
  kill "$BACK_PID" 2>/dev/null || true
  kill "$FRONT_PID" 2>/dev/null || true
}
trap cleanup INT TERM EXIT

wait

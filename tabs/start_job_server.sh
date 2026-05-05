#!/usr/bin/env sh
set -eu

JOB_HOST="${JOB_HOST:-0.0.0.0}"
JOB_PORT="${JOB_PORT:-8765}"

cd "$(dirname "$0")"
python3 network_job_server.py --host "$JOB_HOST" --port "$JOB_PORT"

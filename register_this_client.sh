#!/usr/bin/env sh
set -eu

# Edit SERVER_URL to match your server.
SERVER_URL="${SERVER_URL:-http://192.168.1.50:8765}"
CLIENT_NAME="${CLIENT_NAME:-$(hostname)}"

cd "$(dirname "$0")"
python3 register_client.py --server "$SERVER_URL" --name "$CLIENT_NAME"

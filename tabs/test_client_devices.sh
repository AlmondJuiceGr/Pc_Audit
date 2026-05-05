#!/usr/bin/env sh
set -eu
cd "$(dirname "$0")"
python3 register_client.py --print-diagnostics

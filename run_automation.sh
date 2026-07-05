#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# run_automation.sh
# Quick launcher for blog_image_automation.py
#
# Usage:
#   1. Set your credentials in .env (copy from .env.example)
#   2. Run:  bash run_automation.sh
# ─────────────────────────────────────────────────────────────────────────────

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Load credentials from .env if it exists
if [ -f ".env" ]; then
    echo "[setup] Loading credentials from .env"
    set -a
    source .env
    set +a
else
    echo "[setup] No .env file found. Copy .env.example to .env and fill in your credentials."
    echo "        Or set WP_USER, WP_APP_PASS, OPENAI_API_KEY as environment variables."
    exit 1
fi

# Validate required vars
if [ -z "$WP_USER" ] || [ -z "$WP_APP_PASS" ]; then
    echo "[ERROR] WP_USER and WP_APP_PASS must be set in .env"
    exit 1
fi

echo "[setup] WordPress user: $WP_USER"
echo "[setup] OpenAI key:     ${OPENAI_API_KEY:0:10}..."
echo ""
echo "Starting automation..."
python3 blog_image_automation.py "$@"

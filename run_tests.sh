#!/bin/bash
# ─────────────────────────────────────────────
#  Portfolio Dashboard — Test Runner
#  Usage: bash run_tests.sh
# ─────────────────────────────────────────────

DIR="$(cd "$(dirname "$0")" && pwd)"

echo ""
echo "Checking dependencies..."

# Check Node.js
if ! command -v node &> /dev/null; then
  echo "❌ Node.js not found. Please install Node.js from https://nodejs.org"
  exit 1
fi

# Install jsdom if needed
if [ ! -d "$DIR/node_modules/jsdom" ]; then
  echo "Installing jsdom (first time only)..."
  cd "$DIR" && npm install jsdom --save-dev --silent
fi

echo "Running tests..."
echo ""
node "$DIR/test_dashboard.js"

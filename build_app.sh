#!/bin/bash
# Builds a self-contained "Prism Extractor.app" in dist/ (no Python needed to run it).
# Usage: ./build_app.sh [universal2|arm64|x86_64]   (default: universal2)
set -euo pipefail
cd "$(dirname "$0")"
ARCH="${1:-universal2}"
PY="${PYTHON:-.venv/bin/python}"

"$PY" -m pip install -q -r requirements.txt pyinstaller
# markupsafe._speedups ships as a single-arch binary; markupsafe falls back to pure Python without it.
"$PY" -m PyInstaller --noconfirm --clean --windowed \
  --name "Prism Extractor" \
  --icon icon/AppIcon.icns \
  --add-data "icon/AppIcon.png:icon" \
  --osx-bundle-identifier io.github.markusviks.prism-extractor \
  --target-arch "$ARCH" \
  --exclude-module markupsafe._speedups \
  app.py
echo "Built: dist/Prism Extractor.app"

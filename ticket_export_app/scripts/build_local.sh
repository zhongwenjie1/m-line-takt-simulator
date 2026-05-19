#!/usr/bin/env bash
set -e

cd "$(dirname "$0")/.."

MODE="${1:-check}"
PYTHON_BIN=""
PYINSTALLER_AVAILABLE=0

usage() {
  echo "Usage: $0 [check|build]"
}

select_python() {
  if [ -n "${PYTHON:-}" ]; then
    PYTHON_BIN="$PYTHON"
  elif [ -x "../.venv312/bin/python" ]; then
    PYTHON_BIN="../.venv312/bin/python"
  else
    PYTHON_BIN="python3"
  fi
}

check_environment() {
  echo "== ticket_export_app build environment check =="
  echo "Project dir: $(pwd)"

  if [ ! -f "version.json" ]; then
    echo "ERROR: version.json not found."
    exit 1
  fi

  echo ""
  echo "== version.json =="
  cat version.json
  echo ""

  if [ ! -f "main.py" ]; then
    echo "ERROR: main.py not found."
    exit 1
  fi

  if [ ! -f "packaging/ticket_export_app.spec" ]; then
    echo "ERROR: packaging/ticket_export_app.spec not found."
    exit 1
  fi

  select_python

  if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
    echo "ERROR: Python executable not found: $PYTHON_BIN"
    echo "You can set PYTHON=/path/to/python before running this script."
    exit 1
  fi

  echo ""
  echo "== Python =="
  "$PYTHON_BIN" --version

  echo ""
  echo "== PyInstaller =="
  if PYINSTALLER_VERSION="$("$PYTHON_BIN" -m PyInstaller --version 2>/tmp/ticket_export_pyinstaller_error.txt)"; then
    PYINSTALLER_AVAILABLE=1
    echo "PyInstaller available: $PYINSTALLER_VERSION"
  else
    PYINSTALLER_AVAILABLE=0
    echo "PyInstaller is not installed or not available in this Python environment."
    echo "Install later with:"
    echo "  $PYTHON_BIN -m pip install pyinstaller"
  fi

  echo ""
  echo "Build environment check completed."
}

run_build() {
  check_environment

  if [ "$PYINSTALLER_AVAILABLE" -ne 1 ]; then
    echo "ERROR: PyInstaller is required for build mode."
    exit 1
  fi

  "$PYTHON_BIN" -m PyInstaller packaging/ticket_export_app.spec --clean --noconfirm
  echo "Build completed."
}

case "$MODE" in
  check)
    check_environment
    ;;
  build)
    run_build
    ;;
  *)
    usage
    exit 1
    ;;
esac

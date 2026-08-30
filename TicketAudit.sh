#!/usr/bin/env bash
set -euo pipefail

# Light green on black, matching the brand's green-on-chrome (#3fc16e on #0f1518).
GREEN='\033[0;32m'
BOLD_GREEN='\033[1;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m'

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# Print the FIGlet banner (same file the .bat uses).
if [ -f "$SCRIPT_DIR/banner.txt" ]; then
    printf "${BOLD_GREEN}"
    cat "$SCRIPT_DIR/banner.txt"
    printf "${NC}"
fi

# ---------------------------------------------------------------------------
# Locate Python 3
# ---------------------------------------------------------------------------
PYTHON=""
for cmd in python3 python; do
    if command -v "$cmd" >/dev/null 2>&1; then
        version=$("$cmd" --version 2>&1)
        major=$("$cmd" -c "import sys; print(sys.version_info.major)" 2>/dev/null || echo 0)
        if [ "$major" -ge 3 ]; then
            PYTHON="$cmd"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    printf "${RED}[ERROR]${NC} Python 3 is not installed or not in PATH.\n"
    printf "        Please install Python 3.10+ from https://python.org\n"
    exit 1
fi

minor=$("$PYTHON" -c "import sys; print(sys.version_info.minor)" 2>/dev/null || echo 0)
if [ "$minor" -lt 10 ]; then
    printf "${RED}[ERROR]${NC} Python 3.10+ is required (found %s).\n" "$("$PYTHON" --version 2>&1)"
    printf "        Install a newer version from https://python.org\n"
    printf "        or via your package manager (brew install python, etc.)\n"
    exit 1
fi

printf "${GREEN}[OK]${NC} Python found: %s\n\n" "$("$PYTHON" --version 2>&1)"

# ---------------------------------------------------------------------------
# Ensure main.py exists next to this script
# ---------------------------------------------------------------------------
if [ ! -f "$SCRIPT_DIR/main.py" ]; then
    printf "${RED}[ERROR]${NC} main.py not found next to this script.\n"
    printf "        Please run this script from the TicketAudit folder.\n"
    exit 1
fi

cd "$SCRIPT_DIR"

# ---------------------------------------------------------------------------
# Virtual environment
# ---------------------------------------------------------------------------
if [ ! -d ".venv" ]; then
    printf "${YELLOW}[INFO]${NC} Creating virtual environment...\n"
    "$PYTHON" -m venv .venv
    printf "${GREEN}[OK]${NC} Virtual environment created.\n\n"
fi

# Activate (POSIX venvs use bin/, not Scripts/)
# shellcheck disable=SC1091
source ".venv/bin/activate"

# ---------------------------------------------------------------------------
# Install dependencies if the sentinel import is missing
# ---------------------------------------------------------------------------
if ! python -c "import PySide6" >/dev/null 2>&1; then
    printf "${YELLOW}[INFO]${NC} Installing dependencies in virtual environment...\n"
    printf "        This may take a few minutes on first run...\n\n"
    python -m pip install --upgrade pip >/dev/null 2>&1
    if ! python -m pip install -r "$SCRIPT_DIR/requirements.txt"; then
        printf "\n${RED}[ERROR]${NC} Failed to install dependencies.\n"
        printf "        Try running: pip install -r requirements.txt\n"
        exit 1
    fi
    printf "\n${GREEN}[OK]${NC} Dependencies installed successfully.\n\n"
fi

# ---------------------------------------------------------------------------
# Launch
# ---------------------------------------------------------------------------
printf "${YELLOW}[INFO]${NC} Starting TicketAudit...\n\n"
python main.py

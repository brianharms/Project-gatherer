#!/bin/bash
# Double-click this file on a Mac to launch Project Gatherer.
# (If macOS blocks it the first time: right-click -> Open, or run
#  `chmod +x "Run on Mac.command"` in Terminal once.)

# Move to the folder this launcher lives in, so the app finds its files
# even when started from an external drive.
cd "$(dirname "$0")" || exit 1

# Find a Python 3 interpreter.
PY=""
for cand in python3 python; do
    if command -v "$cand" >/dev/null 2>&1; then
        if "$cand" -c 'import sys; sys.exit(0 if sys.version_info[0]==3 else 1)' >/dev/null 2>&1; then
            PY="$cand"
            break
        fi
    fi
done

if [ -z "$PY" ]; then
    echo ""
    echo "  Python 3 was not found on this Mac."
    echo ""
    echo "  Recent macOS versions include python3, but you may need to install"
    echo "  the Command Line Tools or Python from https://www.python.org/downloads/"
    echo ""
    echo "  Press return to close."
    read -r _
    exit 1
fi

echo "Starting Project Gatherer with: $($PY --version 2>&1)"
"$PY" project_gatherer.py
STATUS=$?

if [ $STATUS -ne 0 ]; then
    echo ""
    echo "  The app exited with an error (code $STATUS)."
    echo "  Press return to close."
    read -r _
fi

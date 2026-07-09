#!/bin/sh
cd "$(dirname "$0")" || exit 1

if command -v python3 >/dev/null 2>&1; then
    python3 gel_gui.py
else
    python gel_gui.py
fi

status=$?
if [ "$status" -ne 0 ]; then
    echo
    echo "GelBandPicker exited with an error. Press Enter to close this window."
    read -r _
fi
exit "$status"

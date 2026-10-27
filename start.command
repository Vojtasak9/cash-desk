#!/bin/bash
# Double-click on macOS to start the cash desk
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
fi
.venv/bin/python app.py

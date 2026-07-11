#!/bin/sh
cd "$(dirname "$0")"
[ -d venv ] || python3 -m venv venv
. venv/bin/activate
python -m pip install -q -r requirements.txt
python -m shift_app.web

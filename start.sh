#!/bin/sh
cd "$(dirname "$0")"
pip install -r requirements.txt
python3 -m shift_app.web

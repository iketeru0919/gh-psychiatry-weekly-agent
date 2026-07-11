@echo off
cd /d %~dp0
pip install -r requirements.txt
python -m shift_app.web
pause

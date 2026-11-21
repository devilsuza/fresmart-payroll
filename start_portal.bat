@echo off
cd /d %~dp0
if not exist data mkdir data
set DATA_DIR=%CD%\data
if not defined PORT set PORT=8080
python -m pip install -r requirements.txt
python app.py
pause

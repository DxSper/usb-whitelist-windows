@echo off

REM Get the directory of the running batch file
SET script_dir=%~dp0

REM Change to that directory
cd /d %script_dir%
start "" pythonw usb-whitelist.py
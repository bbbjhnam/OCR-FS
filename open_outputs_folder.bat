@echo off
cd /d "%~dp0"
if not exist outputs mkdir outputs
start "" "%~dp0outputs"

@echo off
cd /d "%~dp0"
echo Sinkronisasi Excel ke GitHub Pages tiap 60 detik. Biarkan jendela ini terbuka (Ctrl+C untuk berhenti).
py -3 sync_github.py
pause

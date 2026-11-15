@echo off
setlocal
cd /d "%~dp0"

py -3 -c "import openpyxl" >nul 2>&1
if not errorlevel 1 (
  py -3 server.py
  goto :end
)

python -c "import openpyxl" >nul 2>&1
if not errorlevel 1 (
  python server.py
  goto :end
)

set "CODEX_PYTHON=C:\Program Files\WindowsApps\OpenAI.CodexPrimaryRuntime.v26-1112-737-0_26.1112.737.0_x64__3k8sg7r9htsxt\dependencies\python\python.exe"
if exist "%CODEX_PYTHON%" (
  "%CODEX_PYTHON%" server.py
  goto :end
)

echo Python dengan openpyxl tidak ditemukan.
echo Jalankan dari Codex atau pasang Python dan openpyxl, lalu coba lagi.
pause

:end
endlocal

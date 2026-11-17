@echo off
rem Pasang sekali: sinkron ke GitHub jalan otomatis tersembunyi setiap Windows login.
setlocal
set "DIR=%~dp0"
set "VBS=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\SyncDashboardMSFO.vbs"

> "%VBS%" echo Set sh = CreateObject("WScript.Shell")
>> "%VBS%" echo sh.CurrentDirectory = "%DIR%"
>> "%VBS%" echo sh.Run "pyw -3 sync_github.py", 0, False

wscript "%VBS%"
echo.
echo Terpasang. Sinkronisasi berjalan tersembunyi sekarang dan otomatis tiap login.
echo Log: %DIR%sync.log
echo Untuk mematikan: hapus file "%VBS%" lalu akhiri proses pythonw di Task Manager.
pause
endlocal

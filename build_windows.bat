@echo off
setlocal

where pycrucible >nul 2>nul
if errorlevel 1 (
  echo PyCrucible belum terpasang.
  echo Jalankan: py -m pip install pycrucible
  exit /b 1
)

if exist dist rmdir /s /q dist
mkdir dist

pycrucible -e . -o .\dist\HR_Attendance_Dashboard.exe
if errorlevel 1 exit /b 1

echo.
echo Build selesai: dist\HR_Attendance_Dashboard.exe
endlocal

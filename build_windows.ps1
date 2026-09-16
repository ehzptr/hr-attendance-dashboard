$ErrorActionPreference = 'Stop'

if (-not (Get-Command pycrucible -ErrorAction SilentlyContinue)) {
    Write-Host 'PyCrucible belum terpasang.' -ForegroundColor Yellow
    Write-Host 'Jalankan: py -m pip install pycrucible'
    exit 1
}

if (Test-Path .\dist) {
    Remove-Item .\dist -Recurse -Force
}

New-Item .\dist -ItemType Directory | Out-Null
pycrucible -e . -o .\dist\HR_Attendance_Dashboard.exe

Write-Host ''
Write-Host 'Build selesai:' -ForegroundColor Green
Write-Host '.\dist\HR_Attendance_Dashboard.exe'

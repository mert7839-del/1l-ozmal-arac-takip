@echo off
cd /d "%~dp0"
echo Eski Node sunuculari kapatiliyor...
taskkill /F /IM node.exe >nul 2>nul

echo Bozuk node_modules temizleniyor...
if exist node_modules rmdir /s /q node_modules

echo Node paketleri bastan kuruluyor...
npm install
if errorlevel 1 (
  echo.
  echo NPM paket kurulumu basarisiz oldu. Internet baglantisini ve Node.js kurulumunu kontrol et.
  pause
  exit /b 1
)

echo Sistem baslatiliyor...
npm start
pause

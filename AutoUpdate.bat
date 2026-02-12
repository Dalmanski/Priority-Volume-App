@echo off
cd "%USERPROFILE%\OneDrive\Documents\Code\PyQt5\Priority Volume App\PrioVolumeApp.py"

set "VERSION=1.0.0"

echo === Building executable with PyInstaller ===
pyinstaller ^
  "PrioVolumeApp.py"

echo === Creating ZIP archive ===
powershell -NoProfile -Command ^
  "Compress-Archive -Path 'dist\\*' -DestinationPath 'zip\\Priority_Volume_App_%VERSION%.zip' -Force"

echo === Done! ===
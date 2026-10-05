@echo off
chcp 65001 >nul
echo Configurando os videos das apresentacoes para esta pasta...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build\configurar_videos.ps1" -Pasta "%~dp0."
echo.
pause

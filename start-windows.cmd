@echo off
REM SlideTeX - lanceur Windows : installe les dependances au premier lancement puis demarre.
cd /d "%~dp0"
where node >nul 2>nul || (echo Node.js est requis : https://nodejs.org & pause & exit /b 1)
if not exist node_modules (
  echo Installation des dependances...
  call npm install || (pause & exit /b 1)
)
call npm start
pause

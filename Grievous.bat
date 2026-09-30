@echo off
rem Alternativa ao Grievous.exe, para quem prefere rodar pelo .bat.
title Grievous
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js nao encontrado. Instale em https://nodejs.org e rode novamente.
  pause
  exit /b 1
)

if not exist "node_modules" (
  echo Primeira execucao: instalando dependencias...
  call npm run setup
)

node iniciar.js
pause

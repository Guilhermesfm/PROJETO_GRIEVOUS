@echo off
rem Sobe o Grievous e abre o navegador na tela de login.
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

echo Iniciando o Grievous...
start "" /b node server.js

rem Espera o servidor responder antes de abrir o navegador.
for /l %%i in (1,1,20) do (
  timeout /t 1 /nobreak >nul
  curl -s -o nul http://localhost:3000/ && goto :abrir
)

:abrir
start "" http://localhost:3000/
echo.
echo Grievous rodando em http://localhost:3000
echo Feche esta janela para encerrar.
pause >nul

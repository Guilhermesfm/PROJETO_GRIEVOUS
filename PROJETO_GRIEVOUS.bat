@echo off
echo Iniciando a automação...
npx playwright test grievous/example.spec.js --headed
pause

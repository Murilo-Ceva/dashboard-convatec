@echo off
rem Salva a ultima versao do codigo (macros, painel, scripts) no Git. Nao leva dados.
rem Para tambem enviar ao GitHub: arraste para o cmd e acrescente "enviar", ou rode: "SALVAR VERSAO NO GIT.bat" enviar
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "scripts\salvar_versao.ps1" %1
pause

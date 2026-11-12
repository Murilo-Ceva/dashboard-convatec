@echo off
rem Atualiza o painel ConvaTec: exporta a base da planilha, gera os paineis, confere e copia para o OneDrive.
rem Antes: salve a planilha "Base CONVATEC - com tracking.xlsm".
cd /d "%~dp0"
echo 1/3 Exportando a base da planilha...
powershell -NoProfile -ExecutionPolicy Bypass -File "scripts\exportar_base.ps1" || goto erro
echo 2/3 Gerando os paineis...
python -I scripts\gerar_painel.py || goto erro
echo 3/3 Conferindo...
python -I scripts\verificar.py || goto erro
echo.
echo Pronto. Os paineis estao no SharePoint SIT - Gestao Operacional, pasta CONVATEC\PAINEL.
pause
exit /b 0
:erro
echo.
echo ERRO: o painel NAO foi atualizado. Leia a mensagem acima.
pause
exit /b 1

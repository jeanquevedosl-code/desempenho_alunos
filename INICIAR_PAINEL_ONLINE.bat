@echo off
cd /d "%~dp0"
title Painel CGD (online)
echo Iniciando o painel e o link publico...
echo Deixe esta janela aberta. Ctrl+C encerra.
echo.
python painel_online.py
echo.
echo O painel foi encerrado.
pause

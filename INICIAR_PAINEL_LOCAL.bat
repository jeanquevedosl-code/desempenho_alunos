@echo off
cd /d "%~dp0"
title Painel CGD (local)
echo Iniciando o painel apenas neste computador...
echo Endereco: http://127.0.0.1:8000
echo Deixe esta janela aberta. Ctrl+C encerra.
echo.
start "" "http://127.0.0.1:8000"
python painel.py
echo.
echo O painel foi encerrado.
pause

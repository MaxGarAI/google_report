@echo off
chcp 65001 > nul
title Генератор отчетов в Google Sheets

echo Запуск приложения...
if exist venv\Scripts\python.exe (
    start "" venv\Scripts\pythonw.exe app.py
) else (
    start "" pythonw app.py
)
exit

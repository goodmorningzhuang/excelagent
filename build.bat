@echo off
cd /d D:\excelagent-exe\exe
pyinstaller --onefile --name ExcelAgent --add-data "static;static" --add-data "skill;skill" --add-data "ai_skills;ai_skills" --add-data "使用说明.txt;." --hidden-import openpyxl --hidden-import xlrd --hidden-import flask --hidden-import flask_cors --hidden-import matplotlib --hidden-import requests --hidden-import anthropic --hidden-import openai --noconfirm app.py
echo.
echo ===== Build Complete =====
echo Output: D:\excelagent-exe\exe\dist\ExcelAgent.exe
pause

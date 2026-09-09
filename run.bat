@echo off
chcp 65001 >nul
title VIPQC AI
python app.py
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Ung dung ket thuc voi ma loi %ERRORLEVEL%
    pause
)

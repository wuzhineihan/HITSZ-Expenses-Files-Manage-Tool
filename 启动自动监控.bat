@echo off
chcp 65001 >nul
title 社团报销自动化监控
color 0A

REM 设置Python输出编码为UTF-8
set PYTHONIOENCODING=utf-8

echo.
echo ========================================
echo    社团报销自动化监控系统
echo ========================================
echo.
echo 正在启动监控...
echo.

C:\Users\Dau\AppData\Local\Programs\Python\Python311\python.exe auto_watch.py

pause

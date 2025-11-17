@echo off
chcp 65001 >nul
title 社团报销文件夹管理
color 0B

REM 设置Python输出编码为UTF-8
set PYTHONIOENCODING=utf-8

echo.
echo ========================================
echo    社团报销文件夹管理脚本
echo ========================================
echo.
echo 正在执行...
echo.

C:\Users\Dau\AppData\Local\Programs\Python\Python311\python.exe create_folders.py

echo.
echo ========================================
echo    执行完成
echo ========================================
echo.

pause

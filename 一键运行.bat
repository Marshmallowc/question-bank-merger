@echo off
chcp 65001 >nul
title 题库合并工具 - 新手版

echo ========================================
echo 题库合并工具 - 新手版
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未检测到Python！
    echo 请先安装Python：https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

REM 运行Python脚本
echo 正在启动题库合并工具...
echo.
python run.py

pause
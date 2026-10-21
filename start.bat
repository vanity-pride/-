@echo off
title 每日打卡记录系统 - 服务器（关闭本窗口 = 停止）

REM ============================================================
REM  每日打卡记录系统 · 一键启动
REM ------------------------------------------------------------
REM  双击本文件即可启动系统并自动打开浏览器。
REM  本黑窗口就是“服务器”：保持开着才能用，关掉它就停止。
REM ============================================================

REM ---------- 配置 ----------
set "PROJECT_DIR=%~dp0"
set "CONDA=D:\softwares\miniconda3\Scripts\conda.exe"
set "URL=http://127.0.0.1:5000"

REM ---------- 检查项目目录还在不在 ----------
if not exist "%PROJECT_DIR%app.py" (
    echo.
    echo [错误] 找不到项目文件：%PROJECT_DIR%app.py
    echo 项目可能被移动或改名了，请检查本文件所在目录。
    echo.
    pause
    exit /b 1
)

REM ---------- 检查是否已经在运行 ----------
netstat -ano | findstr ":5000 " | findstr "LISTENING" >nul
if %errorlevel%==0 (
    echo 系统已经在运行中，直接为你打开页面...
    start "" %URL%
    timeout /t 3 /nobreak >nul
    exit /b 0
)

REM ---------- 启动前准备 ----------
cd /d "%PROJECT_DIR%"
echo 正在启动每日打卡系统，请稍候...
echo.

REM 延迟 3 秒自动打开浏览器（后台最小化小窗口，打完自动关）
start "" /min cmd /c "timeout /t 3 /nobreak >nul & start %URL%"

REM ---------- 选择 Python 环境并启动（前台运行，窗口即服务器日志） ----------
if exist "%CONDA%" goto use_conda_path
where conda >nul 2>nul
if %errorlevel%==0 goto use_conda_cmd
goto use_plain_python

:use_conda_path
"%CONDA%" run -n work --no-capture-output python app.py
goto :eof

:use_conda_cmd
conda run -n work --no-capture-output python app.py
goto :eof

:use_plain_python
echo [提示] 未找到 conda，使用当前 Python 启动；
echo 若提示缺少 Flask / openpyxl，请改用 conda 环境 work 启动。
python app.py
goto :eof

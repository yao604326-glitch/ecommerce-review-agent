@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo  电商评论分析 Agent - 一键运行
echo ============================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    py analyze.py --offline
) else (
    python analyze.py --offline
)

echo.
echo ============================================
echo 运行结束。按任意键关闭窗口。
echo ============================================
pause >nul

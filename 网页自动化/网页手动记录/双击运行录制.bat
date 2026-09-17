@echo off
rem ============================================================
rem  双击这个文件即可开始录制（.py 双击没反应的备用启动方式）
rem  结束方式：在本窗口输入 c 再按回车
rem  也可以带网址启动：把网址拖到本文件上，或在命令行里
rem      双击运行录制.bat  https://你的网址
rem ============================================================
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo [错误] 没找到 python 命令。
  echo         请先安装 Python 并勾选 "Add Python to PATH"，
  echo         然后在本窗口执行：pip install playwright
  echo.
  pause
  exit /b 1
)

python "%~dp0WebRecorder.py" %*
set RC=%ERRORLEVEL%

echo.
echo ------------------------------------------------------------
if not "%RC%"=="0" echo [提示] 程序退出码为 %RC%（非 0 表示中途出错，请把上面的信息发给 AI 排查）
echo 窗口可以关闭了（按任意键）。
pause >nul

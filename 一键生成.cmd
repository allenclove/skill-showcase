@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set "PYTHONIOENCODING=utf-8"

rem ---- 找 Python：优先系统 python，找不到就用内置运行时 ----
set "PY=python"
where python >nul 2>nul
if errorlevel 1 set "PY=D:\software\DSH Desktop\resources\office-runtime\primary-runtime\dependencies\python\python.exe"

"%PY%" -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo.
  echo [错误] 当前 Python 缺少 openpyxl，请先执行： pip install openpyxl
  echo.
  pause
  exit /b 1
)

echo 正在读取《技能展厅.xlsx》并生成网页...
echo.
"%PY%" build.py
if errorlevel 1 (
  echo.
  echo [失败] 上面列出了具体是哪个 Sheet、哪一行、哪一列有问题。
  echo        改完 Excel 再双击本文件即可。
  echo.
  pause
  exit /b 1
)

echo.
echo 生成成功，正在打开 dist\index.html ...
start "" "dist\index.html"
echo.
pause

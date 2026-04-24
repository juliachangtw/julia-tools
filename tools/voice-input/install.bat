@echo off
chcp 65001 > nul
echo.
echo ╔══════════════════════════════════════════╗
echo ║       語音輸入工具  -  安裝套件          ║
echo ╚══════════════════════════════════════════╝
echo.

:: 先確認 Python 存在
python --version > nul 2>&1
if errorlevel 1 (
    echo [錯誤] 找不到 Python，請先安裝 Python 3.9 以上版本
    echo        https://www.python.org/downloads/
    pause
    exit /b 1
)

echo 正在安裝相依套件...
echo.
pip install openai-whisper pyaudio keyboard pyperclip numpy

if errorlevel 1 (
    echo.
    echo [提示] 若 pyaudio 安裝失敗，請執行：
    echo        pip install pipwin
    echo        pipwin install pyaudio
    echo.
) else (
    echo.
    echo ✓ 安裝完成！
    echo   執行 run.bat 啟動語音輸入工具
    echo.
    echo 首次啟動時會自動下載 Whisper 語音模型（約 244 MB），
    echo 之後就不需要再下載。
)

echo.
pause

@echo off
title AutoWord - Ready to Use
color 0A

echo.
echo   ___        _        _    _               _ 
echo  / _ \      ^| ^|      ^| ^|  ^| ^|             ^| ^|
echo ^| ^|_^| ^|_   _^| ^|_ ___ ^| ^|  ^| ^| ___  _ __ __^| ^|
echo ^|  _  ^| ^| ^| ^| __/ _ \^| ^|/\^| ^|/ _ \^| '__/ _` ^|
echo ^| ^| ^| ^| ^|_^| ^| ^|^| (_) \  /\  / (_) ^| ^| ^| (_^| ^|
echo \_^| ^|_/\__,_^|\__\___/ \/  \/ \___/^|_^|  \__,_^|
echo.
echo Document Automation System
echo.

REM Quick system check
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found
    echo Please install Python 3.8+ from python.org
    pause
    exit /b 1
)

REM Install minimal dependencies
echo Installing required packages...
pip install -q pydantic requests pywin32 >nul 2>&1

echo System ready!
echo.
echo Your document: 郭宇睿-现代汉语言文学在社会发展中的作用与影响研究2稿批注.docx
echo.

REM Run the document processor
python process_document.py

pause
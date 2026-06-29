@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "DEST=%USERPROFILE%\Documents\WordPdfCopyManualInstall"

if not exist "%DEST%" mkdir "%DEST%"

copy /Y "%SCRIPT_DIR%manual-install\WordPdfCopyAddin.bas" "%DEST%\WordPdfCopyAddin.bas" >nul
copy /Y "%SCRIPT_DIR%manual-install\manual_setup_steps.txt" "%DEST%\manual_setup_steps.txt" >nul

echo Manual install files copied to:
echo %DEST%
echo.
echo Opening the folder and setup steps...
start "" "%DEST%"
notepad "%DEST%\manual_setup_steps.txt"

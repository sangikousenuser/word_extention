@echo off
setlocal

set "ADDIN_PATH=%APPDATA%\Microsoft\Word\STARTUP\WordPdfCopyAddin.dotm"

if exist "%ADDIN_PATH%" (
  del "%ADDIN_PATH%"
  echo Removed "%ADDIN_PATH%".
) else (
  echo Add-in file was not found.
)

echo.
echo Restart Microsoft Word.
pause

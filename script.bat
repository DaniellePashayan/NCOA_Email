@Echo Looking for NCOA Files

SETLOCAL
set FILE_PATH=%~dp0
set SCRIPT_PATH=%FILE_PATH%main.py
python -u "%SCRIPT_PATH%"
ENDLOCAL

@Echo Finished
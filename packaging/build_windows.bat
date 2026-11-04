@echo off
REM Run from the project folder (D:\Dev\GrepMail). Requires Python 3.12 and Inno Setup 6.
python -m venv .venv
call .venv\Scripts\activate
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --windowed --name GrepMail --icon grepmail.ico ^
  --hidden-import win32timezone --hidden-import gm_platform --hidden-import gm_security --hidden-import gm_mail --hidden-import gm_update --hidden-import PIL._tkinter_finder ^
  --collect-all ttkbootstrap --collect-all oletools ^
  grepmail.py
if errorlevel 1 goto failed

set ISCC=
where iscc >nul 2>nul && set ISCC=iscc
if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe
if not defined ISCC if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" set ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe
if not defined ISCC (
  echo Inno Setup was not found. Install Inno Setup 6, then run this script again.
  goto failed
)
"%ISCC%" packaging\grepmail.iss
if errorlevel 1 goto failed
echo Installer created in the installer folder.
exit /b 0
:failed
echo BUILD FAILED. See the messages above.
exit /b 1

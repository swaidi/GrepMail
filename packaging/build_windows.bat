@echo off
REM Run from the project folder on a Windows laptop, in a terminal opened there.
REM Requires: Python 3.12, Inno Setup 6 (iscc on PATH)
python -m venv .venv
call .venv\Scripts\activate
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --windowed --name GrepMail --icon grepmail.ico ^
  --hidden-import win32timezone --hidden-import gm_platform --hidden-import gm_security --hidden-import gm_mail --hidden-import PIL._tkinter_finder ^
  --collect-all ttkbootstrap --collect-all oletools ^
  grepmail.py
iscc packaging\grepmail.iss
echo Installer created in the installer folder.

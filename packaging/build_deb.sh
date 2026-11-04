#!/usr/bin/env bash
# Builds grepmail_2.2.0_amd64.deb on Debian/Ubuntu. Run from the project folder.
set -euo pipefail
VER=2.2.0
sudo apt-get install -y python3-venv python3-tk clamav binutils
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --windowed --name GrepMail --icon packaging/grepmail.png \
  --hidden-import gm_platform --hidden-import gm_security --hidden-import gm_mail --hidden-import gm_update --hidden-import gm_import --hidden-import PIL._tkinter_finder \
  --collect-all ttkbootstrap --collect-all oletools grepmail.py
PKG=build/grepmail_${VER}_amd64
rm -rf "$PKG" && mkdir -p "$PKG/DEBIAN" "$PKG/opt" "$PKG/usr/share/applications" "$PKG/usr/share/icons/hicolor/256x256/apps"
cp -r dist/GrepMail "$PKG/opt/grepmail"
mkdir -p "$PKG/usr/bin" && ln -s /opt/grepmail/GrepMail "$PKG/usr/bin/grepmail"
cp packaging/grepmail.desktop "$PKG/usr/share/applications/"
cp packaging/grepmail.png "$PKG/usr/share/icons/hicolor/256x256/apps/grepmail.png"
cat > "$PKG/DEBIAN/control" <<CTL
Package: grepmail
Version: ${VER}
Section: mail
Priority: optional
Architecture: amd64
Recommends: clamav, clamav-freshclam
Maintainer: GrepMail
Description: Saves tender attachments from a mailbox into reference folders
 Monitors a mailbox over IMAP and records everything in an Excel database.
CTL
dpkg-deb --build --root-owner-group "$PKG"
echo "Created $PKG.deb  (install: sudo apt install ./$PKG.deb)"

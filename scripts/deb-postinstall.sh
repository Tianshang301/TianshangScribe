#!/bin/bash
# TianshangScribe .deb post-install script
set -e

PKG_DIR="/usr/lib/tianshang-scribe"

echo "Installing TianshangScribe dependencies..."

if command -v pip3 &>/dev/null; then
    pip3 install --quiet "$PKG_DIR/"
elif command -v pip &>/dev/null; then
    pip install --quiet "$PKG_DIR/"
else
    echo "WARNING: pip not found. Install pip and run:"
    echo "  pip install $PKG_DIR/"
    exit 0
fi

echo "TianshangScribe v${VERSION:-?.?.?} installed successfully."
echo "Usage: tianshang-scribe --help"

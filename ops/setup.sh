#!/usr/bin/env bash
# Відновлює інструменти виробництва продуктів у свіжому контейнері.
set -euo pipefail
pip3 install -q openpyxl xlsxwriter reportlab python-docx pillow matplotlib pypdf pymupdf 2>&1 | grep -v "running pip as the 'root'" || true
if ! dpkg -l | grep -q libreoffice-calc; then
  apt-get update -q >/dev/null 2>&1 || true
  apt-get install -y -q libreoffice-calc >/dev/null
fi
# LibreOffice потребує записуваного профілю
export HOME="${LO_HOME:-/tmp/lohome}"; mkdir -p "$HOME"
echo "setup ok"

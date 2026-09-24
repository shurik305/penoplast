#!/usr/bin/env bash
# Відновлює інструменти виробництва продуктів у свіжому контейнері.
# Використання: source ops/setup.sh   (експортує LO_HOME для qa.py / mockups.py)
set -uo pipefail
pip3 install -q openpyxl xlsxwriter reportlab python-docx pillow matplotlib pypdf pymupdf requests fonttools brotli \
  2>&1 | grep -v "running pip as the 'root'" || true
if ! dpkg -l | grep -q libreoffice-calc; then
  apt-get update -q >/dev/null 2>&1 || true
  apt-get install -y -q libreoffice-calc >/dev/null
fi
# LibreOffice потребує записуваного профілю і має перераховувати формули при відкритті xlsx
export LO_HOME="${LO_HOME:-/tmp/lohome}"
PROFILE="$LO_HOME/.config/libreoffice/4/user"
mkdir -p "$PROFILE"
XCU="$PROFILE/registrymodifications.xcu"
ITEMS='<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>'
if [ ! -f "$XCU" ]; then
  printf '%s\n%s\n%s\n' '<?xml version="1.0" encoding="UTF-8"?>' \
    '<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' \
    "$ITEMS</oor:items>" > "$XCU"
elif ! grep -q OOXMLRecalcMode "$XCU"; then
  python3 - "$XCU" "$ITEMS" <<'EOF'
import sys
p, items = sys.argv[1], sys.argv[2]
s = open(p, encoding="utf-8").read().replace("</oor:items>", items + "\n</oor:items>")
open(p, "w", encoding="utf-8").write(s)
EOF
fi
echo "setup ok (LO_HOME=$LO_HOME)"

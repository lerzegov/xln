#!/bin/bash
# run_mac.sh — run the AppleScript name probe and capture what Excel stored on disk.
#   bash probes/mac/run_mac.sh
# Writes probes/results/mac-<host>-<timestamp>.txt. Quits Excel afterwards only if it was
# not running before. The workbook is saved inside the Office group container, which
# Excel can write without a "Grant File Access" dialog (see excel-models excel_stage.py).
set -u
HERE="$(cd "$(dirname "$0")" && pwd)"
OUT_DIR="$HERE/../results"; mkdir -p "$OUT_DIR"
STAGE="$HOME/Library/Group Containers/UBF8T346G9.Office/excel-dim-probe"; mkdir -p "$STAGE"
OUT="$OUT_DIR/mac-$(hostname -s)-$(date +%Y%m%d-%H%M%S).txt"

was_running=0; pgrep -x "Microsoft Excel" >/dev/null && was_running=1
rm -f "$STAGE/probe_mac.xlsx"

{
  echo "# excel-dim P0 probe — macOS"
  echo "# $(date -u +%FT%TZ)  macOS $(sw_vers -productVersion)  Excel $(defaults read '/Applications/Microsoft Excel.app/Contents/Info' CFBundleShortVersionString 2>/dev/null)"
  echo "# locale: $(defaults read -g AppleLocale 2>/dev/null)"
  echo
  echo "## AppleScript tests"
  osascript "$HERE/probe_names.applescript" "$STAGE" 2>&1
  echo
  echo "## Stored form: <definedNames> in xl/workbook.xml of the saved copy"
  if [ -f "$STAGE/probe_mac.xlsx" ]; then
    unzip -p "$STAGE/probe_mac.xlsx" xl/workbook.xml \
      | python3 -c 'import sys,re,html
x=sys.stdin.read()
for m in re.finditer(r"<definedName ([^>]*)>(.*?)</definedName>", x, re.S):
    print(m.group(1), "=>", html.unescape(m.group(2))[:160].replace("\n","\\n"))'
    cp "$STAGE/probe_mac.xlsx" "$OUT_DIR/"
  else
    echo "(no saved copy)"
  fi
} | tee "$OUT"

[ $was_running -eq 0 ] && osascript -e 'tell application "Microsoft Excel" to quit' >/dev/null 2>&1
echo; echo "report: $OUT"

#!/bin/zsh

set -u
setopt pipefail

SCRIPT_DIR=${0:A:h}
PROJECT_ROOT=${SCRIPT_DIR:h}
DEFAULT_WORKBOOK="$PROJECT_ROOT/outputs/gtm_excel_v0_3/GTM_Excel_Interface_v0.3.xlsx"
WORKBOOK=${1:-$DEFAULT_WORKBOOK}
OUTPUT_ROOT=${2:-$PROJECT_ROOT/outputs/gtm_excel_runs}
ENGINE="$PROJECT_ROOT/.venv/bin/gtm-engine"

if [[ ! -x "$ENGINE" ]]; then
  print "GTM build failed: $ENGINE is missing or is not executable."
  print "Install the project environment before running this launcher."
  exit 2
fi

if [[ ! -f "$WORKBOOK" ]]; then
  print "GTM build failed: workbook not found."
  print "$WORKBOOK"
  print "Drag a GTM .xlsx workbook onto this launcher, or pass its path as argument 1."
  exit 2
fi

print "Building GTM workbook:"
print "$WORKBOOK"
print ""

"$ENGINE" excel-build --workbook "$WORKBOOK" --output "$OUTPUT_ROOT"
STATUS=$?

print ""
if [[ $STATUS -eq 0 ]]; then
  print "Build published successfully. Open:"
  print "$OUTPUT_ROOT/GTM_LATEST.xlsx"
else
  print "Build failed. Read the JSON message above, then open the retained"
  print "GTM_Failed.xlsx under $OUTPUT_ROOT/failed/."
fi

if [[ -t 0 ]]; then
  print ""
  read -k 1 "?Press any key to close."
  print ""
fi

exit $STATUS

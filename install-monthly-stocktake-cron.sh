#!/usr/bin/env bash
set -euo pipefail

APP_DIR="${APP_DIR:-/home/zaialsadmin/StocktakeGUI}"
RUN_SCRIPT="${RUN_SCRIPT:-$APP_DIR/run-stocktake-report.sh}"
LOG_FILE="${LOG_FILE:-$APP_DIR/reports/monthly-stocktake-cron.log}"
CRON_TIME="${CRON_TIME:-0 * * * *}"
LOCAL_TZ="${LOCAL_TZ:-Australia/Sydney}"
LOCAL_DAY_HOUR_MINUTE="${LOCAL_DAY_HOUR_MINUTE:-040800}"
CRON_MARKER="# stocktake-monthly-report"

if [[ ! -f "$RUN_SCRIPT" ]]; then
  echo "Run script not found: $RUN_SCRIPT" >&2
  exit 1
fi

mkdir -p "$(dirname "$LOG_FILE")"
chmod +x "$RUN_SCRIPT"

CRON_LINE="$CRON_TIME [ \"\$(TZ=$LOCAL_TZ date +\\%d\\%H\\%M)\" = \"$LOCAL_DAY_HOUR_MINUTE\" ] && $RUN_SCRIPT >> $LOG_FILE 2>&1 $CRON_MARKER"
CURRENT_CRON="$(crontab -l 2>/dev/null || true)"
FILTERED_CRON="$(printf '%s\n' "$CURRENT_CRON" | grep -vF "$CRON_MARKER" || true)"

{
  printf '%s\n' "$FILTERED_CRON" | sed '/^[[:space:]]*$/d'
  printf '%s\n' "$CRON_LINE"
} | crontab -

echo "Installed monthly stocktake report schedule:"
echo "$CRON_LINE"

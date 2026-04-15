#!/usr/bin/env bash

set -euo pipefail

BASE_URL="${BASE_URL:-http://10.74.43.136:8000}"
EXCEL_FILE="${EXCEL_FILE:-/root/data.xlsx}"
TEMPLATE_ID="${TEMPLATE_ID:-mss_classic_ops}"
USE_RAG="${USE_RAG:-false}"
FOCUS_OPTION="${FOCUS_OPTION:-vulnerability}"
POLL_INTERVAL_SECONDS="${POLL_INTERVAL_SECONDS:-5}"
POLL_TIMEOUT_SECONDS="${POLL_TIMEOUT_SECONDS:-1800}"
OUTPUT_FILE="${OUTPUT_FILE:-report.pptx}"
COOKIE_JAR="${COOKIE_JAR:-report_cookies.txt}"

WORK_DIR="$(mktemp -d)"
trap 'rm -rf "$WORK_DIR"' EXIT

UPLOAD_JSON="$WORK_DIR/upload.json"
CREATE_JSON="$WORK_DIR/create.json"
STATUS_JSON="$WORK_DIR/status.json"

require_command() {
  if ! command -v "$1" >/dev/null 2>&1; then
    echo "Missing required command: $1" >&2
    exit 1
  fi
}

json_get() {
  local file="$1"
  local expr="$2"
  python3 - "$file" "$expr" <<'PY'
import json
import sys

path = sys.argv[1]
expr = sys.argv[2]

with open(path, "r", encoding="utf-8") as f:
    data = json.load(f)

value = data
for part in expr.split("."):
    if not part:
        continue
    if isinstance(value, dict) and part in value:
        value = value[part]
    else:
        sys.exit(2)

if isinstance(value, bool):
    print("true" if value else "false")
elif value is None:
    print("")
else:
    print(value)
PY
}

fail_with_response() {
  local prefix="$1"
  local file="$2"
  echo "$prefix" >&2
  cat "$file" >&2
  exit 1
}

require_command curl
require_command python3

if [[ ! -f "$EXCEL_FILE" ]]; then
  echo "Excel file not found: $EXCEL_FILE" >&2
  exit 1
fi

echo "Uploading Excel: $EXCEL_FILE"
curl -sS -f -X POST "$BASE_URL/api/v1/inputs/excel" \
  -F "file=@${EXCEL_FILE};type=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" \
  -F "template_id=${TEMPLATE_ID}" \
  > "$UPLOAD_JSON" || fail_with_response "Upload request failed." "$UPLOAD_JSON"

SESSION_ID="$(json_get "$UPLOAD_JSON" "data.session_id" || true)"
if [[ -z "$SESSION_ID" ]]; then
  fail_with_response "Upload response missing data.session_id." "$UPLOAD_JSON"
fi

echo "Upload succeeded. session_id=$SESSION_ID"

CREATE_BODY="$(python3 - "$SESSION_ID" "$TEMPLATE_ID" "$USE_RAG" "$FOCUS_OPTION" <<'PY'
import json
import sys

session_id, template_id, use_rag_raw, focus_option = sys.argv[1:5]
use_rag = use_rag_raw.lower() == "true"

body = {
    "input_id": session_id,
    "template_id": template_id,
    "session_id": session_id,
    "use_rag": use_rag,
    "focus_options": [focus_option],
}

print(json.dumps(body, ensure_ascii=True))
PY
)"

echo "Creating report job"
curl -sS -f -c "$COOKIE_JAR" -X POST "$BASE_URL/api/v1/reports" \
  -H "Content-Type: application/json" \
  -d "$CREATE_BODY" \
  > "$CREATE_JSON" || fail_with_response "Create report request failed." "$CREATE_JSON"

JOB_ID="$(json_get "$CREATE_JSON" "data.job_id" || true)"
JOB_STATUS="$(json_get "$CREATE_JSON" "data.status" || true)"

if [[ -z "$JOB_ID" ]]; then
  fail_with_response "Create report response missing data.job_id." "$CREATE_JSON"
fi

echo "Job created. job_id=$JOB_ID status=$JOB_STATUS"

if [[ "$JOB_STATUS" != "completed" ]]; then
  START_TS="$(date +%s)"
  while true; do
    NOW_TS="$(date +%s)"
    ELAPSED="$((NOW_TS - START_TS))"
    if (( ELAPSED > POLL_TIMEOUT_SECONDS )); then
      echo "Polling timed out after ${POLL_TIMEOUT_SECONDS}s" >&2
      exit 1
    fi

    sleep "$POLL_INTERVAL_SECONDS"

    curl -sS -f -b "$COOKIE_JAR" \
      "$BASE_URL/api/v1/jobs/${JOB_ID}/status" \
      > "$STATUS_JSON" || fail_with_response "Status request failed." "$STATUS_JSON"

    JOB_STATUS="$(json_get "$STATUS_JSON" "data.status" || true)"
    JOB_PROGRESS="$(json_get "$STATUS_JSON" "data.progress" || true)"
    JOB_MESSAGE="$(json_get "$STATUS_JSON" "data.message" || true)"

    echo "Status: ${JOB_STATUS:-unknown} progress=${JOB_PROGRESS:-?} message=${JOB_MESSAGE:-}"

    if [[ "$JOB_STATUS" == "completed" ]]; then
      break
    fi

    if [[ "$JOB_STATUS" == "failed" || "$JOB_STATUS" == "cancelled" ]]; then
      fail_with_response "Job did not complete successfully." "$STATUS_JSON"
    fi
  done
fi

echo "Downloading report to $OUTPUT_FILE"
curl -sS -f -L -b "$COOKIE_JAR" \
  "$BASE_URL/api/v1/reports/${JOB_ID}/download" \
  -o "$OUTPUT_FILE"

echo "Done. report_file=$OUTPUT_FILE job_id=$JOB_ID"

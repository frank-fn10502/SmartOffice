#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
API_SCRIPT="$SCRIPT_DIR/../outlook-api.sh"
INBOX_SCRIPT="$SCRIPT_DIR/inbox.sh"
TAKE="100"
LOOKBACK_HOURS="168"
MAX_COUNT="30"
BASE_URL_ARGS=()

usage() {
  cat <<'USAGE'
Read recent mails from the primary mailbox Inbox.

Usage:
  recent-mails.sh [--base-url URL] [--lookback-hours N] [--max-count N] [--take N]

Output:
  JSON object: { "folderPath": "...", "request": {...}, "mails": [...] }
USAGE
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --base-url)
      BASE_URL_ARGS=(--base-url "$2")
      shift 2
      ;;
    --lookback-hours)
      LOOKBACK_HOURS="$2"
      shift 2
      ;;
    --max-count)
      MAX_COUNT="$2"
      shift 2
      ;;
    --take)
      TAKE="$2"
      shift 2
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "Unknown option: $1" >&2
      usage >&2
      exit 2
      ;;
  esac
done

inbox_json="$("$INBOX_SCRIPT" "${BASE_URL_ARGS[@]+"${BASE_URL_ARGS[@]}"}" --take "$TAKE")"
folder_path="$(jq -r '.folder.folderPath' <<< "$inbox_json")"
body="$(jq -nc --arg folderPath "$folder_path" --argjson lookbackHours "$LOOKBACK_HOURS" --argjson maxCount "$MAX_COUNT" \
  '{folderPath:$folderPath,lookbackHours:$lookbackHours,maxCount:$maxCount}')"
result="$("$API_SCRIPT" "${BASE_URL_ARGS[@]+"${BASE_URL_ARGS[@]}"}" request-fetch /api/outlook/request-mails "$body" --take "$TAKE")"

jq -nc --arg folderPath "$folder_path" --argjson request "$body" --argjson result "$result" \
  '{folderPath:$folderPath,request:$request,mails:[ $result.fetchResult.pages[].data.mails[]? ]}'

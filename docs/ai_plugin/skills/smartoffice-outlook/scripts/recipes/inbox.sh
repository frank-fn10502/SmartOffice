#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
API_SCRIPT="$SCRIPT_DIR/../outlook-api.sh"
TAKE="100"
BASE_URL_ARGS=()

usage() {
  cat <<'USAGE'
Locate the primary mailbox Inbox through the SmartOffice Outlook API.

Usage:
  inbox.sh [--base-url URL] [--take N]

Output:
  JSON object: { "folder": <FolderTreeNodeDto> }
USAGE
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --base-url)
      BASE_URL_ARGS=(--base-url "$2")
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

body='{"name":"","folderPath":"","folderType":"Inbox","storeId":"","includeHidden":false,"maxResults":20}'
result="$("$API_SCRIPT" "${BASE_URL_ARGS[@]+"${BASE_URL_ARGS[@]}"}" request-fetch /api/outlook/request-find-folder "$body" --take "$TAKE")"
match_count="$(jq -r '[.fetchResult.pages[].data.matchCount?] | add // 0' <<< "$result")"
is_ambiguous="$(jq -r '[.fetchResult.pages[].data.isAmbiguous?] | any' <<< "$result")"
folder="$(jq -c '[.fetchResult.pages[].data.folders[]?] | .[0] // null' <<< "$result")"

if [[ "$match_count" != "1" || "$is_ambiguous" == "true" || "$folder" == "null" ]]; then
  echo "$result"
  echo "Inbox could not be uniquely located." >&2
  exit 1
fi

jq -nc --argjson folder "$folder" '{folder:$folder}'

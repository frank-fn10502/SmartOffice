#!/usr/bin/env bash
set -euo pipefail

BASE_URL="${SMARTOFFICE_OUTLOOK_BASE_URL:-http://localhost:2805}"
TAKE="100"

usage() {
  cat <<'USAGE'
SmartOffice Outlook HTTP API helper.

Outputs JSON to stdout. Diagnostic text goes to stderr.

Usage:
  outlook-api.sh [--base-url URL] status
  outlook-api.sh [--base-url URL] post <path> <json-or-@file>
  outlook-api.sh [--base-url URL] fetch <fetch-result-path> <request-id> [--take N]
  outlook-api.sh [--base-url URL] request-fetch <request-path> <json-or-@file> [--take N]
  outlook-api.sh [--base-url URL] inbox
  outlook-api.sh [--base-url URL] recent-mails [--lookback-hours N] [--max-count N] [--take N]

Examples:
  ./scripts/outlook-api.sh status
  ./scripts/outlook-api.sh inbox
  ./scripts/outlook-api.sh recent-mails --lookback-hours 168 --max-count 30
  ./scripts/outlook-api.sh request-fetch /api/outlook/request-calendar '{"daysForward":31,"startDate":null,"endDate":null}'
  ./scripts/outlook-api.sh post /api/outlook/request-mail-search @request.json

Rules implemented by this helper:
  - request-* responses are not treated as success until paired fetch-result completes.
  - fetch-result pagination continues while next.hasMore=true, even when state=completed.
  - failed, unavailable, and timeout states stop the helper with a non-zero exit code.
USAGE
}

need_tool() {
  if ! command -v "$1" >/dev/null 2>&1; then
    echo "Missing required tool: $1" >&2
    exit 127
  fi
}

json_arg() {
  local value="$1"
  if [[ "$value" == @* ]]; then
    local file="${value#@}"
    if [[ ! -f "$file" ]]; then
      echo "JSON file not found: $file" >&2
      exit 2
    fi
    cat "$file"
  else
    printf '%s' "$value"
  fi
}

http_get() {
  local path="$1"
  curl -sS "${BASE_URL}${path}"
}

http_post() {
  local path="$1"
  local body="$2"
  local response status content
  response="$(curl -sS -w '\n__SMARTOFFICE_HTTP_STATUS__:%{http_code}' \
    -H 'Content-Type: application/json' \
    -X POST \
    --data "$body" \
    "${BASE_URL}${path}")"
  status="${response##*$'\n'__SMARTOFFICE_HTTP_STATUS__:}"
  content="${response%$'\n'__SMARTOFFICE_HTTP_STATUS__:*}"
  printf '%s' "$content"
  if [[ "$status" -lt 200 || "$status" -ge 300 ]]; then
    return 1
  fi
}

fetch_all() {
  local endpoint="$1"
  local request_id="$2"
  local take="$3"
  local cursor=""
  local pages='[]'
  local state has_more next_cursor page body

  for _ in $(seq 1 200); do
    body="$(jq -nc --arg requestId "$request_id" --arg cursor "$cursor" --argjson take "$take" \
      '{requestId:$requestId,cursor:$cursor,take:$take}')"
    page="$(http_post "$endpoint" "$body")"
    pages="$(jq -c --argjson page "$page" '. + [$page]' <<< "$pages")"
    state="$(jq -r '.state // ""' <<< "$page")"
    has_more="$(jq -r '.next.hasMore // false' <<< "$page")"
    next_cursor="$(jq -r '.next.cursor // ""' <<< "$page")"

    case "$state" in
      failed|unavailable|timeout)
        jq -nc --arg endpoint "$endpoint" --arg requestId "$request_id" --argjson pages "$pages" \
          '{endpoint:$endpoint,requestId:$requestId,state:"failed",pages:$pages}'
        return 1
        ;;
    esac

    if [[ "$has_more" == "true" ]]; then
      cursor="$next_cursor"
      continue
    fi

    if [[ "$state" == "completed" ]]; then
      jq -nc --arg endpoint "$endpoint" --arg requestId "$request_id" --argjson pages "$pages" \
        '{endpoint:$endpoint,requestId:$requestId,state:"completed",pages:$pages}'
      return 0
    fi

    sleep 0.2
  done

  jq -nc --arg endpoint "$endpoint" --arg requestId "$request_id" --argjson pages "$pages" \
    '{endpoint:$endpoint,requestId:$requestId,state:"timeout",message:"fetch-result loop exceeded 200 attempts",pages:$pages}'
  return 1
}

request_fetch() {
  local request_path="$1"
  local request_body="$2"
  local take="$3"
  local response request_id endpoint result
  response="$(http_post "$request_path" "$request_body")"
  request_id="$(jq -r '.requestId // ""' <<< "$response")"
  endpoint="$(jq -r '.data.fetchResultEndpoint // ""' <<< "$response")"
  if [[ -z "$request_id" || -z "$endpoint" ]]; then
    echo "$response"
    echo "request response did not include requestId or data.fetchResultEndpoint" >&2
    return 1
  fi
  result="$(fetch_all "$endpoint" "$request_id" "$take")"
  jq -nc --argjson requestResponse "$response" --argjson fetchResult "$result" \
    '{requestResponse:$requestResponse,fetchResult:$fetchResult}'
}

inbox() {
  local body result folder
  body='{"name":"","folderPath":"","folderType":"Inbox","storeId":"","includeHidden":false,"maxResults":20}'
  result="$(request_fetch "/api/outlook/request-find-folder" "$body" "$TAKE")"
  folder="$(jq -c '[.fetchResult.pages[].data.folders[]?] | .[0] // null' <<< "$result")"
  if [[ "$folder" == "null" ]]; then
    echo "$result"
    echo "Inbox could not be uniquely located." >&2
    return 1
  fi
  jq -nc --argjson folder "$folder" '{folder:$folder}'
}

recent_mails() {
  local lookback_hours="$1"
  local max_count="$2"
  local inbox_json folder_path body result
  inbox_json="$(inbox)"
  folder_path="$(jq -r '.folder.folderPath' <<< "$inbox_json")"
  body="$(jq -nc --arg folderPath "$folder_path" --argjson lookbackHours "$lookback_hours" --argjson maxCount "$max_count" \
    '{folderPath:$folderPath,lookbackHours:$lookbackHours,maxCount:$maxCount}')"
  result="$(request_fetch "/api/outlook/request-mails" "$body" "$TAKE")"
  jq -nc --arg folderPath "$folder_path" --argjson request "$body" --argjson result "$result" \
    '{folderPath:$folderPath,request:$request,mails:[ $result.fetchResult.pages[].data.mails[]? ]}'
}

need_tool curl
need_tool jq

if [[ $# -eq 0 ]]; then
  usage
  exit 2
fi

while [[ $# -gt 0 ]]; do
  case "$1" in
    --base-url)
      BASE_URL="${2%/}"
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
      break
      ;;
  esac
done

command_name="${1:-}"
shift || true

case "$command_name" in
  status)
    http_get "/api/outlook/admin/status"
    ;;
  post)
    if [[ $# -lt 2 ]]; then usage >&2; exit 2; fi
    http_post "$1" "$(json_arg "$2")"
    ;;
  fetch)
    if [[ $# -lt 2 ]]; then usage >&2; exit 2; fi
    endpoint="$1"
    request_id="$2"
    shift 2
    while [[ $# -gt 0 ]]; do
      case "$1" in
        --take) TAKE="$2"; shift 2 ;;
        *) echo "Unknown option: $1" >&2; exit 2 ;;
      esac
    done
    fetch_all "$endpoint" "$request_id" "$TAKE"
    ;;
  request-fetch)
    if [[ $# -lt 2 ]]; then usage >&2; exit 2; fi
    request_path="$1"
    request_body="$(json_arg "$2")"
    shift 2
    while [[ $# -gt 0 ]]; do
      case "$1" in
        --take) TAKE="$2"; shift 2 ;;
        *) echo "Unknown option: $1" >&2; exit 2 ;;
      esac
    done
    request_fetch "$request_path" "$request_body" "$TAKE"
    ;;
  inbox)
    inbox
    ;;
  recent-mails)
    lookback_hours="168"
    max_count="30"
    while [[ $# -gt 0 ]]; do
      case "$1" in
        --lookback-hours) lookback_hours="$2"; shift 2 ;;
        --max-count) max_count="$2"; shift 2 ;;
        --take) TAKE="$2"; shift 2 ;;
        *) echo "Unknown option: $1" >&2; exit 2 ;;
      esac
    done
    recent_mails "$lookback_hours" "$max_count"
    ;;
  *)
    usage >&2
    exit 2
    ;;
esac

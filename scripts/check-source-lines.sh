#!/usr/bin/env bash
set -euo pipefail

MAX_LINES="${SMARTOFFICE_MAX_SOURCE_LINES:-800}"
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OVERSIZED=()

while IFS= read -r -d '' file; do
  lines="$(wc -l < "${file}")"
  lines="${lines//[[:space:]]/}"
  if [ "${lines}" -gt "${MAX_LINES}" ]; then
    relative="${file#"${ROOT_DIR}/"}"
    OVERSIZED+=("${lines} ${relative}")
  fi
done < <(
  find "${ROOT_DIR}" \
    \( -path "${ROOT_DIR}/.git" \
      -o -path "${ROOT_DIR}/.vs" \
      -o -path "*/bin" \
      -o -path "*/obj" \
      -o -path "*/node_modules" \
      -o -path "*/dist" \) -prune \
    -o -type f \
    \( -name '*.cs' \
      -o -name '*.ts' \
      -o -name '*.vue' \
      -o -name '*.js' \
      -o -name '*.mjs' \
      -o -name '*.css' \
      -o -name '*.sh' \) \
    -print0
)

if [ "${#OVERSIZED[@]}" -gt 0 ]; then
  printf 'Source file line-count gate failed. Max allowed lines: %s.\n' "${MAX_LINES}" >&2
  printf '%s\n' "${OVERSIZED[@]}" | sort -nr >&2
  exit 1
fi

printf 'Source file line-count gate passed. Max allowed lines: %s.\n' "${MAX_LINES}"

#!/usr/bin/env bash
set -euo pipefail

CONFIGURATION="${CONFIGURATION:-Debug}"
PREINSTALLED_WEBUI_DIR="/opt/smartoffice-hub/webui"
WEBUI_DIR="webui"
LOCK_MARKER="node_modules/.smartoffice-package-lock"

sync_webui_dependencies() {
  if [ ! -f "${WEBUI_DIR}/package.json" ]; then
    return
  fi

  pushd "${WEBUI_DIR}" >/dev/null

  if [ -d "${PREINSTALLED_WEBUI_DIR}/node_modules" ] \
    && [ -f "${PREINSTALLED_WEBUI_DIR}/package-lock.json" ] \
    && cmp -s package-lock.json "${PREINSTALLED_WEBUI_DIR}/package-lock.json"; then
    if [ ! -f "${LOCK_MARKER}" ] || ! cmp -s package-lock.json "${LOCK_MARKER}"; then
      find node_modules -mindepth 1 -maxdepth 1 -exec rm -rf {} +
      cp -a "${PREINSTALLED_WEBUI_DIR}/node_modules/." node_modules/
      cp package-lock.json "${LOCK_MARKER}"
    fi
  elif [ -f package-lock.json ]; then
    npm ci --no-audit --no-fund
  else
    npm install --no-audit --no-fund
  fi

  npm run build
  popd >/dev/null
}

sync_webui_dependencies
dotnet build SmartOffice.Hub.sln --configuration "${CONFIGURATION}"

#!/usr/bin/env bash
set -euo pipefail

IMAGE="${SMARTOFFICE_BUILD_IMAGE:-smartoffice-hub-devcontainer-node22:local}"
CONFIGURATION="${CONFIGURATION:-Debug}"
NODE_MODULES_VOLUME="${SMARTOFFICE_WEBUI_NODE_MODULES_VOLUME:-smartoffice-hub-webui-node-modules}"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

if ! docker image inspect "${IMAGE}" >/dev/null 2>&1; then
  docker build \
    -f "${REPO_ROOT}/.devcontainer/Dockerfile" \
    -t "${IMAGE}" \
    "${REPO_ROOT}"
fi

docker run --rm \
  -e CONFIGURATION="${CONFIGURATION}" \
  -v "${REPO_ROOT}:/workspace" \
  -v "${NODE_MODULES_VOLUME}:/workspace/webui/node_modules" \
  -w /workspace \
  "${IMAGE}" \
  bash -lc 'if [ -f webui/package.json ]; then cd webui && if [ -f package-lock.json ]; then npm ci --no-audit --no-fund; else npm install --no-audit --no-fund; fi && npm run build && cd ..; fi; dotnet build SmartOffice.Hub.sln --configuration "$CONFIGURATION"'

rmdir "${REPO_ROOT}/webui/node_modules" 2>/dev/null || true

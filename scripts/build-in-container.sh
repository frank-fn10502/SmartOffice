#!/usr/bin/env bash
set -euo pipefail

IMAGE="${SMARTOFFICE_BUILD_IMAGE:-smartoffice-hub-devcontainer-node22:local}"
CONFIGURATION="${CONFIGURATION:-Debug}"

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
  -w /workspace \
  "${IMAGE}" \
  bash -lc 'set -euo pipefail; if [ -f webui/package.json ]; then pushd webui >/dev/null; if [ ! -d node_modules ]; then if [ -f package-lock.json ]; then npm ci --no-audit --no-fund; else npm install --no-audit --no-fund; fi; fi; npm run build; popd >/dev/null; fi; dotnet build SmartOffice.Hub.sln --configuration "$CONFIGURATION"'

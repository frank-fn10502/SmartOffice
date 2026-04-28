#!/usr/bin/env bash
set -euo pipefail

IMAGE="${SMARTOFFICE_BUILD_IMAGE:-smartoffice-hub-dev:local}"
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
  -v "${REPO_ROOT}:/workspace" \
  -w /workspace \
  "${IMAGE}" \
  dotnet build SmartOffice.Hub.sln --configuration "${CONFIGURATION}"

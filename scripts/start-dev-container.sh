#!/usr/bin/env bash
set -euo pipefail

IMAGE="${SMARTOFFICE_BUILD_IMAGE:-smartoffice-hub-devcontainer-node22:local}"
CONTAINER_NAME="${SMARTOFFICE_DEV_CONTAINER:-smartoffice-hub-dev}"
HOST_PORT="${SMARTOFFICE_HOST_PORT:-2805}"
ASPNETCORE_ENVIRONMENT="${SMARTOFFICE_ASPNETCORE_ENVIRONMENT:-Mock}"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

if ! docker image inspect "${IMAGE}" >/dev/null 2>&1; then
  docker build \
    -f "${REPO_ROOT}/.devcontainer/Dockerfile" \
    -t "${IMAGE}" \
    "${REPO_ROOT}"
fi

if docker container inspect "${CONTAINER_NAME}" >/dev/null 2>&1; then
  docker rm -f "${CONTAINER_NAME}" >/dev/null
fi

if [ ! -d "${REPO_ROOT}/webui/node_modules" ] \
  || [ ! -f "${REPO_ROOT}/wwwroot/index.html" ] \
  || find "${REPO_ROOT}/webui/src" "${REPO_ROOT}/webui/index.html" "${REPO_ROOT}/webui/package.json" "${REPO_ROOT}/webui/vite.config.ts" -newer "${REPO_ROOT}/wwwroot/index.html" | grep -q .; then
  "${SCRIPT_DIR}/build-in-container.sh"
fi

docker run -d \
  --name "${CONTAINER_NAME}" \
  -e ASPNETCORE_ENVIRONMENT="${ASPNETCORE_ENVIRONMENT}" \
  -e ASPNETCORE_URLS=http://0.0.0.0:2805 \
  -p "${HOST_PORT}:2805" \
  -v "${REPO_ROOT}:/workspace" \
  -w /workspace \
  "${IMAGE}" \
  bash -lc 'set -euo pipefail; dotnet run --project SmartOffice.Hub.csproj --urls http://0.0.0.0:2805'

printf 'SmartOffice.Hub is starting in Docker container "%s" with ASPNETCORE_ENVIRONMENT=%s.\n' "${CONTAINER_NAME}" "${ASPNETCORE_ENVIRONMENT}"
printf 'Open http://localhost:%s/\n' "${HOST_PORT}"
printf 'Logs: docker logs -f %s\n' "${CONTAINER_NAME}"

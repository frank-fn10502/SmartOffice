#!/usr/bin/env bash
set -euo pipefail

# 從 Docker Hub 下載 image
# 版號從 .env.docker 檔案讀取，並同步拉取 latest
# 使用 multi-arch manifest，自動選擇可執行的架構
#
# 用法:
#   ./scripts/pull-from-dockerhub.sh

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
source "${REPO_ROOT}/scripts/dockerhub/common.sh"

if ! ENV_FILE="$(resolve_env_file "${REPO_ROOT}")"; then
  echo "錯誤：找不到 ${REPO_ROOT}/scripts/dockerhub/.env.docker"
  exit 1
fi

DOCKER_IMAGE_VERSION="$(read_docker_image_version "${ENV_FILE}")"

# 使用 manifest tag (無架構後綴)，Docker 會自動根據系統選擇
TAG="${DOCKER_IMAGE_VERSION}"
SOURCE_IMAGE="${DOCKER_HUB_REPO}:${TAG}"
LATEST_IMAGE="${DOCKER_HUB_REPO}:latest"
LOCAL_IMAGE_NAME="smartoffice-hub-devcontainer-node22:local"

echo "從 Docker Hub 下載 image (版號: ${DOCKER_IMAGE_VERSION})"
echo "=============================================="
echo ""
echo "設定檔: ${ENV_FILE}"
echo "下載版號 tag: ${SOURCE_IMAGE}"
docker pull "${SOURCE_IMAGE}"

echo "下載 latest tag: ${LATEST_IMAGE}"
docker pull "${LATEST_IMAGE}"

echo ""
echo "為本地鏡像重新標記: ${SOURCE_IMAGE} -> ${LOCAL_IMAGE_NAME}"
docker tag "${SOURCE_IMAGE}" "${LOCAL_IMAGE_NAME}"

echo ""
echo "✓ 完成！鏡像已下載"
echo ""
echo "現在您可以使用以下命令進行構建："
echo "  SMARTOFFICE_BUILD_IMAGE=${LOCAL_IMAGE_NAME} ./scripts/build-in-container.sh"
echo ""
echo "或直接運行："
echo "  ./scripts/build-in-container.sh"

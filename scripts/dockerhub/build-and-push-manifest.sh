#!/usr/bin/env bash
set -euo pipefail

# 建立並推送多架構 manifest
# 將 x86-linux 和 arm64 的 image 聯繫在同一個 manifest 下
# 注意：單架構 image 目前由 docker buildx 推送為 OCI image index，
# 因此這裡需使用 docker buildx imagetools create，而不是 docker manifest create。
# 版號從 .env.docker 檔案讀取，並同步推送 latest
#
# 用法:
#   ./scripts/dockerhub/build-and-push-manifest.sh

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/../.." && pwd)"
source "${SCRIPT_DIR}/common.sh"

if ! ENV_FILE="$(resolve_env_file "${REPO_ROOT}")"; then
  echo "錯誤：找不到 ${REPO_ROOT}/scripts/dockerhub/.env.docker"
  exit 1
fi

DOCKER_IMAGE_VERSION="$(read_docker_image_version "${ENV_FILE}")"

X86_IMAGE="${DOCKER_HUB_REPO}:${DOCKER_IMAGE_VERSION}-x86-linux"
ARM64_IMAGE="${DOCKER_HUB_REPO}:${DOCKER_IMAGE_VERSION}-arm64"
MANIFEST_TAGS=("${DOCKER_IMAGE_VERSION}" "latest")

echo "正在建立 multi-arch manifest..."
echo "設定檔: ${ENV_FILE}"
echo "版號: ${DOCKER_IMAGE_VERSION}"
echo "X86-Linux: ${X86_IMAGE}"
echo "ARM64: ${ARM64_IMAGE}"
echo ""

for MANIFEST_TAG in "${MANIFEST_TAGS[@]}"; do
  MANIFEST_IMAGE="${DOCKER_HUB_REPO}:${MANIFEST_TAG}"

  echo "建立並推送 manifest: ${MANIFEST_IMAGE}"
  docker buildx imagetools create \
    --tag "${MANIFEST_IMAGE}" \
    "${X86_IMAGE}" \
    "${ARM64_IMAGE}"
  echo ""
done

echo "成功！version 與 latest manifest 已推送"
echo "工作機可以使用 pull 命令下載："
echo "  ./scripts/pull-from-dockerhub.sh"
echo ""
echo "或直接使用通用 tag："
echo "  docker pull ${DOCKER_HUB_REPO}:${DOCKER_IMAGE_VERSION}"
echo "  docker pull ${DOCKER_HUB_REPO}:latest"

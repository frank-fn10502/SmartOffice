#!/usr/bin/env bash
set -euo pipefail

# 為 arm64 平台建立並推送 image 到 Docker Hub
# 在 macOS 上使用 docker buildx 來支援 arm64 (linux/arm64) 架構的構建
# 版號從 .env.docker 檔案讀取，tag 格式為: {version}-arm64
#
# 用法:
#   ../build-and-push-arm64.sh

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/../.." && pwd)"
source "${SCRIPT_DIR}/common.sh"

if ! ENV_FILE="$(resolve_env_file "${REPO_ROOT}")"; then
  echo "錯誤：找不到 ${REPO_ROOT}/scripts/dockerhub/.env.docker"
  exit 1
fi

DOCKER_IMAGE_VERSION="$(read_docker_image_version "${ENV_FILE}")"

TAG="${DOCKER_IMAGE_VERSION}-arm64"
TARGET_IMAGE="${DOCKER_HUB_REPO}:${TAG}"

ensure_buildx_builder "smartoffice-builder"

echo "正在為 linux/arm64 平台建立並推送 image..."
echo "設定檔: ${ENV_FILE}"
echo "來源: ${REPO_ROOT}/.devcontainer/Dockerfile"
echo "目標: ${TARGET_IMAGE}"

docker buildx build \
  --platform linux/arm64 \
  -f "${REPO_ROOT}/.devcontainer/Dockerfile" \
  -t "${TARGET_IMAGE}" \
  --push \
  "${REPO_ROOT}"

echo "成功！arm64 image 已推送到 ${TARGET_IMAGE}"
echo ""
echo "下一步："
echo "  1. 如果尚未推送 x86-linux 版本，執行："
echo "     ./scripts/dockerhub/build-and-push-x86-linux.sh"
echo "  2. 建立並推送多架構 manifest（含 latest）："
echo "     ./scripts/dockerhub/build-and-push-manifest.sh"

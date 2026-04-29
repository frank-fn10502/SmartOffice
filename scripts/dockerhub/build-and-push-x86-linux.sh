#!/usr/bin/env bash
set -euo pipefail

# 為 x86-linux 平台建立並推送 image 到 Docker Hub
# 在 macOS 上使用 docker buildx 來支援 x86-linux (amd64) 架構的構建
# 版號從 .env.docker 檔案讀取，tag 格式為: {version}-x86-linux
#
# 用法:
#   ../build-and-push-x86-linux.sh

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/../.." && pwd)"
source "${SCRIPT_DIR}/common.sh"

if ! ENV_FILE="$(resolve_env_file "${REPO_ROOT}")"; then
  echo "錯誤：找不到 ${REPO_ROOT}/scripts/dockerhub/.env.docker"
  exit 1
fi

DOCKER_IMAGE_VERSION="$(read_docker_image_version "${ENV_FILE}")"

TAG="${DOCKER_IMAGE_VERSION}-x86-linux"
TARGET_IMAGE="${DOCKER_HUB_REPO}:${TAG}"

ensure_buildx_builder "smartoffice-builder"

echo "正在為 linux/amd64 平台建立並推送 image..."
echo "設定檔: ${ENV_FILE}"
echo "來源: ${REPO_ROOT}/.devcontainer/Dockerfile"
echo "目標: ${TARGET_IMAGE}"

docker buildx build \
  --platform linux/amd64 \
  -f "${REPO_ROOT}/.devcontainer/Dockerfile" \
  -t "${TARGET_IMAGE}" \
  --push \
  "${REPO_ROOT}"

echo "成功！x86-linux image 已推送到 ${TARGET_IMAGE}"

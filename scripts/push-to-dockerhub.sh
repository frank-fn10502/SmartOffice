#!/usr/bin/env bash
set -euo pipefail

# 推送 image 到 Docker Hub
# 版號從 .env.docker 檔案讀取
# 流程：推送 x86-linux、arm64，並建立 version 與 latest manifest
#
# 用法:
#   ./scripts/push-to-dockerhub.sh

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
DOCKERHUB_DIR="${SCRIPT_DIR}/dockerhub"
source "${DOCKERHUB_DIR}/common.sh"

if ! ENV_FILE="$(resolve_env_file "${REPO_ROOT}")"; then
  echo "錯誤：找不到 ${REPO_ROOT}/scripts/dockerhub/.env.docker"
  echo "請先建立 scripts/dockerhub/.env.docker 並設定 DOCKER_IMAGE_VERSION"
  exit 1
fi

DOCKER_IMAGE_VERSION="$(read_docker_image_version "${ENV_FILE}")"

echo "推送 image 到 Docker Hub (版號: ${DOCKER_IMAGE_VERSION})"
echo "=============================================="
echo ""
echo "設定檔: ${ENV_FILE}"
echo ""

echo "[步驟 1/3] 構建並推送 x86-linux image..."
"${DOCKERHUB_DIR}/build-and-push-x86-linux.sh"
echo ""

echo "[步驟 2/3] 構建並推送 arm64 image..."
"${DOCKERHUB_DIR}/build-and-push-arm64.sh"
echo ""

echo "[步驟 3/3] 建立並推送 manifest (含 latest)..."
"${DOCKERHUB_DIR}/build-and-push-manifest.sh"
echo ""

echo "✓ 完成！image 已推送到 Docker Hub"
echo ""
echo "工作機可以使用以下命令下載："
echo "  ./scripts/pull-from-dockerhub.sh"

#!/usr/bin/env bash

DOCKER_HUB_REPO="${DOCKER_HUB_REPO:-frank10502/smart-office-dev}"

resolve_repo_root() {
  local script_dir
  script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
  cd "${script_dir}/../.." && pwd
}

resolve_env_file() {
  local repo_root="${1}"
  printf '%s\n' "${repo_root}/scripts/dockerhub/.env.docker"
}

read_docker_image_version() {
  local env_file="${1}"
  local version

  if [ ! -f "${env_file}" ]; then
    echo "錯誤：${env_file} 不存在" >&2
    echo "請先建立 .env.docker 檔案並設定 DOCKER_IMAGE_VERSION" >&2
    exit 1
  fi

  if ! version=$(grep '^DOCKER_IMAGE_VERSION=' "${env_file}" | head -n 1 | cut -d '=' -f 2-); then
    echo "錯誤：無法從 ${env_file} 讀取 DOCKER_IMAGE_VERSION" >&2
    exit 1
  fi

  version="$(printf '%s' "${version}" | tr -d '[:space:]')"

  if [ -z "${version}" ]; then
    echo "錯誤：DOCKER_IMAGE_VERSION 未設定或為空" >&2
    exit 1
  fi

  printf '%s\n' "${version}"
}

ensure_buildx_builder() {
  local builder_name="${1:-smartoffice-builder}"

  echo "檢查 docker buildx 支援..."
  if ! docker buildx version >/dev/null 2>&1; then
    echo "錯誤：docker buildx 不可用" >&2
    echo "請確保已安裝 Docker Desktop 或啟用 buildx" >&2
    exit 1
  fi

  if ! docker buildx inspect "${builder_name}" >/dev/null 2>&1; then
    echo "建立新的 builder instance: ${builder_name}"
    docker buildx create --name "${builder_name}" --use
  else
    echo "使用既有的 builder instance: ${builder_name}"
    docker buildx use "${builder_name}"
  fi

  docker buildx inspect --bootstrap >/dev/null
}
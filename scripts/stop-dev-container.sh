#!/usr/bin/env bash
set -euo pipefail

CONTAINER_NAME="${SMARTOFFICE_DEV_CONTAINER:-smartoffice-hub-dev}"

if docker container inspect "${CONTAINER_NAME}" >/dev/null 2>&1; then
  docker rm -f "${CONTAINER_NAME}" >/dev/null
  printf 'Stopped Docker container "%s".\n' "${CONTAINER_NAME}"
else
  printf 'Docker container "%s" is not running.\n' "${CONTAINER_NAME}"
fi

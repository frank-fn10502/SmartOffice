#!/usr/bin/env bash
set -euo pipefail

SKILL_NAME="smartoffice-hub-outlook"
SKILL_ID="smartoffice-hub-outlook.skill.smartoffice-hub.2026-05"
MARKER_FILE=".smartoffice-skill-id"

usage() {
  cat <<'USAGE'
安裝 SmartOffice.Hub Outlook Agents SKILL。

用法:
  install.sh [options]

預設:
  安裝到 user skill folder:
  ${CODEX_HOME:-$HOME/.codex}/skills/smartoffice-hub-outlook

Options:
  --user
      安裝到 user skill folder。這是預設行為。

  --project <path>
      安裝到指定 project 的 .codex/skills folder。
      例如: --project /path/to/project

  --dest <path>
      安裝到指定 skills root 或完整 skill folder。
      若 path basename 是 smartoffice-hub-outlook，會直接使用該 path；
      否則會安裝到 <path>/smartoffice-hub-outlook。

  --force
      保留相容參數；目前安裝預設就是全新重裝。

  --dry-run
      只顯示將會安裝的位置，不寫入檔案。

  -h, --help
      顯示說明。

範例:
  ./scripts/install.sh
  ./scripts/install.sh --project /path/to/project
  ./scripts/install.sh --dest /tmp/codex-skills
USAGE
}

absolute_path() {
  local path="$1"
  if [[ "$path" = /* ]]; then
    printf '%s\n' "$path"
  else
    printf '%s/%s\n' "$PWD" "$path"
  fi
}

resolve_destination() {
  local destination="$1"
  local base
  base="$(basename "$destination")"
  if [[ "$base" == "$SKILL_NAME" ]]; then
    printf '%s\n' "$destination"
  else
    printf '%s/%s\n' "$destination" "$SKILL_NAME"
  fi
}

TARGET_MODE="user"
PROJECT_DIR=""
DEST_ROOT=""
DRY_RUN="false"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --user)
      TARGET_MODE="user"
      PROJECT_DIR=""
      DEST_ROOT=""
      shift
      ;;
    --project)
      if [[ $# -lt 2 ]]; then
        echo "錯誤: --project 需要 path。" >&2
        exit 2
      fi
      TARGET_MODE="project"
      PROJECT_DIR="$2"
      DEST_ROOT=""
      shift 2
      ;;
    --dest)
      if [[ $# -lt 2 ]]; then
        echo "錯誤: --dest 需要 path。" >&2
        exit 2
      fi
      TARGET_MODE="dest"
      DEST_ROOT="$2"
      PROJECT_DIR=""
      shift 2
      ;;
    --force)
      shift
      ;;
    --dry-run)
      DRY_RUN="true"
      shift
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "錯誤: 不支援的參數: $1" >&2
      usage >&2
      exit 2
      ;;
  esac
done

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILL_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

SOURCE_MARKER="$SKILL_DIR/$MARKER_FILE"
if [[ ! -f "$SOURCE_MARKER" ]]; then
  echo "錯誤: source skill 缺少識別檔 $MARKER_FILE。" >&2
  exit 1
fi

SOURCE_SKILL_ID="$(tr -d '\r\n' < "$SOURCE_MARKER")"
if [[ "$SOURCE_SKILL_ID" != "$SKILL_ID" ]]; then
  echo "錯誤: source skill id 不符，預期 $SKILL_ID，實際 $SOURCE_SKILL_ID。" >&2
  exit 1
fi

case "$TARGET_MODE" in
  user)
    CODEX_HOME_DIR="${CODEX_HOME:-$HOME/.codex}"
    DEST_DIR="$CODEX_HOME_DIR/skills/$SKILL_NAME"
    ;;
  project)
    PROJECT_ABS="$(absolute_path "$PROJECT_DIR")"
    DEST_DIR="$PROJECT_ABS/.codex/skills/$SKILL_NAME"
    ;;
  dest)
    DEST_ABS="$(absolute_path "$DEST_ROOT")"
    DEST_DIR="$(resolve_destination "$DEST_ABS")"
    ;;
  *)
    echo "錯誤: 未知 target mode: $TARGET_MODE" >&2
    exit 2
    ;;
esac

echo "來源: $SKILL_DIR"
echo "目標: $DEST_DIR"

if [[ "$DRY_RUN" == "true" ]]; then
  exit 0
fi

if [[ -e "$DEST_DIR" ]]; then
  TARGET_MARKER="$DEST_DIR/$MARKER_FILE"
  if [[ ! -f "$TARGET_MARKER" ]]; then
    echo "錯誤: 目標已存在，但缺少 ${MARKER_FILE}；為避免覆蓋其他同名 skill，已停止。" >&2
    exit 1
  fi

  TARGET_SKILL_ID="$(tr -d '\r\n' < "$TARGET_MARKER")"
  if [[ "$TARGET_SKILL_ID" != "$SKILL_ID" ]]; then
    echo "錯誤: 目標 skill id 不符，為避免覆蓋其他同名 skill，已停止。" >&2
    echo "預期: $SKILL_ID" >&2
    echo "實際: $TARGET_SKILL_ID" >&2
    exit 1
  fi

  echo "移除既有安裝: $DEST_DIR"
  rm -rf "$DEST_DIR"
fi

mkdir -p "$(dirname "$DEST_DIR")"
cp -R "$SKILL_DIR" "$DEST_DIR"

echo "已全新安裝 ${SKILL_NAME}。"

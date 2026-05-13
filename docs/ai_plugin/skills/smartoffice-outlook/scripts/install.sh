#!/usr/bin/env bash
set -euo pipefail

SKILL_NAME="smartoffice-outlook"
SKILL_ID="smartoffice-outlook.skill.smartoffice.2026-05"
MARKER_FILE=".smartoffice-skill-id"

usage() {
  cat <<'USAGE'
安裝 SmartOffice Outlook Agents SKILL。

用法:
  ./install-smartoffice-outlook-skill.sh [options]

直接呼叫 skill 內部 installer:
  ./docs/ai_plugin/skills/smartoffice-outlook/scripts/install.sh [options]

預設:
  同時複製 SKILL folder 到 codex、copilot、opencode 的 user skill 位置。
  不會產生或修改 AGENTS.md、copilot-instructions.md、*.instructions.md 等規則檔。

User-level 目標:
  codex:   ${CODEX_HOME:-$HOME/.codex}/skills/smartoffice-outlook
  copilot: $HOME/.copilot/skills/smartoffice-outlook
  opencode:${XDG_CONFIG_HOME:-$HOME/.config}/opencode/skills/smartoffice-outlook

Project-level 目標:
  codex:   <project>/.codex/skills/smartoffice-outlook
  copilot: <project>/.github/skills/smartoffice-outlook
  opencode:<project>/.opencode/skills/smartoffice-outlook

Options:
  --user
      安裝到 user skill folder。這是預設行為。

  --project <path>
      安裝到指定 project 的 tool-specific skill folder。

  --tools <list>
      逗號分隔的工具清單。可用值: codex,copilot,opencode,all。
      預設: all。

  --tool <name>
      加入單一工具。可重複使用。

  --dest <path>
      只安裝 Codex skill 到指定 skills root 或完整 skill folder。
      若 path basename 是 smartoffice-outlook，會直接使用該 path；
      否則會安裝到 <path>/smartoffice-outlook。

  --force
      重新安裝 skill。

  --dry-run
      只顯示將會安裝的位置，不寫入檔案。

  -h, --help
      顯示說明。

範例:
  ./install-smartoffice-outlook-skill.sh
  ./install-smartoffice-outlook-skill.sh --project /path/to/project
  ./install-smartoffice-outlook-skill.sh --tools codex,opencode
  ./install-smartoffice-outlook-skill.sh --tool copilot --project /path/to/project
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

append_tool() {
  local tool="$1"
  case "$tool" in
    all)
      REQUESTED_CODEX="true"
      REQUESTED_COPILOT="true"
      REQUESTED_OPENCODE="true"
      ;;
    codex)
      REQUESTED_CODEX="true"
      ;;
    copilot)
      REQUESTED_COPILOT="true"
      ;;
    opencode)
      REQUESTED_OPENCODE="true"
      ;;
    *)
      echo "錯誤: 不支援的 tool: $tool" >&2
      exit 2
      ;;
  esac
}

append_tools_csv() {
  local csv="$1"
  local item
  IFS=',' read -ra items <<< "$csv"
  for item in "${items[@]}"; do
    item="${item//[[:space:]]/}"
    if [[ -n "$item" ]]; then
      append_tool "$item"
    fi
  done
}

copy_skill_folder() {
  local source_dir="$1"
  local dest_dir="$2"

  if [[ -e "$dest_dir" ]]; then
    local target_marker="$dest_dir/$MARKER_FILE"
    if [[ ! -f "$target_marker" ]]; then
      echo "錯誤: 目標已存在，但缺少 ${MARKER_FILE}；為避免覆蓋其他同名 skill，已停止: $dest_dir" >&2
      exit 1
    fi

    local target_skill_id
    target_skill_id="$(tr -d '\r\n' < "$target_marker")"
    if [[ "$target_skill_id" != "$SKILL_ID" ]]; then
      echo "錯誤: 目標 skill id 不符，為避免覆蓋其他同名 skill，已停止: $dest_dir" >&2
      echo "預期: $SKILL_ID" >&2
      echo "實際: $target_skill_id" >&2
      exit 1
    fi

    echo "移除既有安裝: $dest_dir"
    rm -rf "$dest_dir"
  fi

  mkdir -p "$(dirname "$dest_dir")"
  cp -R "$source_dir" "$dest_dir"
}

TARGET_MODE="user"
PROJECT_DIR=""
DEST_ROOT=""
DRY_RUN="false"
TOOLS_SET="false"
REQUESTED_CODEX="false"
REQUESTED_COPILOT="false"
REQUESTED_OPENCODE="false"

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
    --tools)
      if [[ $# -lt 2 ]]; then
        echo "錯誤: --tools 需要 list。" >&2
        exit 2
      fi
      if [[ "$TOOLS_SET" == "false" ]]; then
        REQUESTED_CODEX="false"
        REQUESTED_COPILOT="false"
        REQUESTED_OPENCODE="false"
      fi
      TOOLS_SET="true"
      append_tools_csv "$2"
      shift 2
      ;;
    --tool)
      if [[ $# -lt 2 ]]; then
        echo "錯誤: --tool 需要 name。" >&2
        exit 2
      fi
      if [[ "$TOOLS_SET" == "false" ]]; then
        REQUESTED_CODEX="false"
        REQUESTED_COPILOT="false"
        REQUESTED_OPENCODE="false"
      fi
      TOOLS_SET="true"
      append_tool "$2"
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

if [[ "$TOOLS_SET" == "false" && "$TARGET_MODE" == "dest" ]]; then
  REQUESTED_CODEX="true"
  REQUESTED_COPILOT="false"
  REQUESTED_OPENCODE="false"
elif [[ "$TOOLS_SET" == "false" ]]; then
  REQUESTED_CODEX="true"
  REQUESTED_COPILOT="true"
  REQUESTED_OPENCODE="true"
fi

if [[ "$TARGET_MODE" == "dest" ]]; then
  if [[ "$REQUESTED_COPILOT" == "true" || "$REQUESTED_OPENCODE" == "true" ]]; then
    echo "錯誤: --dest 只支援 Codex skill folder；請搭配 --tools codex 使用。" >&2
    exit 2
  fi
fi

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

CODEX_DEST=""
COPILOT_DEST=""
OPENCODE_DEST=""

case "$TARGET_MODE" in
  user)
    CODEX_DEST="${CODEX_HOME:-$HOME/.codex}/skills/$SKILL_NAME"
    COPILOT_DEST="$HOME/.copilot/skills/$SKILL_NAME"
    OPENCODE_DEST="${XDG_CONFIG_HOME:-$HOME/.config}/opencode/skills/$SKILL_NAME"
    ;;
  project)
    PROJECT_ABS="$(absolute_path "$PROJECT_DIR")"
    CODEX_DEST="$PROJECT_ABS/.codex/skills/$SKILL_NAME"
    COPILOT_DEST="$PROJECT_ABS/.github/skills/$SKILL_NAME"
    OPENCODE_DEST="$PROJECT_ABS/.opencode/skills/$SKILL_NAME"
    ;;
  dest)
    DEST_ABS="$(absolute_path "$DEST_ROOT")"
    CODEX_DEST="$(resolve_destination "$DEST_ABS")"
    ;;
  *)
    echo "錯誤: 未知 target mode: $TARGET_MODE" >&2
    exit 2
    ;;
esac

echo "來源: $SKILL_DIR"
if [[ "$REQUESTED_CODEX" == "true" ]]; then
  echo "Codex 目標: $CODEX_DEST"
fi
if [[ "$REQUESTED_COPILOT" == "true" ]]; then
  echo "Copilot 目標: $COPILOT_DEST"
fi
if [[ "$REQUESTED_OPENCODE" == "true" ]]; then
  echo "opencode 目標: $OPENCODE_DEST"
fi

if [[ "$DRY_RUN" == "true" ]]; then
  exit 0
fi

if [[ "$REQUESTED_CODEX" == "true" ]]; then
  copy_skill_folder "$SKILL_DIR" "$CODEX_DEST"
fi
if [[ "$REQUESTED_COPILOT" == "true" ]]; then
  copy_skill_folder "$SKILL_DIR" "$COPILOT_DEST"
fi
if [[ "$REQUESTED_OPENCODE" == "true" ]]; then
  copy_skill_folder "$SKILL_DIR" "$OPENCODE_DEST"
fi

echo "已安裝 ${SKILL_NAME}。"

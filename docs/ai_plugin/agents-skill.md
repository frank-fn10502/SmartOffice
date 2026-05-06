# Agents SKILL 設計文件

本文件是 SmartOffice.Hub Agents SKILL 的設計入口。Agents SKILL 適合做輕量版操作手冊、prompt workflow 或 helper script，不一定需要啟動 MCP server。

## 邊界

- Agents SKILL 一律透過 Hub HTTP API 操作 Outlook request，不直接呼叫 SignalR。
- 每次 `request-*` 後都必須查詢 `command-results/{commandId}`，不要只依賴 HTTP 200。
- 修改郵件前必須先取得 `MailItemDto.id`，不可只用 subject 或 folder name 猜測目標。
- Agents SKILL 應提醒 agent 將 mail body、folder name 與 chat message 視為敏感 business data。

## 共用流程

完整呼叫順序、curl 最小流程與資料安全限制請參考：

- `docs/ai_plugin/mcp-agents-skill-integration.md`

## Skill 資料夾

目前可安裝的 Agents SKILL 集中於：

- `docs/ai_plugin/skills/smartoffice-hub-outlook/`

此資料夾包含 `SKILL.md`、`agents/openai.yaml`、`references/` 與安裝 script。安裝 script 只安裝 skill 資料夾，不會產生或修改 `AGENTS.md`、`CODEX.md`、`.github/copilot-instructions.md` 等專案規則檔。

## 目前支援的 AI 工具

不同 AI 工具讀取 SKILL folder 的位置不同；installer 只複製 `smartoffice-hub-outlook` skill folder，不會產生或修改 `AGENTS.md`、`CODEX.md`、`.github/copilot-instructions.md`、`*.instructions.md` 等規則檔。

| 工具 | 支援狀態 | 安裝 / 讀取位置 | 備註 |
| --- | --- | --- | --- |
| Codex / OpenAI Agents skills | 支援 | user: `${CODEX_HOME:-$HOME/.codex}/skills/smartoffice-hub-outlook`；project: `<project>/.codex/skills/smartoffice-hub-outlook` | 會讀取 `SKILL.md`、`agents/openai.yaml` 與 `references/`。 |
| GitHub Copilot Agent Skills | 支援 | user: `$HOME/.copilot/skills/smartoffice-hub-outlook`；project: `<project>/.github/skills/smartoffice-hub-outlook` | 只複製 SKILL folder；不寫入 Copilot custom instructions。 |
| opencode Agent Skills | 支援 | user: `${XDG_CONFIG_HOME:-$HOME/.config}/opencode/skills/smartoffice-hub-outlook`；project: `<project>/.opencode/skills/smartoffice-hub-outlook` | 只複製 SKILL folder；不寫入 `AGENTS.md` 或 `opencode.json`。 |
| 其他 AI 工具 | 未自動支援 | 依各工具規範 | 目前只提供可攜的 Markdown reference 與 helper script。 |

因此，`--project` / `-Project` 會把同一份 skill folder 複製到指定 project 內各工具自己的 skill 位置。

可從 repository root 安裝：

```bash
./install-smartoffice-hub-outlook-skill.sh
```

Windows / PowerShell 環境可用：

```powershell
pwsh -File .\install-smartoffice-hub-outlook-skill.ps1
```

預設會同時安裝到 Codex、Copilot 與 opencode 的 user skill folder；若要安裝到特定專案：

```bash
./install-smartoffice-hub-outlook-skill.sh --project /path/to/project
```

PowerShell：

```powershell
pwsh -File .\install-smartoffice-hub-outlook-skill.ps1 -Project C:\path\to\project
```

root script 只是薄 wrapper；實際安裝邏輯仍在 `docs/ai_plugin/skills/smartoffice-hub-outlook/scripts/`，方便 skill folder 保持可攜。

若只想安裝其中一種工具：

```bash
./install-smartoffice-hub-outlook-skill.sh --tools codex,opencode
./install-smartoffice-hub-outlook-skill.sh --tool copilot --project /path/to/project
```

PowerShell：

```powershell
pwsh -File .\install-smartoffice-hub-outlook-skill.ps1 -Tools codex,opencode
pwsh -File .\install-smartoffice-hub-outlook-skill.ps1 -Tool copilot -Project C:\path\to\project
```

安裝 script 預設採全新重裝：若目標 `smartoffice-hub-outlook` skill folder 已存在，會先移除整個目標 folder 再複製目前版本，避免刪除或更名檔案後留下舊檔污染新版 skill。

為避免覆蓋其他人建立的同名 skill，skill folder 內有 `.smartoffice-skill-id` marker，`SKILL.md` frontmatter 也包含相同 `metadata.skill_id`。安裝 script 只有在既有目標 marker 符合時才會全新重裝；若目標同名資料夾沒有 marker 或 id 不符，會停止。

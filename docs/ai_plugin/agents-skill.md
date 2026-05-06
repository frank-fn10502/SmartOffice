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

此資料夾包含 `SKILL.md`、`agents/openai.yaml` 與 `references/`。若之後需要安裝 script，建議從這個資料夾複製到 user skill 位置，例如 `${CODEX_HOME:-$HOME/.codex}/skills/smartoffice-hub-outlook`。

可從 repository root 安裝：

```bash
./install-smartoffice-hub-outlook-skill.sh
```

Windows / PowerShell 環境可用：

```powershell
pwsh .\install-smartoffice-hub-outlook-skill.ps1
```

預設安裝到 user skill folder；若要安裝到特定專案：

```bash
./install-smartoffice-hub-outlook-skill.sh --project /path/to/project
```

PowerShell：

```powershell
pwsh .\install-smartoffice-hub-outlook-skill.ps1 -Project C:\path\to\project
```

root script 只是薄 wrapper；實際安裝邏輯仍在 `docs/ai_plugin/skills/smartoffice-hub-outlook/scripts/`，方便 skill folder 保持可攜。

安裝 script 預設採全新重裝：若目標 `smartoffice-hub-outlook` skill folder 已存在，會先移除整個目標 folder 再複製目前版本，避免刪除或更名檔案後留下舊檔污染新版 skill。

為避免覆蓋其他人建立的同名 skill，skill folder 內有 `.smartoffice-skill-id` marker，`SKILL.md` frontmatter 也包含相同 `metadata.skill_id`。安裝 script 只有在既有目標 marker 符合時才會全新重裝；若目標同名資料夾沒有 marker 或 id 不符，會停止。

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

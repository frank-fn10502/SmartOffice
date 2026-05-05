# AI Plugin 文件

本資料夾專門放置給 AI plugin、MCP adapter 與 Agents SKILL 使用的文件。這些文件說明外部 AI client 如何安全地透過 Hub HTTP API 操作 `SmartOffice.Hub`，不應混入工作機 Outlook AddIn 的實作規格。

目前分成兩類：

- `MCP`：把 Hub HTTP workflow 包成 MCP tool 的設計、tool schema 與呼叫順序。
- `Agents SKILL`：給 coding agent 或 AI assistant 使用的操作手冊、helper workflow 與資料安全注意事項。

## 文件

- `docs/ai_plugin/mcp.md`：MCP adapter 與 tool 設計入口。
- `docs/ai_plugin/agents-skill.md`：Agents SKILL 設計入口。
- `docs/ai_plugin/mcp-agents-skill-integration.md`：MCP 與 Agents SKILL 共用的 Hub integration 建議流程。

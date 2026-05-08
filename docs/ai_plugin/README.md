# AI Plugin 文件

本資料夾專門放置給 AI plugin、MCP adapter 與 Agents SKILL 使用的文件。這些文件說明外部 AI client 如何安全地透過 SmartOffice Outlook HTTP API 操作 Outlook，不應混入工作機 Outlook AddIn 的實作規格。

目前分成兩類：

- `MCP`：把 SmartOffice API HTTP workflow 包成 MCP tool 的設計、tool schema 與呼叫順序。
- `Agents SKILL`：給 coding agent 或 AI assistant 使用的操作手冊、helper workflow 與資料安全注意事項。

## 文件

- `docs/ai_plugin/mcp.md`：MCP adapter 與 tool 設計入口。
- `docs/ai_plugin/agents-skill.md`：Agents SKILL 設計入口。
- `docs/ai_plugin/mcp-agents-skill-integration.md`：MCP 與 Agents SKILL 共用的 SmartOffice API integration 建議流程。
- `docs/ai_plugin/acceptance-scenarios.md`：開發驗收用的使用者情境清單，不屬於可安裝 SKILL 內容。
- `docs/ai_plugin/skills/smartoffice-outlook/`：可安裝的 SmartOffice Outlook Agents SKILL 資料夾，包含 bash 與 PowerShell 安裝 script。

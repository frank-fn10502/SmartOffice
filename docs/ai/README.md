# AI 協作文件

本資料夾是給本 repository 的 AI coding agent 與維護者看的內部文件，包含專案邊界、coding 規範、前端限制、後端擴充原則與驗證方式。

如果你要實作 Outlook AddIn，請先看：

- `../SmartOffice/docs/outlook-addin/README.md`
- `../SmartOffice/docs/outlook-addin/features-checklist.md`
- `../SmartOffice/docs/outlook-addin/signalr-contract.md`

如果你要擴充 Hub 後端支援新的 Office AddIn domain，請先看：

- `backend.md`

如果你要讓 AI client 透過 MCP 或 Agents SKILL 操作 Hub，請看：

- `docs/ai_plugin/README.md`
- `docs/ai_plugin/mcp.md`
- `docs/ai_plugin/agents-skill.md`
- `docs/ai_plugin/mcp-agents-skill-integration.md`

`docs/ai/` 不再放 AddIn 實作者的正式 contract，也不放專門給 MCP 或 Agents SKILL 的 plugin 文件；避免 AI 在內部協作文件、AddIn 規格與 AI plugin 設計之間來回跳轉。

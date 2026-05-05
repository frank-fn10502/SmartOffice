# MCP 設計文件

本文件是 SmartOffice.Hub MCP adapter 的設計入口。MCP server 應保持薄層，只負責把 MCP tool call 轉成 Hub HTTP API call，並等待 `command-results/{commandId}` 回報。

## 邊界

- MCP 不直接連 Outlook、Office COM 或 `/hub/outlook-addin`。
- MCP 不保存長期 mailbox state；Hub cached snapshot 仍是主要讀取來源。
- MCP tool schema 應保守，優先沿用 Hub request DTO，避免 adapter 發明另一套 contract。
- MCP 回覆不應大量輸出 mail body、folder name 或 chat message，這些內容都可能含有敏感 business data。

## 共用流程

完整呼叫順序、建議 tool 清單與 progress endpoint 請參考：

- `docs/ai_plugin/mcp-agents-skill-integration.md`

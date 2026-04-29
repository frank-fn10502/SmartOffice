# Project Notes for AI Agents

## 專案使命

SmartOffice.Hub 是 Office 2016 Add-in、Web UI 與 AI/MCP client 之間的本機中介層。請維持這個邊界：

- Add-in 負責 Office automation 與 Office-specific integration。
- Hub 負責 HTTP API、SignalR notification、command routing、temporary state，以及未來 AI/MCP integration。
- Web UI 負責使用者檢視、手動 request、chat 與 diagnostics。

修改時請偏好小而明確的 contract，避免隱性耦合。Office 2016 與受限企業環境是設計約束，不只是 implementation detail。

## Repository Layout

- `Program.cs`：application startup 與 dependency registration。
- `Controllers/`：HTTP API boundary。Office-specific route 要保持命名清楚。
- `Hubs/`：用於 browser live update 的 SignalR hub。
- `Models/`：Hub、Add-in、Web UI 與可能的 MCP client 共用的 DTO。
- `Services/`：in-memory store、queue 與 application service。
- `webui/`：Vue 3 + Vite Web UI source。
- `wwwroot/`：ASP.NET Core static file root；目前保留既有 dashboard，Vue build 暫時輸出到 `wwwroot/dist/`。

## 目前技術選擇

- ASP.NET Core on .NET 8。
- SignalR 用於 dashboard real-time update。
- Swagger through Swashbuckle。
- Web UI 採用 Vue 3 + Vite + Element Plus，production output 仍是 static assets。
- Prototype 階段使用 in-memory store。

除非任務明確需要，請不要引入 database、Nuxt、Vue Router、Pinia、background job framework 或 AI SDK。

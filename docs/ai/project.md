# Project Notes for AI Agents

## 專案使命

SmartOffice.Hub 是 Office 2016 Add-in、Web UI 與 AI/MCP client 之間的本機中介層。請維持這個邊界：

- Add-in 負責 Office automation 與 Office-specific integration。
- Hub 負責 HTTP API、SignalR notification、command routing、temporary state，以及未來 AI/MCP integration。
- Web UI 負責使用者檢視、手動 request、chat 與 diagnostics；Web UI 應優先走 HTTP API，作為 Hub API 的手動測試工具。

修改時請偏好小而明確的 contract，避免隱性耦合。Office 2016 與受限企業環境是設計約束，不只是 implementation detail。

## Solution 關係

本 repository 是 `SmartOffice.Hub` 專案，負責 Hub contract、Web UI、mock、文件與本機驗證。完整 SmartOffice solution 會用相對路徑參考：

```text
..\SmartOffice.Hub\SmartOffice.Hub.csproj
```

這代表：

- Hub 專案是 AddIn 的 API contract 與本機服務參考。
- Outlook COM/VSTO automation、連線 Hub SignalR、接收 `OutlookCommand`、回報 folders/mails/rules/calendar、修改 mail/category/folder 等實作位於 SmartOffice / Outlook AddIn 專案。
- 修改 Hub contract、HTTP API、SignalR command/result 或 mock 時，請同步檢查 AddIn 文件與 Web UI。
- 涉及真實 Outlook COM/VSTO 行為時，仍需 Windows / Outlook / Office 環境編譯與實測；Hub/mock/Web UI 可先在本機驗證 contract 與使用者流程。

## Repository Layout

- `Program.cs`：application startup 與 dependency registration。
- `Controllers/`：HTTP API boundary。Office-specific route 要保持命名清楚。
- `Hubs/`：用於 notification 與 Outlook AddIn command/result channel 的 SignalR hub；Web UI 不應把 notification 當主要資料來源。
- `Models/`：Hub、Add-in、Web UI 與可能的 MCP client 共用的 DTO。
- `Services/`：in-memory store、SignalR command dispatcher 與 application service。
- `webui/`：Vue 3 + Vite Web UI source。Outlook domain 程式集中在 `webui/src/features/outlook/`，不要把 Outlook-specific api、model、component、composable、util 分散回 `webui/src` 根層資料夾。
- `wwwroot/`：ASP.NET Core static file root；Vue build 會輸出 `index.html` 與 `assets/` 到這裡。

## 目前技術選擇

- ASP.NET Core on .NET 8。
- SignalR 用於 Outlook AddIn command/result channel，以及 status、log、progress notification；Web UI 的 folders、mails、rules、categories、calendar、chat 等資料應由 HTTP data endpoint 讀回。
- Swagger through Swashbuckle。
- Web UI 採用 Vue 3 + Vite + Element Plus，production output 仍是 static assets。
- Prototype 階段使用 in-memory store。

除非任務明確需要，請不要引入 database、Nuxt、Vue Router、Pinia、background job framework 或 AI SDK。

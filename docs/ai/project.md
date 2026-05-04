# Project Notes for AI Agents

## 專案使命

SmartOffice.Hub 是 Office 2016 Add-in、Web UI 與 AI/MCP client 之間的本機中介層。請維持這個邊界：

- Add-in 負責 Office automation 與 Office-specific integration。
- Hub 負責 HTTP API、SignalR notification、command routing、temporary state，以及未來 AI/MCP integration。
- Web UI 負責使用者檢視、手動 request、chat 與 diagnostics。

修改時請偏好小而明確的 contract，避免隱性耦合。Office 2016 與受限企業環境是設計約束，不只是 implementation detail。

## 雙電腦 / 雙 Solution 關係

本 repository 是開發機上的 `SmartOffice.Hub` 專案，負責 Hub contract、Web UI 與文件。這台開發機不一定有工作機的完整 SmartOffice / Outlook AddIn solution。

工作機上才有完整 SmartOffice solution 與 Outlook AddIn 專案。該 solution 會用相對路徑參考：

```text
..\SmartOffice.Hub\SmartOffice.Hub.csproj
```

這代表：

- Hub 專案是工作機 AddIn 的 API contract 與本機服務參考。
- Outlook COM/VSTO automation、連線 Hub SignalR、接收 `OutlookCommand`、回報 folders/mails/rules/calendar、修改 mail/category/folder 等實作，必須在工作機 SmartOffice / Outlook AddIn 專案中完成。
- 在本 repository 修改 `Controllers/`、`Models/`、`Services/` 只適合 Hub contract 真的不足時進行。
- `Plan/` 內的任務預設是交給工作機 AI 在完整 SmartOffice solution 中執行，不是要求本 Hub repo 的 agent 修改 Hub 程式碼。

## Repository Layout

- `Program.cs`：application startup 與 dependency registration。
- `Controllers/`：HTTP API boundary。Office-specific route 要保持命名清楚。
- `Hubs/`：用於 browser live update 與 Outlook AddIn command/result channel 的 SignalR hub。
- `Models/`：Hub、Add-in、Web UI 與可能的 MCP client 共用的 DTO。
- `Services/`：in-memory store、SignalR command dispatcher 與 application service。
- `webui/`：Vue 3 + Vite Web UI source。
- `wwwroot/`：ASP.NET Core static file root；Vue build 會輸出 `index.html` 與 `assets/` 到這裡。

## 目前技術選擇

- ASP.NET Core on .NET 8。
- SignalR 用於 dashboard real-time update，以及 Outlook AddIn command/result channel。
- Swagger through Swashbuckle。
- Web UI 採用 Vue 3 + Vite + Element Plus，production output 仍是 static assets。
- Prototype 階段使用 in-memory store。

除非任務明確需要，請不要引入 database、Nuxt、Vue Router、Pinia、background job framework 或 AI SDK。

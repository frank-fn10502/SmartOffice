# SmartOffice.Hub

SmartOffice.Hub 是 Office 2016 Add-in、Web UI 與 AI/MCP tooling 之間的本機中介服務。這個專案主要服務於受限的 Windows office 工作環境，讓桌面版 Office 不需要直接與 cloud AI service 溝通，也能取得 AI 協助。

目前實作重點放在 Outlook。未來 Word、Excel、PowerPoint 或其他 Office Add-in 可以沿用同一個 Hub pattern。

## 專案目的

Hub 位在三種角色中間：

- Office Add-in：開啟聊天室窗、讀取 Office context、推送結果，並透過 polling 取得命令。
- Web UI：讓使用者檢視 Outlook data、發出操作要求、聊天，以及監看 Add-in 狀態。
- AI/MCP client：透過 Hub API 讀取 Office context，或要求 Add-in 執行操作。

這個設計讓 Office 2016 automation 維持在本機且明確可控。Add-in 負責 Office COM/VSTO interaction，Hub 負責 API boundary、command routing、real-time UI notification 與暫存狀態。

## 目前架構

```text
Web UI / AI / MCP
       |
       | REST + SignalR
       v
SmartOffice.Hub
       |
       | long-poll commands + push results
       v
Office 2016 Add-in
```

重要檔案：

- `Program.cs`：ASP.NET Core startup、CORS、Swagger、static files、SignalR 與 in-memory service 註冊。
- `Controllers/OutlookController.cs`：Web UI、AI/MCP client 與 Outlook Add-in 使用的 REST API。
- `Hubs/NotificationHub.cs`：提供 Web UI real-time update 的 SignalR endpoint。
- `Services/Stores.cs`：in-memory mail/folder/chat/status store 與 command queue。
- `Models/Dtos.cs`：Hub、Web UI、Add-in 之間共用的 DTO contract。
- `wwwroot/`：靜態 dashboard，用於 mail browsing、chat 與 Add-in diagnostics。

## 執行模型

Web UI 與 AI/MCP 端透過下列 endpoint enqueue command：

- `POST /api/outlook/request-folders`
- `POST /api/outlook/request-mails`

Outlook Add-in 透過 long-poll 取得 pending command：

- `GET /api/outlook/poll`

Add-in 執行完本機 Office automation 後，將結果推回 Hub：

- `POST /api/outlook/push-folders`
- `POST /api/outlook/push-mails`

Web UI 透過 SignalR 接收更新：

- `/hub/notifications`
- 事件包含 `FoldersUpdated`、`MailsUpdated`、`NewChatMessage`、`AddinStatus`、`AddinLog`。

## 本機執行

需求：

- .NET 8 SDK

啟動 Hub：

```bash
dotnet run
```

預設 `http` launch profile 目前監聽：

```text
http://localhost:2805
```

預設 profile 不啟用 Hub 端 Add-in mock，適合在工作電腦由 Outlook Add-in 直接連線。離線開發時可改用 mock profile：

```bash
dotnet run --launch-profile http-mock
```

常用網址：

- Dashboard：`http://localhost:2805/`
- Swagger：`http://localhost:2805/swagger`

## 開發模式

本專案支援三種開發方式。

### 本機模式 Host Mode

使用本機已安裝的 .NET SDK：

```bash
dotnet run
dotnet build
```

如果本機已經有相容的 .NET 8 SDK，這是最直接的模式。

### 快速模式 Quick Mode

Quick Mode 保持 editor 與日常開發環境在本機，只把 compilation 放進暫存 Docker container。

```bash
./scripts/build-in-container.sh
```

這是目前偏好的 build workflow，適合不想在本機安裝或維護 .NET SDK 的情境。腳本會在需要時從 `.devcontainer/Dockerfile` 建立 reusable local image，接著用暫存 container 執行 compilation。build container 結束後會被移除。這個腳本只做 build 驗證，不會改變 runtime profile；實際 F5 是否連真 Add-in 由 `Properties/launchSettings.json` 的 profile 決定。

可以調整 local image tag 或 build configuration：

```bash
SMARTOFFICE_BUILD_IMAGE=smartoffice-hub-devcontainer:local CONFIGURATION=Release ./scripts/build-in-container.sh
```

### 完整容器模式 Full Container Mode

可選的 `.devcontainer` 資料夾讓 VS Code 將整個 workspace 重新開在 .NET 8 development container 裡。

當你希望 editor terminal、SDK 與 C# tooling 都在 Docker 裡執行時，使用這個模式。devcontainer 會使用 `.devcontainer/Dockerfile`，讓未來 native package 與 tooling 可以集中維護。

devcontainer 不會自動執行 `dotnet restore`。restore 與 run command 需要手動執行，避免開啟 container 時意外下載 package。

請參考 `.devcontainer/README.md`。

## API 說明

Outlook route prefix：

```text
/api/outlook
```

主要 Web UI / AI request endpoint：

- `POST /request-folders`：enqueue folder fetch command。
- `POST /request-mails`：enqueue mail fetch command。
- `GET /folders`：讀取 cached folders。
- `GET /mails`：讀取 cached mails。
- `POST /chat`：新增並 broadcast chat message。
- `GET /chat`：讀取 cached chat messages。

主要 Add-in endpoint：

- `GET /poll`：long-poll 取得一筆 pending command，timeout 為 30 秒。
- `POST /push-folders`：取代 cached folders 並 broadcast update。
- `POST /push-mails`：取代 cached mails 並 broadcast update。

Admin endpoint：

- `GET /admin/status`
- `GET /admin/logs`
- `POST /admin/log`

## 安全假設 Security Assumptions

目前專案假設執行在可信任的本機或 intranet 環境：

- CORS 允許任意 origin 搭配 credentials。
- Swagger 目前總是啟用。
- Data 只存在 process-local memory。
- 目前尚未加入 authentication / authorization。

如果要放到受控 workstation 或 lab network 以外的環境，請先加入 authentication、限制 CORS、決定 Swagger 是否只在 development 啟用，並檢查 mail content 是否能暴露給 AI/MCP client。

## 後續方向

近期適合投入的方向：

- 加入 provider-agnostic AI service layer。
- 加入面向 MCP 的 endpoint/tool，用於 Office context 與 command dispatch。
- 當更多 Add-in 出現時，拆出 Office-specific API，例如 `/api/word`、`/api/excel`、`/api/powerpoint`。
- 如果狀態需要跨 process restart 保留，加入 durable storage 或 bounded cache。
- 加入 command correlation 與 completion/error reporting，讓 Web UI 與 AI client 可以追蹤 request end-to-end。
- 補上 command queue behavior、DTO contract 與 controller response 的 tests。

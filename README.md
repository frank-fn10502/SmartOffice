# SmartOffice.Hub

SmartOffice.Hub 是 Office 2016 Add-in、Web UI 與外部整合工具之間的本機中介服務。這個專案主要服務於受限的 Windows office 工作環境，讓桌面版 Office 不需要直接與 cloud AI service 溝通，也能取得 AI 協助。

目前實作重點放在 Outlook。未來 Word、Excel、PowerPoint 或其他 Office Add-in 可以沿用同一個 Hub pattern；未來也可讓支援 MCP 或其他協定的外部工具連接到這個 Hub。

## 專案目的

Hub 位在三種角色中間：

- Office Add-in：開啟聊天室窗、讀取 Office context，並透過 SignalR 接收命令與回報結果。
- Web UI：讓使用者檢視 Outlook data、發出操作要求、聊天，以及監看 Add-in 狀態。
- 外部整合工具：透過 Hub 提供的 API、MCP 或其他整合方式讀取 Office context，或要求 Add-in 執行操作。

這個設計讓 Office 2016 automation 維持在本機且明確可控。Add-in 負責 Office COM/VSTO interaction，Hub 負責 API boundary、command routing、real-time UI notification 與暫存狀態。

## 目前架構

```text
Web UI / 外部工具
       |
       | HTTP / SignalR / MCP
       v
SmartOffice.Hub
       |
       | SignalR commands + results
       v
Office 2016 Add-in
```

重要檔案：

- `Program.cs`：ASP.NET Core startup、CORS、Swagger、static files、SignalR 與 in-memory service 註冊。
- `Controllers/OutlookController.cs`：Web UI 與外部整合工具使用的 REST API。
- `Hubs/NotificationHub.cs`：提供 Web UI real-time update 的 SignalR endpoint。
- `Hubs/OutlookAddinHub.cs`：提供 Outlook AddIn 正式 SignalR command/result channel。
- `Services/Stores.cs`：in-memory mail/folder/chat/status store 與 SignalR command dispatcher。
- `Models/Dtos.cs`：Hub、Web UI、Add-in 之間共用的 DTO contract。
- `wwwroot/`：靜態 dashboard，用於 mail browsing、chat 與 Add-in diagnostics。

## 執行模型

Web UI 與外部整合工具可透過下列 endpoint 發出 request；Hub 會透過 SignalR dispatch command 給 Outlook AddIn：

- `POST /api/outlook/request-folders`
- `POST /api/outlook/request-mails`

Outlook AddIn 透過 SignalR 取得 command：

- `/hub/outlook-addin`
- client event：`OutlookCommand`

AddIn 執行完本機 Office automation 後，透過 SignalR server method 將結果回報 Hub：

- `BeginFolderSync`
- `PushFolderBatch`
- `CompleteFolderSync`
- `PushMails`
- `PushRules`
- `PushCategories`
- `PushCalendar`
- `ReportAddinLog`
- `ReportCommandResult`

Web UI 透過 SignalR 接收更新：

- `/hub/notifications`
- 事件包含 `FolderSyncStarted`、`FoldersPatched`、`FolderSyncCompleted`、`MailsUpdated`、`NewChatMessage`、`AddinStatus`、`AddinLog`。

## 預設開發與執行方式

目前專案的標準流程是：

- 本機編輯程式碼
- 用 container 編譯 Vue/Vite 與 .NET
- 用 container 執行 SmartOffice.Hub

也就是說，開發環境的重點不是先在每台機器上安裝完整的前端與 .NET 建置工具，而是優先使用專案提供的 Docker image。

### 1. 先從 Docker Hub 下載 image

預設 Docker Hub repository：`frank10502/smart-office-dev`

```bash
./scripts/pull-from-dockerhub.sh
```

這個 script 會下載專案使用的 build image，並同步拉取 `latest` tag。image 提供 `linux/amd64` 與 `linux/arm64` multi-arch manifest，Docker 會自動選擇可用版本。

### 2. 用 container 編譯前端與 .NET

```bash
./scripts/build-in-container.sh
```

這個 script 會在 container 內完成整個建置流程：

- 處理 `webui` 的 npm dependencies
- 執行 Vue/Vite build，產出 static html 到 `wwwroot`
- 執行 `dotnet build`

如果本機還沒有對應的 image，script 會依照 `.devcontainer/Dockerfile` 自動建立。

`.devcontainer/Dockerfile` 目前固定的主要工具鏈是：

- .NET 8 SDK
- Node.js 22 LTS

如需調整 image tag 或 build configuration，可使用：

```bash
SMARTOFFICE_BUILD_IMAGE=smartoffice-hub-devcontainer:local CONFIGURATION=Release ./scripts/build-in-container.sh
```

### 3. 用 container 執行 Hub

```bash
./scripts/start-dev-container.sh
```

這個 script 會先確認前端 static 檔案與 build 結果是否可用，必要時會先執行 `./scripts/build-in-container.sh`，然後再用 container 啟動 Hub。

預設會使用 `Mock` environment 啟動，適合開發時檢查 Web UI、API 與 SignalR endpoint；Outlook 資料仍需工作機 AddIn 透過 `/hub/outlook-addin` 回推。

常用網址：

- Dashboard：`http://localhost:2805/`
- Swagger：`http://localhost:2805/swagger`

停止 container：

```bash
./scripts/stop-dev-container.sh
```

如果希望連 editor terminal、SDK 與工具鏈都放在 container 內，也可以使用 `.devcontainer`。詳細說明請參考 `.devcontainer/README.md`。

## 本機開發（需要自行處理環境）

如果不使用上面的 container 流程，而要改成純本機開發，就必須自己把開發機環境補齊。

請至少比照 `.devcontainer/Dockerfile` 準備：

- .NET 8 SDK
- Node.js 22 LTS
- `webui` 所需的 npm dependencies

本機最常見的方式，通常是：

- 用 Visual Studio 20xx 開啟專案處理 .NET / ASP.NET Core
- 另外執行 `./scripts/build-in-container.sh`，先把前端編譯成 `wwwroot` 下的 static html

換句話說，就算是本機開發，實務上也常常仍然會用 build script 來處理前端產物，再交給 Visual Studio 20xx 處理 .NET 端。

## API 說明

Outlook route prefix：

```text
/api/outlook
```

主要 Web UI / 外部整合工具 request endpoint：

- `POST /request-folders`：dispatch folder fetch command。
- `POST /request-mails`：dispatch mail fetch command。
- `GET /folders`：讀取 cached folders。
- `GET /mails`：讀取 cached mails。
- `POST /chat`：新增並 broadcast chat message。
- `GET /chat`：讀取 cached chat messages。
- `GET /command-results/{commandId}`：讓 AI / MCP client 等待 AddIn command result。
- `GET /command-results`：讀取最近 command result。

主要 AddIn SignalR endpoint：

- `/hub/outlook-addin`：正式 Outlook AddIn command/result channel。

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

如果要放到受控 workstation 或 lab network 以外的環境，請先加入 authentication、限制 CORS、決定 Swagger 是否只在 development 啟用，並檢查 mail content 是否能暴露給外部整合工具或其他連接端。

## AI / MCP / SKILL 整合

AI client 建議不要直接連 Outlook AddIn SignalR channel，而是透過 MCP tool 或 SKILL helper 呼叫 Hub HTTP API。標準流程是 dispatch `/api/outlook/request-*`、取得 `commandId`、輪詢 `/api/outlook/command-results/{commandId}`，完成後再讀取 `/api/outlook/mails`、`/folders`、`/calendar` 等 cached snapshot。

詳細 tool 設計與 SKILL 寫法請看 `docs/ai/mcp-skill-integration.md`。

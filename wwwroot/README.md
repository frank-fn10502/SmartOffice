# SmartOffice Web UI

這個資料夾包含 SmartOffice.Hub 提供的 static dashboard。

UI 刻意保持 lightweight，因為目標環境可能受限，未必允許 npm、bundler 或頻繁 dependency change。

## 檔案

```text
wwwroot/
├── index.html       # Dashboard markup 與 page-level JavaScript
├── styles.css       # Dashboard styling
└── folder-tree.js   # Outlook folder tree rendering logic
```

## 職責

- 向 Hub request Outlook folder 與 mail fetch。
- 顯示 cached Outlook folders 與 mails。
- 發送與接收 chat message。
- 顯示 Outlook Add-in connection status 與 logs。
- 訂閱 SignalR event 以接收 real-time update。

## 外部執行期相依 External Runtime Dependency

頁面會從 CDN 載入 SignalR：

```html
https://cdnjs.cloudflare.com/ajax/libs/microsoft-signalr/8.0.0/signalr.min.js
```

如果部署環境需要完全 offline 或 intranet-only，請將此檔案 vendor 到本機，並更新 `index.html`。

## API 使用方式 API Usage

UI 使用 same-origin endpoint 與 Hub 溝通：

- `POST /api/outlook/request-folders`
- `POST /api/outlook/request-mails`
- `GET /api/outlook/folders`
- `GET /api/outlook/mails`
- `GET /api/outlook/chat`
- `POST /api/outlook/chat`
- `GET /api/outlook/admin/status`
- `GET /api/outlook/admin/logs`

SignalR endpoint：

- `/hub/notifications`

UI 會消費的 SignalR event：

- `FoldersUpdated`
- `MailsUpdated`
- `NewChatMessage`
- `AddinStatus`
- `AddinLog`

## 開發筆記

- 除非專案明確移往 frontend build pipeline，否則請保持 UI dependency-light。
- 不要將 secret 或 AI provider key 放在 client-side file。
- rendered mail content 視為敏感資料。目前 HTML mail view 使用 iframe 隔離顯示，但 iframe 不是完整 sanitization boundary。

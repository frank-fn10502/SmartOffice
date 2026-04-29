# SmartOffice Web UI

這個資料夾包含 SmartOffice.Hub 提供的 static dashboard 與前端 build output。

既有 `index.html` 仍是目前 dashboard 入口。新的 Vue 3 + Vite Web UI source 放在 `webui/`，導入階段的 build output 會輸出到 `wwwroot/dist/`。

## 檔案

```text
wwwroot/
├── index.html       # Dashboard markup 與 page-level JavaScript
├── styles.css       # Dashboard styling
└── folder-tree.js   # Outlook folder tree rendering logic
└── dist/            # Vue/Vite build output，產生檔不 commit
```

## 職責

- 向 Hub request Outlook folder 與 mail fetch。
- 顯示 cached Outlook folders 與 mails。
- 發送與接收 chat message。
- 顯示 Outlook Add-in connection status 與 logs。
- 訂閱 SignalR event 以接收 real-time update。

## Frontend Build

新的 Vue Web UI 使用：

```bash
cd webui
npm run build
```

Vue 版本的 SignalR client 應透過 npm dependency bundle，不再從 CDN 載入。

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

- Web UI 已明確採用 Vue 3 + Vite + Element Plus，但仍請保持 dependency-light。
- 不要預設加入 Nuxt、Vue Router、Pinia、Axios、Tailwind 或第二套 UI kit。
- 不要將 secret 或 AI provider key 放在 client-side file。
- rendered mail content 視為敏感資料。目前 HTML mail view 使用 iframe 隔離顯示，但 iframe 不是完整 sanitization boundary。

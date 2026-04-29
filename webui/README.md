# SmartOffice.Hub Web UI

這個資料夾是新的 Vue 3 + Vite Web UI source。

它取代了既有手寫 `wwwroot/index.html` dashboard。

## Stack

- Vue 3
- Vite
- Element Plus
- npm

## 指令

不要在 host 直接執行 npm 指令。啟動可瀏覽的 Hub + Web UI：

```bash
./scripts/start-dev-container.sh
```

開啟：

```text
http://localhost:2805/
```

停止：

```bash
./scripts/stop-dev-container.sh
```

需要互動式前端開發時，請在 devcontainer 裡執行：

```bash
npm install
npm run dev
npm run build
```

development server 預設監聽：

```text
http://localhost:5173/
```

Vite 會 proxy `/api` 與 `/hub` 到：

```text
http://localhost:2805
```

production build output 目前輸出到：

```text
../wwwroot/
```

`wwwroot/` 是 build output 目錄，Vite build 會清空後重建。
`start-dev-container.sh` 會在缺少 `webui/node_modules/` 或 `wwwroot/index.html` 時自動準備。`webui/node_modules/` 會留在 workspace 作為 cache，並由 `.gitignore` 排除，不要 commit。

## Dependency 原則

- `node_modules/` 不可 commit。
- 不要在 host 直接執行 npm install/build。
- 優先使用 Vue Composition API 與 project-local components。
- UI component 優先使用 Element Plus，不加入第二套 UI kit。
- API request 優先使用原生 `fetch`。
- SignalR client 使用 `@microsoft/signalr` npm package，不使用 CDN。

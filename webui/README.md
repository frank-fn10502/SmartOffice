# SmartOffice.Hub Web UI

這個資料夾是新的 Vue 3 + Vite Web UI source。

目前它只是前端 toolchain baseline，尚未取代既有 `wwwroot/index.html` dashboard。

## Stack

- Vue 3
- Vite
- Element Plus
- npm

## 指令

不要在 host 直接執行 npm 指令。日常驗證請從 repository root 執行：

```bash
./scripts/build-in-container.sh
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
../wwwroot/dist/
```

## Dependency 原則

- `node_modules/` 不可 commit。
- 不要在 host 直接執行 npm install/build。
- 優先使用 Vue Composition API 與 project-local components。
- UI component 優先使用 Element Plus，不加入第二套 UI kit。
- API request 優先使用原生 `fetch`。
- SignalR client 使用 `@microsoft/signalr` npm package，不使用 CDN。

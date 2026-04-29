# Frontend Framework Decision

## 選擇

Web UI 採用 `Vue 3 + Vite + Element Plus`。

不採用 `Nuxt`、`Vue Router`、`Pinia` 或大型前端 application framework 作為預設。只有當 UI 真的出現多頁 routing、跨頁 global state 或複雜 client workflow 時，才重新評估。

## 目的

這個選擇的目標是讓 Web UI 從目前的 static HTML/CSS/JavaScript 演進到較穩定的 component model，同時保持：

- 完全本機編譯。
- production output 仍是靜態 HTML/CSS/JavaScript。
- ASP.NET Core 繼續從 `wwwroot` 提供 Web UI。
- 前端 dependency 有明確邊界，不讓套件生態自然膨脹。

## UI Framework

採用 `Element Plus` 作為 Vue UI component framework。

選擇理由：

- 原生支援 Vue 3。
- 文件完整，元件覆蓋 dashboard、form、table、tree、dialog、notification 等常見管理介面需求。
- 風格偏企業工具與 admin UI，符合 SmartOffice.Hub 的 dashboard / diagnostics 使用情境。
- 支援 on-demand import，可避免一開始 full import 整套樣式與元件。

限制：

- 不要預設導入第二套 UI kit。
- 不要預設導入 Tailwind、Bootstrap、Material Design 或 CSS utility framework。
- 如果需要 icons，優先使用 `@element-plus/icons-vue` 或專案本地 SVG，不額外加入大型 icon library。

## 專案結構

前端 source code 放在：

```text
webui/
```

目前導入階段的 build output 放在：

```text
wwwroot/dist/
```

這是為了避免尚未移植完成時覆蓋既有 `wwwroot/index.html`。等 Vue UI 完成並切換為正式入口後，再調整 Vite `outDir` 或 ASP.NET Core static file routing。

## Dependency 原則

- `node_modules/` 不可 commit。
- 不要在 host 直接執行 `npm install`、`npm run build` 或其他 npm script。
- npm install/build 必須透過 devcontainer、`./scripts/build-in-container.sh`，或明確的 Docker container 執行。
- 使用 npm lockfile 固定 dependency resolution。
- SignalR client 不再從 CDN 載入；改使用 npm dependency 並 bundle 進本機 build output。
- API 使用原生 `fetch`，不要預設加入 Axios 或 data fetching framework。
- 優先使用 Vue built-in Composition API 管理 local state，不要預設加入 Pinia。
- 優先使用 project-local components，不要因單一畫面需求引入大型附加套件。

## 預期指令

```bash
./scripts/build-in-container.sh
```

Docker Quick Mode 會在 `webui/package.json` 存在時執行前端 build，再執行 .NET build。腳本會把 `webui/node_modules` 掛到 Docker volume，避免 npm packages 寫入 repository 工作目錄。

需要互動式前端開發時，請在 devcontainer 內執行：

```bash
cd webui
npm install
npm run dev
```

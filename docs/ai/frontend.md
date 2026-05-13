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

build output 放在：

```text
wwwroot/
```

`wwwroot/` 是 Vue/Vite 產物目錄，build 時會清空後重建。不要在 `wwwroot/` 放手寫檔案，包括 README；修改 UI 時請改 `webui/src/`。

Outlook domain code 採 feature folder：

```text
webui/src/features/outlook/
  api/
  components/
  composables/
  models/
  styles/
  utils/
```

`webui/src` 根層只保留 app shell、views 與 global style。Outlook-specific 檔案不要再新增到根層 `api/`、`components/`、`composables/`、`models/` 或 `utils/`；若未來新增其他 domain，請建立新的 `features/<domain>/`，不要把不同 domain 混在同一個泛用資料夾。

## 檔案切分原則

前端要避免單一檔案長期膨脹。修改 `webui/src/` 時，行數限制是硬性規則：

- `webui/src/` 內 `.ts` 與 `.vue` 檔案不得超過 800 行。
- 這條規則由 `webui/scripts/check-file-lines.mjs` 檢查，`npm run build` 會先執行 `npm run check:file-lines`。
- `./scripts/build-in-container.sh` 也會經過同一個 gate；若任一 `.ts` 或 `.vue` 超過 800 行，build 必須失敗。
- 600 行以上視為預警區。繼續新增功能前，請先找自然切分點；接近 800 行時不得再把新功能塞進同一檔案。
- 先抽出純資料與 pure helper，例如 enum mapping、formatting、normalizer、date helper、color helper。
- UI 很自然成為獨立區塊時，再拆成 component；例如 folder tree、mail row、category editor、calendar grid。
- CSS 可依畫面區塊或 feature 拆分，但不要把每個 selector 拆成獨立檔案；以「能一次理解一個 UI 區塊」為準。
- 不要為了追求短檔案而建立大量只有 trivial function 的檔案。偏好少量、穩定、命名清楚的模組。

若當次任務會讓檔案明顯變長，應同時安排小幅切分。超過 800 行時不能以 change summary 說明原因代替切分；必須先修到 gate 通過。完成 Web UI 變更時，回報中需列出 line-count gate 或 container build 結果。

## Dependency 原則

- `node_modules/` 不可 commit。
- 不要在 host 直接執行 `npm install`、`npm run build` 或其他 npm script。
- npm install/build 必須透過 devcontainer、`./scripts/build-in-container.sh`，或明確的 Docker container 執行。
- `webui/node_modules/` 可留在 workspace 作為 Docker 內 npm install/cache 的結果，但必須由 `.gitignore` 排除，不可 commit。
- 使用 npm lockfile 固定 dependency resolution。
- SignalR client 不再從 CDN 載入；改使用 npm dependency 並 bundle 進本機 build output。
- Web UI 是 Hub HTTP API 的手動測試工具。SignalR notification 只能用於 AddIn status、log、progress 或 diagnostics 顯示；folders、mails、folder mail results、mail body、attachments、search results、rules、categories、calendar 與 chat 等資料必須透過 HTTP request / command result / data endpoint 更新。
- API 使用原生 `fetch`，不要預設加入 Axios 或 data fetching framework。
- 優先使用 Vue built-in Composition API 管理 local state，不要預設加入 Pinia。
- 優先使用 project-local components，不要因單一畫面需求引入大型附加套件。

## 預期指令

```bash
./scripts/start-dev-container.sh
```

`start-dev-container.sh` 是人類使用的主要入口；缺少 `webui/node_modules/` 或 `wwwroot/index.html` 時會自動呼叫 Docker build flow。這個 script 預設使用 `ASPNETCORE_ENVIRONMENT=Mock`，方便沒有 Office Add-in 時檢查 Web UI；需要連真 Add-in 時可用 `SMARTOFFICE_ASPNETCORE_ENVIRONMENT=Production ./scripts/start-dev-container.sh`。`build-in-container.sh` 保留給 CI 或只需要 build 驗證的情境。

需要互動式前端開發時，請在 devcontainer 內執行：

```bash
cd webui
npm install
npm run dev
```

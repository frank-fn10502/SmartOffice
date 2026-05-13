# Validation Notes for AI Agents

## Quick Mode

偏好的驗證模式是在本機編輯，使用暫存 Docker container 編譯：

```bash
./scripts/build-in-container.sh
```

如果 `webui/package.json` 存在，Quick Mode 會先執行前端 build，再執行 .NET build。
不要在 host 直接執行 npm 指令；前端 npm install/build 必須在 Docker/devcontainer 內執行。

## Web UI 行數 Gate

`webui/src/` 內 `.ts` 與 `.vue` 檔案不得超過 800 行。這是硬性驗收規則，任何 Web UI 變更都必須經過檢查。

Preferred validation：

```bash
./scripts/build-in-container.sh
```

此流程會執行 `npm run check:file-lines`。若只需要在 devcontainer 內快速檢查行數，可執行：

```bash
cd webui
npm run check:file-lines
```

若 gate 失敗，必須先依自然職責切分檔案，再繼續 build 或回覆使用者。不要把「超過 800 行」列為已知問題留到之後處理。

## Host Mode

如果本機已安裝 .NET 8 SDK，也可以執行：

```bash
dotnet build
```

## API 檢查

如果修改 API behavior，也請透過 Swagger 或 `.http` example 檢查。

常用網址：

```text
http://localhost:2805/swagger
```

## Mock-first Contract 驗證

若修改的功能最後需要 `SmartOffice/OutlookAddIn` 真實 VSTO/COM 實作，請先在 `SmartOffice.Hub` 完成並驗證 mock workflow：

1. 實作 Hub 前，先查 Microsoft 官方文件確認 Outlook/Office API 概念可行；若是已知高風險行為，例如 conversation、search、大量 mail/folder 枚舉、rule、attachment export、COM lifetime 或 Office UI thread，需在回報中附上官方依據。官方文件不足時，才補查 Microsoft Q&A、Stack Overflow 或 issue，且清楚標示為社群經驗。
2. 更新 Hub DTO、HTTP route、SignalR command/result contract、store、mock backend、Web UI 與文件。
3. 執行 `./scripts/build-in-container.sh`，確認前端 build 與 .NET build 都通過。
4. 用 `./scripts/start-dev-container.sh` 啟動 Mock 環境，透過 Swagger、`curl` 或 `.http` 檔檢查實際對外 HTTP API：request endpoint、paired `fetch-result-*` endpoint、必要的 `GET` data endpoint、錯誤狀態與分頁資訊都要能被非 Web UI caller 清楚理解。
5. API smoke test 最低標準：response 必須包含可追蹤的 `requestId`、穩定的 `request` 名稱、可判斷的 `state`、可讀的 `message`、明確命名的 `data` 欄位，以及必要時的 `next.cursor` / `next.hasMore`。如果只看 raw JSON 很難知道下一步、資料含義、錯誤原因或是否還有下一頁，請修正 API，而不是只讓 Web UI 補說明。
6. 從使用者角度操作 Web UI：確認主要入口找得到、按鈕狀態合理、loading/empty/error state 不重疊、文字不擠出容器、dialog 或 list 不出現明顯破版。此階段不是追求完美，只要求沒有最基礎的 UI 錯誤。
7. 確認 `docs/ai/protocols.md` 與 `../SmartOffice/docs/outlook-addin/signalr-contract.md` 描述的 command type、DTO 欄位、route 與 mock 回傳一致。
8. 回頭檢查 Hub/mock 是否能替 Add-in 減輕負擔：能由 Hub 合併、分頁、slice、節流或以單一 command 表達的工作，不應改成一次 dispatch 數百或數千個 Outlook command；mock 應能呈現足夠資料量與邊界情境來驗證這個設計。
9. 同步更新 Agents SKILL：`docs/ai_plugin/skills/smartoffice-outlook/SKILL.md`、`references/http-api.md`、`references/workflows.md`，必要時更新 `docs/ai_plugin/acceptance-scenarios.md` 與 `docs/ai_plugin/mcp-agents-skill-integration.md`。外部 AI 主要只讀 SKILL folder；如果只更新內部 AGENTS/docs，外部 AI 仍會照舊 contract 操作。
10. 完成上述檢查後，才修改 `../SmartOffice/OutlookAddIn` 的 Outlook COM/VSTO 真實實作；本機回報必須標註該部分仍需 Windows 主機編譯/實測。

## API 可理解性檢查

Hub HTTP API 是 Web UI、AI、MCP client 與其他工具的共同入口。修改 API behavior 時，不能只檢查 Web UI 畫面是否能用，也要直接看 raw HTTP response：

- request endpoint 要明確告訴 caller 這是 accepted、running、completed、unavailable、timeout 還是 failed。
- paired `fetch-result-*` 要能用同一個 `requestId` 查回狀態與資料，不需要 caller 猜內部 command。
- `data` 物件內的集合名稱要清楚，例如 `mails`、`folders`、`attachments`、`calendarEvents`，避免只回匿名 array 或意義不明的 payload。
- 分頁結果要提供 `next.cursor` 與 `next.hasMore`；沒有下一頁時 cursor 應為空字串。
- 錯誤 response 要有穩定 `status` 或 `state`、可讀 `message`，必要時附上 caller 可採取的下一步，例如重試、先載入 folders、縮小 scope 或改用另一個 endpoint。
- 若為了讓 Web UI 好看而把 API 設計得難以單獨理解，應優先修 API contract，Web UI 只負責呈現，不負責補足 API 語意。

## Web UI 檢查

如果修改 Web UI，請用 Docker 啟動 Hub 後檢查：

```bash
./scripts/start-dev-container.sh
```

這個 script 預設載入 `appsettings.Mock.json`，讓 Web UI 檢查不依賴真實 Office Add-in。

```text
http://localhost:2805/
```

檢查完請停止 container：

```bash
./scripts/stop-dev-container.sh
```

只需要 build 驗證時，請執行：

```bash
./scripts/build-in-container.sh
```

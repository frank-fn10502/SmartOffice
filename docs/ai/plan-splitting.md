# Plan 任務切分與工作機 AI 交付規範

本文件定義未來將 `Plan/` 任務切分給另一台主機 AI 時的粒度、必要文件、狀態追蹤與驗證要求。目標是讓任務可以被 VS Code Copilot custom agent、工作機 AI 或其他 coding agent 穩定地逐項執行。

## 適用範圍

`Plan/` 內的任務預設是給工作機完整 SmartOffice / Outlook AddIn solution 執行。這個 repository 是 `SmartOffice.Hub`，只包含 Hub、Web UI、contract 與 mock。

除非任務明確標示需要修改 Hub contract、API、Web UI、mock 或文件，否則切分後的任務不應要求修改本 repository 的程式碼。

## 任務大小

每個 Plan 任務應符合：

- 單一目標：只完成一個 AddIn 能力、Hub contract 調整、Web UI 顯示或文件產出。
- 單一主要驗證：完成後可以用一組明確步驟判定成功或失敗。
- 可在一個 agent session 內完成；若需要跨多個大型 class、跨 Hub 與 AddIn、或需要人工 Outlook 操作多輪，應再切小。
- 不超過 8 個實作步驟；若超過，拆成多個依序任務。
- 不要求 agent 同時設計 contract、實作 AddIn、改 Web UI 與寫完整測試；這些應拆開。

## 任務檔案格式

新增或重寫 Plan 任務時，使用以下章節：

```md
# Task NNN：任務名稱

## 執行位置

說明要在工作機 SmartOffice solution、SmartOffice.Hub，或兩者哪一邊執行。

## 目標

用 1 到 3 句話描述完成後的使用者可觀察結果。

## 必讀

- `..\SmartOffice.Hub\AGENTS.md`
- 依任務需要列出 Hub 文件、contract、controller、DTO 或工作機狀態檔。

## 輸入

列出 command type、endpoint、DTO、現有 class、設定值或前置任務。

## 實作步驟

1. 每一步只做一件事。
2. 步驟要能讓小模型照著執行。

## 驗證

1. 明確列出啟動、操作與預期結果。
2. 若需要人工 Outlook 操作，寫出匿名化測試資料要求。

## 完成後更新

說明要更新哪個 status 檔、任務狀態與下一個任務。
```

## 必要文件

切分給工作機 AI 的任務，至少應要求讀取：

- `..\SmartOffice.Hub\AGENTS.md`
- `..\SmartOffice.Hub\docs\ai\workstation-solution.md`
- `..\SmartOffice.Hub\docs\ai\protocols.md`
- `..\SmartOffice.Hub\docs\ai\office2016-workstation-contract.md`
- 當次 `..\SmartOffice.Hub\Plan\NNN-*.md`
- 工作機 `Plan\WORKSTATION-STATUS.md`

若任務涉及 Hub contract 或 DTO，額外加入：

- `..\SmartOffice.Hub\Models\Dtos.cs`
- `..\SmartOffice.Hub\Controllers\OutlookController.cs`
- `..\SmartOffice.Hub\Services\OutlookCommandQueue.cs`

若任務涉及 Web UI，額外加入：

- `..\SmartOffice.Hub\webui\src\models\outlook.ts`
- `..\SmartOffice.Hub\webui\src\App.vue`
- 相關 component 檔案

## 狀態追蹤

Hub repository 使用：

```text
Plan/status.md
```

工作機 SmartOffice solution 使用：

```text
Plan/WORKSTATION-STATUS.md
```

工作機狀態檔至少包含：

- Hub URL。
- AddIn 專案名稱。
- command handler 位置。
- Outlook automation 主要 class。
- 已完成任務列表。
- 目前任務。
- 下一個任務。
- blocker 與驗證紀錄。

## 切分流程

當使用者要求「切分 Plan 給另一台主機 AI」時：

1. 先讀取 `AGENTS.md`、本文件與現有 `Plan/`。
2. 判斷任務應在 Hub repo 或工作機 SmartOffice solution 執行。
3. 將大型需求拆成依賴順序清楚的小任務。
4. 每個任務都寫清楚必讀文件、輸入、步驟、驗證與完成後更新。
5. 更新 `Plan/status.md`，依執行順序列出所有任務。
6. 若任務需要工作機才能完成，明確寫 `不要修改 SmartOffice.Hub 程式碼`。
7. 若任務需要先改 Hub contract，再讓工作機實作，拆成兩個任務：先 Hub contract，後工作機 AddIn。

## 驗證原則

Hub repo 的預設驗證：

```bash
./scripts/build-in-container.sh
```

如果本機已安裝 .NET SDK，也可使用：

```bash
dotnet build
```

工作機 AddIn 的驗證應以任務檔案明確列出，不要只寫「測試通過」。至少包含：

- Hub 是否啟動。
- Outlook AddIn 是否啟動。
- 觸發哪個 request 或 command。
- 預期 push 回 Hub 的 endpoint 或 UI 狀態。
- 失敗時要記錄哪些匿名化 diagnostics。

## 敏感資料

任務與驗證紀錄不得包含真實：

- mail body。
- folder name。
- rule name。
- calendar subject。
- attendee。
- 客戶名稱。
- 公司內部資訊。

需要範例時使用匿名化資料，例如 `Sample Folder A`、`sample@example.com`、`Test Meeting`。

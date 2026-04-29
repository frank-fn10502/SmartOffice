# Task 010：AI Suggestion Confirm 後才執行

## 新 Session 起手

本任務可以在全新 session 單獨執行。請先讀：

1. `AGENTS.md`
2. `Plan/000-session-handoff.md`
3. `Models/Dtos.cs`
4. `Services/Stores.cs`
5. `Controllers/OutlookController.cs`
6. `webui/src/App.vue`
7. 本檔

本任務依賴 `009-ai-suggestion-format.md` 以及至少一個低風險操作 command 已完成。如果缺少 suggestion endpoint 或操作 command，請停止並回報缺少前置任務。

## 目標

讓使用者可以在 Web UI 確認 AI suggestion，確認後才 enqueue Outlook command。

## 前置任務

需要先完成：

- `009-ai-suggestion-format.md`
- 至少一個操作 command，例如 `006-mail-marker-commands.md` 或 `007-folder-create-and-move-mail.md`

## 建議流程

1. AI 建立 suggestion。
2. Web UI 顯示 suggestion。
3. 使用者點開 preview。
4. 使用者按 Confirm。
5. Hub 將 suggestion 轉成對應 `PendingCommand`。
6. Outlook Add-in poll 到 command 並執行。
7. Add-in 回報 command result。

## Hub 實作步驟

1. 新增 endpoint：

```http
POST /api/outlook/ai/suggestions/{id}/confirm
```

2. endpoint 讀取 suggestion。
3. 驗證 `proposedCommandType` 是否在允許清單。
4. 建立 `PendingCommand`。
5. 將 suggestion 標記為 confirmed。

## 注意事項

- 不允許 AI suggestion 執行未知 command type。
- 第一版只支援低風險 command。
- 不支援 delete、send mail、批次移動。

## 驗證

1. 建立一筆低風險 suggestion。
2. Web UI confirm。
3. 確認 Add-in poll 到 command。
4. 確認 command result 回報成功或失敗。

## 完成回報

請回報 confirm endpoint、允許的 command type 清單、Web UI confirmation 行為，以及 build 結果。

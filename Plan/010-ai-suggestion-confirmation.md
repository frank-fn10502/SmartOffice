# Task 010：AI Suggestion Confirm 後才執行

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

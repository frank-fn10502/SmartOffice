# Task 010：工作機補齊 Mail Metadata

## 這個任務的定位

本任務在公司電腦的 Outlook Add-in 補齊 mail metadata。這是線性 Plan 的最後一步。

## 開始前必讀

- `AGENTS.md`
- `Plan/STATUS.md`
- `Plan/CONTRACT-INVENTORY.md`
- 本檔
- `docs/ai/office2016-workstation-contract.md`
- `Models/Dtos.cs`

## 目標

讓工作機推送 mails 時，除了 body，也回傳 Web UI 與 AI 操作需要的 metadata。

## 需要補的欄位

```json
{
  "id": "...",
  "categories": "Customer, Follow-up",
  "isRead": false,
  "isMarkedAsTask": true,
  "importance": "high",
  "sensitivity": "normal"
}
```

## 實作步驟

1. 在工作機既有 fetch mail flow 找到轉換 `MailItem` 的地方。
2. `id` 優先使用 Outlook `EntryID`；取不到就留空。
3. `categories` 對應 Outlook `Categories`。
4. `isRead` 由 `UnRead` 反向轉換。
5. `isMarkedAsTask` 對應 `IsMarkedAsTask`。
6. `importance` 轉成 `low`、`normal`、`high`。
7. `sensitivity` 轉成 `normal`、`personal`、`private`、`confidential`。
8. 保持既有 `body` 與 `bodyHtml` 欄位，不要 rename。
9. 更新 `Plan/STATUS.md`。

## 注意事項

- 如果某欄位取不到，使用預設值，不要讓整批 mails 失敗。
- 不要把完整 mail body 寫進 log。
- `bodyHtml` 是 best-effort，不保證保留 Outlook 視覺樣式。

## 驗證

1. Web UI Fetch Mails。
2. 切到 `Outlook` 分頁。
3. 確認 unread、flagged、importance、categories 統計有合理數字。

## 更新 STATUS

- `010-workstation-mail-metadata` 改成 `done`，或標記 `blocked` 並說明沒有工作機 repo。
- 下一個任務改成 `全部 Plan 任務完成`。

## 完成時請回報

- 每個 metadata 欄位對應的 Outlook property。
- 取不到時的預設值。
- 匿名化測試結果。

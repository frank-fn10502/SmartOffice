# Task 004：AddIn 補齊 Mail Metadata

## 執行位置

本任務請在工作機的完整 SmartOffice / Outlook AddIn solution 中執行。

不要修改 `SmartOffice.Hub` 程式碼。

## 目標

讓 `push-mails` payload 包含 Web UI 與未來 AI 操作需要的 metadata。

## 需要輸出的欄位

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

## 對應建議

- `id`：Outlook `EntryID`，取不到則空字串。
- `categories`：Outlook `Categories`。
- `isRead`：由 Outlook `UnRead` 反向轉換。
- `isMarkedAsTask`：Outlook `IsMarkedAsTask`，取不到則 `false`。
- `importance`：轉成 `low`、`normal`、`high`。
- `sensitivity`：轉成 `normal`、`personal`、`private`、`confidential`。

## 實作步驟

1. 找到 mail 轉 DTO 的程式。
2. 補齊上述欄位。
3. 每個欄位都要 safe read，避免 COM exception 讓整封轉換失敗。
4. 保持既有 JSON field 不 rename。

## 驗證

1. Web UI Fetch Mails。
2. 切到 Outlook 分頁。
3. 確認 unread、flagged、importance、categories 統計有合理數字。

## 完成後更新

更新工作機 `Plan/WORKSTATION-STATUS.md`：

- `004-addin-mail-metadata` 改為 done。
- 下一個任務改為 `005-addin-fetch-rules.md`。

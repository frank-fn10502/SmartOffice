# Task 003：AddIn 實作 Fetch Mails

## 執行位置

本任務請在工作機的完整 SmartOffice / Outlook AddIn solution 中執行。

不要修改 `SmartOffice.Hub` 程式碼。

## 目標

AddIn 收到 `fetch_mails` 後，依 `folderPath`、`range`、`maxCount` 讀取郵件並 push 到 Hub。

## Hub Contract

Poll 會收到：

```json
{
  "type": "fetch_mails",
  "mailsRequest": {
    "folderPath": "\\\\Mailbox - User\\Inbox",
    "range": "1d",
    "maxCount": 10
  }
}
```

完成後呼叫：

```http
POST /api/outlook/push-mails
```

## 實作步驟

1. 在 command handler 加入 `fetch_mails`。
2. 使用 `folderPath` 定位 Outlook folder。
3. 依 `range` 過濾時間：
   - `1d`
   - `1w`
   - `1m`
4. 限制最多 `maxCount` 封。
5. 轉成 Hub 的 `MailItemDto` JSON。
6. `bodyHtml` 取不到時可留空，Web UI 會 fallback 到 `body`。
7. POST `/api/outlook/push-mails`。

## 注意事項

- 不要把真實 mail body 寫入 log。
- HTML 樣式只能 best-effort，不保證還原 Outlook 畫面。
- 單封 mail 轉換失敗時，盡量跳過該封並記錄匿名化錯誤，不要讓整批失敗。

## 驗證

1. Web UI 選 folder。
2. 按 Fetch Mails。
3. 確認郵件列表顯示。
4. 確認 `body` fallback 正常。

## 完成後更新

更新工作機 `Plan/WORKSTATION-STATUS.md`：

- `003-addin-fetch-mails` 改為 done。
- 下一個任務改為 `004-addin-mail-metadata.md`。

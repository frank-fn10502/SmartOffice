# Task 006：新增 Mail Marker 操作 Commands

## 目標

讓 Web UI 或 AI 可以要求工作機修改郵件標記，但第一版只做單封郵件。

## 建議 Command Types

- `mark_mail_read`
- `mark_mail_unread`
- `mark_mail_task`
- `clear_mail_task`
- `set_mail_categories`

## 建議 Request DTO

```json
{
  "mailId": "...",
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "categories": "Customer, Follow-up"
}
```

## Hub 實作步驟

1. 在 `Models/Dtos.cs` 新增 `MailMarkerCommandRequest`。
2. 在 `PendingCommand` 新增 `MailMarkerRequest`。
3. 在 `Controllers/OutlookController.cs` 新增 request endpoints，例如：
   - `POST /api/outlook/request-mark-mail-read`
   - `POST /api/outlook/request-set-mail-categories`
4. endpoint 只 enqueue command，不直接改 Hub cache。
5. 工作機完成後重新 push mails 或回報 command result。

## 工作機實作步驟

1. 用 `mailId` 優先定位 Outlook item。
2. 如果 `mailId` 無法使用，再用 `folderPath` 加其他條件做保守定位。
3. 修改 `UnRead`、`Categories` 或 task flag。
4. 呼叫 `Save()`。
5. 回報 command result。

## 注意事項

- 第一版不要支援批次。
- 找不到唯一郵件時不要修改，回報失敗。
- category 名稱可能含敏感專案名稱，log 要匿名化。

## 驗證

1. 用 Web UI 或 curl enqueue 單封標記 command。
2. 工作機執行後重新 Fetch Mails。
3. 確認 Web UI 顯示已變更。

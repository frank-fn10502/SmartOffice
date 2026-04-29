# Task 009：AddIn 實作移動單封郵件

## 執行位置

本任務請在工作機的完整 SmartOffice / Outlook AddIn solution 中執行。

不要修改 `SmartOffice.Hub` 程式碼。

## 目標

AddIn 支援 Hub 發出的 `move_mail` command。

## Request Shape

```json
{
  "mailId": "...",
  "sourceFolderPath": "\\\\Mailbox - User\\Inbox",
  "destinationFolderPath": "\\\\Mailbox - User\\Projects\\Sample"
}
```

## 實作步驟

1. 在 command handler 加入 `move_mail`。
2. 使用 `mailId` 優先定位 mail。
3. 使用 `destinationFolderPath` 定位 destination folder。
4. 找不到唯一 mail 或 destination folder 時拒絕操作。
5. 呼叫 Outlook `Move()`。
6. 成功後 push mails 或 folders，讓 Web UI 更新。

## 注意事項

- 第一版只支援單封郵件。
- 不要批次移動。
- 不要刪信。
- 不要只靠 subject 定位郵件。

## 驗證

1. enqueue `move_mail`。
2. 確認 Outlook 中郵件移到目標 folder。
3. Web UI 重新 Fetch Mails 後結果正確。

## 完成後更新

更新工作機 `Plan/WORKSTATION-STATUS.md`：

- `009-addin-move-mail` 改為 done。
- 下一個任務改為 `010-addin-command-result-log.md`。

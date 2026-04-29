# Task 007：AddIn 實作單封郵件標記

## 執行位置

本任務請在工作機的完整 SmartOffice / Outlook AddIn solution 中執行。

不要修改 `SmartOffice.Hub` 程式碼。

## 目標

AddIn 支援 Hub 發出的單封郵件標記 command。

## Command Types

- `mark_mail_read`
- `mark_mail_unread`
- `mark_mail_task`
- `clear_mail_task`
- `set_mail_categories`

## 實作步驟

1. 在 command handler 加入上述 command type。
2. 使用 `mailId` 優先定位 Outlook item。
3. 找不到唯一 mail 時不要修改，回報失敗 log。
4. 修改 `UnRead`、task flag 或 `Categories`。
5. 呼叫 `Save()`。
6. 操作完成後建議重新 push mails，讓 Web UI 更新。

## 注意事項

- 第一版只支援單封郵件。
- 不要批次修改。
- 不要直接依 subject 定位郵件。
- category 名稱可能含敏感資料，log 要匿名化。

## 驗證

1. 從 Web UI 或 curl enqueue 一個標記 command。
2. 確認 Outlook 中該封郵件狀態改變。
3. 重新 Fetch Mails，確認 Web UI 更新。

## 完成後更新

更新工作機 `Plan/WORKSTATION-STATUS.md`：

- `007-addin-mark-mail` 改為 done。
- 下一個任務改為 `008-addin-create-folder.md`。

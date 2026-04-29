# Task 010：AddIn 回報 Command Result / Admin Log

## 執行位置

本任務請在工作機的完整 SmartOffice / Outlook AddIn solution 中執行。

不要修改 `SmartOffice.Hub` 程式碼。

## 目標

讓 AddIn 每次執行 command 後，都能回報成功或失敗摘要。若 Hub 尚未提供 command result endpoint，先使用既有：

```http
POST /api/outlook/admin/log
```

## 實作步驟

1. 在 command handler 外層包成功/失敗處理。
2. 成功時記錄 command type 與匿名化摘要。
3. 失敗時記錄 command type 與錯誤類型。
4. 不要把 mail body、folder full path、rule name、calendar subject 寫入 log。
5. 若 Hub 有 command result endpoint，優先使用；否則使用 admin log。

## 建議 Log

```json
{
  "level": "info",
  "message": "Command fetch_calendar completed. Items: 3"
}
```

```json
{
  "level": "error",
  "message": "Command move_mail failed. Reason: target folder not found"
}
```

## 驗證

1. 執行至少三種 command。
2. Web UI Admin 分頁可看到匿名化 log。
3. 故意給錯 folder path，確認 failure log 不含敏感資料。

## 完成後更新

更新工作機 `Plan/WORKSTATION-STATUS.md`：

- `010-addin-command-result-log` 改為 done。
- 下一個任務改為 `全部工作機 AddIn 任務完成`。

# Task 001：AddIn 確認 Hub Poll Command

## 執行位置

本任務請在工作機的完整 SmartOffice / Outlook AddIn solution 中執行。

不要修改 `SmartOffice.Hub` 程式碼。

## 目標

確認 AddIn 可以呼叫：

```http
GET /api/outlook/poll
```

並能依照 `type` 分派 command。

## 必讀

- 工作機 `Plan/WORKSTATION-STATUS.md`
- `..\SmartOffice.Hub\docs\ai\protocols.md`
- `..\SmartOffice.Hub\docs\ai\office2016-workstation-contract.md`

## 實作步驟

1. 找到 AddIn 目前與 Hub 溝通的 HTTP client。
2. 確認 Hub URL 可以設定，不要 hard-code 到無法切換。
3. 實作或修正 long-poll 呼叫 `/api/outlook/poll`。
4. 處理 `{ "type": "none" }`。
5. 對未知 command type 寫匿名化 log，不要 crash。
6. 在 command handler 中保留後續任務會用到的分派結構。

## 驗證

1. 啟動 Hub。
2. 啟動 Outlook AddIn。
3. 確認 Hub Admin status 顯示 AddIn 有 poll。
4. 無 command 時不應造成錯誤。

## 完成後更新

更新工作機 `Plan/WORKSTATION-STATUS.md`：

- `001-addin-poll-command` 改為 done。
- 下一個任務改為 `002-addin-fetch-folders.md`。

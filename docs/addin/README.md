# AddIn 實作者文件

本資料夾是給工作機 Outlook AddIn 實作者與工作機 AI 看的文件。請從 checklist 開始，不要直接跳到 DTO 細節。

## 建議閱讀順序

1. `features-checklist.md`：Web UI 需要 AddIn 實作的功能、完成定義與驗收項目。
2. `signalr-contract.md`：SignalR method、command payload、DTO 欄位與 JSON 範例。
3. `outlook-references.md`：需要確認 Outlook / Office 2016 行為時再查看的官方文件入口。
4. `test-report.md`：工作機測到差異、錯誤或真實資料形狀時的回報格式。

## 使用原則

- checklist 是任務入口；contract 是欄位規格。
- AddIn 不應用 mock data 反推 Outlook object model 行為。
- 真實 mail body、folder name、PST path、category name 與 chat message 都可能含敏感 business data；回報時必須匿名化。
- 郵件列表採兩段式載入：`fetch_mails` 只回 metadata，不應載入或回推完整 `body` / `bodyHtml`；使用者點開單封郵件時，Web UI 會送 `fetch_mail_body`，AddIn 再回推該封內容。

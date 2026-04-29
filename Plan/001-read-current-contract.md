# Task 001：讀懂目前 Outlook Hub Contract

## 新 Session 起手

本任務可以在全新 session 單獨執行。請先讀：

1. `AGENTS.md`
2. `Plan/000-session-handoff.md`
3. 本檔

本任務只做理解與筆記，不修改程式碼。

## 目標

先讓實作者理解目前 Hub、Web UI、Outlook Add-in 的邊界，不修改程式碼。

## 必讀檔案

- `AGENTS.md`
- `docs/ai/project.md`
- `docs/ai/protocols.md`
- `docs/ai/office2016-workstation-contract.md`
- `Models/Dtos.cs`
- `Controllers/OutlookController.cs`

## 需要確認的重點

1. Web UI、AI 或 MCP client 只能透過 `/api/outlook/request-*` enqueue command。
2. Outlook Add-in 只能透過 `/api/outlook/poll` 取得 command。
3. Outlook Add-in 執行本機 Office automation 後，再透過 `/api/outlook/push-*` 回傳結果。
4. Hub 只保存 temporary state，不引入 database。
5. mail body、folder name、chat message 都可能含有敏感 business data。

## 交付內容

建立一份工作機實作筆記，列出：

- 目前已支援的 command type。
- 目前已支援的 push endpoint。
- 哪些欄位可能因 Office 2016 或公司環境而取不到。

## 驗證

不需要 build。只需要確認筆記沒有包含真實客戶資料或信件內容。

## 完成回報

請依 `Plan/000-session-handoff.md` 的完成回報格式回覆，並附上筆記檔案路徑。

# Task 007：建立 Folder 與移動單封郵件

## 新 Session 起手

本任務可以在全新 session 單獨執行。請先讀：

1. `AGENTS.md`
2. `Plan/000-session-handoff.md`
3. `Models/Dtos.cs`
4. `Controllers/OutlookController.cs`
5. `Services/MockAddins/OutlookMockAddinWorker.cs`
6. `docs/ai/protocols.md`
7. `docs/ai/office2016-workstation-contract.md`
8. 本檔

如果 `005-command-result-log.md` 尚未完成，仍可實作 enqueue command；完成回報中註明 command result 尚未接上。

## 目標

支援低風險的 folder 操作：建立 folder、移動單封郵件。

## 建議 Command Types

- `create_folder`
- `move_mail`

## 建議 Request DTO

建立 folder：

```json
{
  "parentFolderPath": "\\\\Mailbox - User\\Projects",
  "name": "Customer A"
}
```

移動 mail：

```json
{
  "mailId": "...",
  "sourceFolderPath": "\\\\Mailbox - User\\Inbox",
  "destinationFolderPath": "\\\\Mailbox - User\\Projects\\Customer A"
}
```

## Hub 實作步驟

1. 新增 request DTO。
2. 新增 `PendingCommand` 欄位。
3. 新增 request endpoints。
4. endpoint 只 enqueue，不直接改 cache。

## 工作機實作步驟

1. `create_folder`：定位 parent folder，確認同名 folder 不存在，再建立。
2. `move_mail`：定位單封 mail 與 destination folder。
3. 找不到唯一 mail 時不要移動。
4. 成功後回報 command result。
5. 成功後建議重新 push folders 或 mails。

## 注意事項

- 不要刪除 folder。
- 不要移動整個 folder。
- 不要批次移動多封郵件。
- folder name 可能含敏感資料，log 要匿名化。

## 驗證

1. 建立測試 folder。
2. 移動一封測試郵件。
3. 在 Outlook client 確認郵件位置。
4. Web UI 重新 Fetch Folders / Fetch Mails。

## 完成回報

請回報新增的 command type、request endpoint、DTO、mock 行為，以及哪些副作用有使用者確認保護。

# Task 004：Hub 新增 Folder 與 Move Mail Commands

## 這個任務的定位

本任務只做 Hub command contract，不做 Web UI preview，也不做工作機 Outlook 實作。

## 開始前必讀

- `AGENTS.md`
- `Plan/STATUS.md`
- `Plan/CONTRACT-INVENTORY.md`
- 本檔
- `Models/Dtos.cs`
- `Controllers/OutlookController.cs`
- `Services/MockAddins/OutlookMockAddinWorker.cs`
- `docs/ai/protocols.md`
- `docs/ai/office2016-workstation-contract.md`

## 目標

支援低風險 folder 操作 command：建立 folder、移動單封郵件。

## Command Types

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

## 實作步驟

1. 在 `Models/Dtos.cs` 新增 request DTO。
2. 在 `PendingCommand` 新增對應 request 欄位。
3. 在 `Controllers/OutlookController.cs` 新增 request endpoints。
4. endpoint 只 enqueue，不直接改 Hub cache。
5. mock worker 收到 command 時，先記錄 command result 或 admin log。
6. 更新 protocol 與 workstation contract 文件。
7. 更新 `Plan/STATUS.md`。

## 注意事項

- 不要刪除 folder。
- 不要移動整個 folder。
- 不要批次移動多封郵件。
- folder name 可能含敏感資料，測試與 log 要匿名化。

## 驗證

執行：

```bash
./scripts/build-in-container.sh
```

用匿名化 JSON 呼叫 request endpoint，確認 `/api/outlook/poll` 可取到 command。

## 更新 STATUS

- `004-hub-folder-mail-commands` 改成 `done`。
- 下一個任務改成 `Plan/005-webui-action-preview.md`。

## 完成時請回報

- 新增的 command type。
- 新增的 endpoint。
- 是否有接上 command result log。
- build 結果。

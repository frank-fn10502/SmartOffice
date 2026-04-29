# Task 003：Hub 新增 Mail Marker Commands

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

讓 Web UI 或 AI 可以 enqueue 單封郵件標記 command。實際修改 Outlook 狀態由未來工作機任務處理。

## Command Types

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

## 實作步驟

1. 在 `Models/Dtos.cs` 新增 `MailMarkerCommandRequest`。
2. 在 `PendingCommand` 新增 `MailMarkerRequest`。
3. 在 `Controllers/OutlookController.cs` 新增 request endpoints。
4. endpoint 只 enqueue command，不直接改 Hub cache。
5. mock worker 收到這些 command 時，先記錄 command result 或 admin log，不需要真的改 mock mails。
6. 更新 protocol 與 workstation contract 文件。
7. 更新 `Plan/STATUS.md`。

## 注意事項

- 第一版只支援單封郵件。
- 找不到 command result store 時，不要自行重做上一個任務；改用 admin log 並在完成回報註明。
- 不要做批次操作。

## 驗證

執行：

```bash
./scripts/build-in-container.sh
```

用匿名化 JSON 呼叫其中一個 request endpoint，確認 `/api/outlook/poll` 可取到 command。

## 更新 STATUS

- `003-hub-mail-marker-commands` 改成 `done`。
- 下一個任務改成 `Plan/004-hub-folder-mail-commands.md`。

## 完成時請回報

- 新增的 command type。
- 新增的 endpoint。
- 是否有接上 command result log。
- build 結果。

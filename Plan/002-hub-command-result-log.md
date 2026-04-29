# Task 002：Hub 新增 Command Result Log

## 這個任務的定位

這是 Hub 端第一個實作任務。它提供後續操作 command 的共同結果回報機制。

## 開始前必讀

- `AGENTS.md`
- `Plan/STATUS.md`
- `Plan/CONTRACT-INVENTORY.md`
- 本檔
- `Models/Dtos.cs`
- `Services/Stores.cs`
- `Controllers/OutlookController.cs`
- `docs/ai/protocols.md`
- `docs/ai/office2016-workstation-contract.md`

## 目標

讓 Outlook Add-in 執行 command 後，可以把成功或失敗摘要回報到 Hub。Hub 只用 in-memory store，不新增 database。

## 建議 DTO

新增 `CommandResultDto`：

```json
{
  "commandId": "...",
  "type": "fetch_rules",
  "success": true,
  "message": "Fetched 3 rules",
  "timestamp": "2026-04-29T10:00:00+08:00"
}
```

## 建議 Endpoint

```http
POST /api/outlook/admin/command-result
GET /api/outlook/admin/command-results
```

## 實作步驟

1. 在 `Models/Dtos.cs` 新增 `CommandResultDto`。
2. 在 `Services/Stores.cs` 新增 in-memory command result store，最多保留 200 筆。
3. 在 `Controllers/OutlookController.cs` 注入並使用 store。
4. 新增 POST/GET endpoint。
5. 更新 `docs/ai/protocols.md`。
6. 更新 `docs/ai/office2016-workstation-contract.md`。
7. 更新 `Plan/STATUS.md`。

## 注意事項

- `message` 不可包含真實 mail body、folder name、rule name、calendar subject。
- 不新增 SignalR event，除非實作時已有明確需求。
- 不修改現有 route 行為。

## 驗證

執行：

```bash
./scripts/build-in-container.sh
```

可以用 curl 測試 POST/GET，但測試內容必須匿名化。

## 更新 STATUS

- `002-hub-command-result-log` 改成 `done`。
- 下一個任務改成 `Plan/003-hub-mail-marker-commands.md`。

## 完成時請回報

- 修改的檔案。
- 新增的 DTO 與 endpoint。
- build 結果。

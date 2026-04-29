# Task 005：新增 Command Result Log

## 目標

讓每次 Outlook Add-in 執行 command 後，都可以回報成功或失敗摘要，方便 Web UI 與 AI 知道操作結果。

## 範圍

只新增輕量 log，不新增 database。

## 建議 DTO

可新增 `CommandResultDto`：

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

## 建議實作步驟

1. 在 `Models/Dtos.cs` 新增 `CommandResultDto`。
2. 在 `Services/Stores.cs` 新增 in-memory list，最多保留 200 筆。
3. 在 `Controllers/OutlookController.cs` 新增 POST/GET admin endpoint。
4. Web UI admin 分頁可以先不做，或只顯示簡單 list。
5. 工作機每次 command 完成後呼叫 POST endpoint。

## 注意事項

- `message` 不可包含真實 mail body。
- 失敗訊息只放錯誤類型與匿名化摘要。

## 驗證

1. 用 curl POST 一筆 command result。
2. 用 GET 取回。
3. 確認超過上限時舊資料會被移除。
4. 執行 `./scripts/build-in-container.sh`。

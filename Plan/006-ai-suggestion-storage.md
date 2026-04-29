# Task 006：Hub 新增 AI Suggestion Storage

## 這個任務的定位

本任務只定義 AI 建議的儲存與查詢 API，不讓 AI 直接執行 Outlook command。

## 開始前必讀

- `AGENTS.md`
- `Plan/STATUS.md`
- `Plan/CONTRACT-INVENTORY.md`
- 本檔
- `Models/Dtos.cs`
- `Services/Stores.cs`
- `Controllers/OutlookController.cs`
- `docs/ai/protocols.md`

## 目標

讓外部 AI 或未來 client 可以把建議送進 Hub，Web UI 可讀取建議。

## 建議 DTO

```json
{
  "id": "...",
  "source": "ai",
  "title": "Move vendor invoice",
  "reason": "The subject contains invoice and the sender is a vendor.",
  "proposedCommandType": "move_mail",
  "proposedPayloadJson": "{\"mailId\":\"...\"}",
  "risk": "low",
  "status": "pending",
  "createdAt": "2026-04-29T10:00:00+08:00"
}
```

## 建議 Endpoint

```http
POST /api/outlook/ai/suggestions
GET /api/outlook/ai/suggestions
```

## 實作步驟

1. 在 `Models/Dtos.cs` 新增 `AiSuggestionDto`。
2. 在 `Services/Stores.cs` 新增 in-memory suggestion store。
3. 在 `Controllers/OutlookController.cs` 新增 POST/GET endpoint。
4. 不新增 AI SDK。
5. 更新 docs。
6. 更新 `Plan/STATUS.md`。

## 注意事項

- `proposedPayloadJson` 不可包含完整 mail body。
- `status` 第一版可用 `pending`。
- 本任務不做 confirm。

## 驗證

執行：

```bash
./scripts/build-in-container.sh
```

用匿名化 JSON 測試 POST/GET。

## 更新 STATUS

- `006-ai-suggestion-storage` 改成 `done`。
- 下一個任務改成 `Plan/007-ai-suggestion-confirmation.md`。

## 完成時請回報

- 新增的 DTO 與 endpoint。
- store 保留上限。
- build 結果。

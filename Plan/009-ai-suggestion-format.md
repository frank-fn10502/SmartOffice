# Task 009：定義 AI Suggestion Format

## 目標

先定義 AI 只能產生建議，不直接執行 Outlook 修改。

## 建議 DTO

```json
{
  "id": "...",
  "source": "ai",
  "title": "Move vendor invoice",
  "reason": "The subject contains invoice and the sender is a vendor.",
  "proposedCommandType": "move_mail",
  "proposedPayload": {
    "mailId": "...",
    "sourceFolderPath": "\\\\Mailbox - User\\Inbox",
    "destinationFolderPath": "\\\\Mailbox - User\\Vendors"
  },
  "risk": "low",
  "createdAt": "2026-04-29T10:00:00+08:00"
}
```

## 建議實作步驟

1. 在 `Models/Dtos.cs` 新增 `AiSuggestionDto`。
2. 在 `Services/Stores.cs` 新增 in-memory suggestion list。
3. 新增 endpoint：
   - `POST /api/outlook/ai/suggestions`
   - `GET /api/outlook/ai/suggestions`
4. Web UI 先只顯示 suggestions，不提供執行按鈕。
5. 確認 suggestion 不包含完整 mail body。

## 注意事項

- AI 建議不能直接變成 Outlook command。
- 必須經過使用者確認。
- `proposedPayload` 如果用 JSON object，請保持 backward-compatible。

## 驗證

1. 用 curl POST 一筆 suggestion。
2. Web UI 或 GET endpoint 可取回。
3. 不啟動 AI 也能測試。
4. 執行 `./scripts/build-in-container.sh`。

# Task 007：AI Suggestion Confirm 後 Enqueue Command

## 這個任務的定位

本任務讓使用者確認 AI suggestion 後，Hub 才把 suggestion 轉成 Outlook command。

## 開始前必讀

- `AGENTS.md`
- `Plan/STATUS.md`
- `Plan/CONTRACT-INVENTORY.md`
- 本檔
- `Models/Dtos.cs`
- `Services/Stores.cs`
- `Controllers/OutlookController.cs`
- `webui/src/App.vue`

## 前置檢查

請確認：

- `006-ai-suggestion-storage` 已完成。
- 至少一個低風險 operation command 已完成，例如 `mark_mail_read` 或 `move_mail`。

如果缺少前置功能，請停止並更新 `Plan/STATUS.md` 的交接注意事項。

## 目標

新增 confirm endpoint，將允許清單內的 AI suggestion 轉成 `PendingCommand`。

## 建議 Endpoint

```http
POST /api/outlook/ai/suggestions/{id}/confirm
```

## 實作步驟

1. 在 suggestion store 加入 status 更新能力。
2. 新增 confirm endpoint。
3. 驗證 `proposedCommandType` 是否在允許清單。
4. 將 suggestion payload 轉成對應 `PendingCommand`。
5. suggestion 標記為 `confirmed`。
6. Web UI 顯示 confirm button 與結果。
7. 更新 `Plan/STATUS.md`。

## 注意事項

- 不允許 unknown command type。
- 不允許 delete、send mail、批次移動。
- 使用者確認前不得 enqueue command。

## 驗證

執行：

```bash
./scripts/build-in-container.sh
```

用匿名化 suggestion 測試 confirm，確認 `/api/outlook/poll` 可取得 command。

## 更新 STATUS

- `007-ai-suggestion-confirmation` 改成 `done`。
- 下一個任務改成 `Plan/008-workstation-fetch-rules.md`。

## 完成時請回報

- confirm endpoint。
- 允許的 command type 清單。
- build 結果。

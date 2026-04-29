# Task 000：初始化 Outlook AI 操作流水線

## 這個任務的定位

這是整套 Plan 的第一個任務。請先執行本任務，再依序執行 `001` 到 `010`。

每次只會有一個 agent 處理一個任務，所以本任務要建立後續 agent 可以接手的共同狀態檔。

## 目標

建立 `Plan/STATUS.md`，記錄整體進度、執行規則與下一個任務。後續每個任務完成後都要更新 `Plan/STATUS.md`。

## 必讀檔案

- `AGENTS.md`
- `docs/ai/project.md`
- `docs/ai/protocols.md`
- `docs/ai/office2016-workstation-contract.md`
- `Models/Dtos.cs`
- `Controllers/OutlookController.cs`

## 請建立檔案

新增：

```text
Plan/STATUS.md
```

內容請使用下列格式：

```markdown
# Outlook AI 操作流水線狀態

## 執行規則

- 每次只處理一個 `Plan/NNN-*.md` 任務。
- 每個新 session 開始時，先讀 `AGENTS.md`、`Plan/STATUS.md`、當次任務檔。
- 完成任務後更新本檔。
- 不要假設模型記得前一個 session 的內容。
- 不要把真實 mail body、folder name、calendar subject、rule name、客戶名稱或公司內部資訊寫入文件、測試資料或 log。

## 任務狀態

| 任務 | 狀態 | 備註 |
| --- | --- | --- |
| 000-bootstrap-plan | done | 已建立狀態檔 |
| 001-contract-inventory | pending |  |
| 002-hub-command-result-log | pending |  |
| 003-hub-mail-marker-commands | pending |  |
| 004-hub-folder-mail-commands | pending |  |
| 005-webui-action-preview | pending |  |
| 006-ai-suggestion-storage | pending |  |
| 007-ai-suggestion-confirmation | pending |  |
| 008-workstation-fetch-rules | pending |  |
| 009-workstation-fetch-calendar | pending |  |
| 010-workstation-mail-metadata | pending |  |

## 已完成變更

- 000：建立本狀態檔。

## 下一個任務

請執行 `Plan/001-contract-inventory.md`。

## 交接注意事項

- Hub 目前使用 `/api/outlook/request-*` enqueue command。
- Outlook Add-in 透過 `/api/outlook/poll` 取得 command。
- Outlook Add-in 用 `/api/outlook/push-*` 回傳結果。
- 預設驗證是 `./scripts/build-in-container.sh`。
```

## 驗證

不需要 build。確認 `Plan/STATUS.md` 已建立且沒有真實敏感資料。

## 完成時請回報

- 已建立 `Plan/STATUS.md`。
- 下一個任務是 `Plan/001-contract-inventory.md`。

# Task 005：Web UI 加入 Outlook 操作 Preview

## 這個任務的定位

本任務只做 Web UI 手動操作 preview。AI suggestion 之後再做。

## 開始前必讀

- `AGENTS.md`
- `Plan/STATUS.md`
- `Plan/CONTRACT-INVENTORY.md`
- 本檔
- `webui/src/App.vue`
- `webui/src/styles.css`
- `Models/Dtos.cs`
- `Controllers/OutlookController.cs`

## 前置檢查

請先確認下列 endpoint 至少部分存在：

- mail marker request endpoint
- folder / move mail request endpoint

如果完全不存在，請停止並回報需要先完成 `003` 或 `004`。

## 目標

在 Web UI 的 Outlook 分頁中，讓使用者可以先看到操作 preview，再手動確認 enqueue command。

## 第一版支援

至少完成其中兩種：

- 標記已讀 / 未讀。
- 設定 category。
- 建立 folder。
- 移動單封郵件。

## 實作步驟

1. 在 `webui/src/App.vue` 新增 action button。
2. 使用 Element Plus dialog 顯示 preview。
3. Dialog 只顯示 subject、folderPath、目標動作，不顯示完整 body。
4. 使用者按確認後才 POST request endpoint。
5. 顯示 queued 狀態，不直接假裝操作成功。
6. 視需要更新 `webui/src/styles.css`。
7. 更新 `Plan/STATUS.md`。

## 注意事項

- 不要在前端直接修改 mail 狀態。
- 等工作機 push 新資料後再更新畫面。
- 不要加入 AI SDK。

## 驗證

執行：

```bash
./scripts/build-in-container.sh
```

啟動 Hub 後，手動確認 dialog 與 enqueue request。

## 更新 STATUS

- `005-webui-action-preview` 改成 `done`。
- 下一個任務改成 `Plan/006-ai-suggestion-storage.md`。

## 完成時請回報

- 完成哪些操作 preview。
- 呼叫哪些 endpoint。
- build 結果。

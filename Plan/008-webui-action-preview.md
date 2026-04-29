# Task 008：Web UI 加入操作 Preview

## 新 Session 起手

本任務可以在全新 session 單獨執行。請先讀：

1. `AGENTS.md`
2. `Plan/000-session-handoff.md`
3. `webui/src/App.vue`
4. `webui/src/styles.css`
5. `Models/Dtos.cs`
6. `Controllers/OutlookController.cs`
7. 本檔

本任務依賴至少一個操作 endpoint 已存在。如果找不到對應 `/api/outlook/request-*` endpoint，請停止並回報缺少前置任務，不要自行設計新 endpoint。

## 目標

在真正執行 Outlook 修改動作前，Web UI 先顯示 preview 與確認按鈕。

## 第一版範圍

只做 UI，不需要 AI。

可支援：

- 設定單封郵件 category。
- 標記已讀/未讀。
- 移動單封郵件。

## 建議實作步驟

1. 在 `webui/src/App.vue` 的 mail item 顯示 action button。
2. 點擊 action button 後開啟 Element Plus dialog。
3. Dialog 顯示：
   - mail subject
   - current folder
   - target action
   - target folder 或 category
4. 使用者按確認後才呼叫 `/api/outlook/request-*` endpoint。
5. 顯示 queued 狀態，不假裝已成功。

## 注意事項

- 不要在前端直接修改 mail list 狀態。
- 等工作機 push 新資料後再更新畫面。
- Dialog 不要顯示完整 mail body。

## 驗證

1. 點 action button。
2. 確認 dialog 資訊正確。
3. 按取消不應 enqueue command。
4. 按確認會 enqueue command。
5. 執行 `./scripts/build-in-container.sh`。

## 完成回報

請回報新增的 UI 入口、dialog 行為、呼叫的 endpoint，以及 build 結果。

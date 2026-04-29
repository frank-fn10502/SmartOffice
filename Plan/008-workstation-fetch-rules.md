# Task 008：工作機實作 Fetch Rules

## 這個任務的定位

本任務在公司電腦的 Outlook Add-in 實作 `fetch_rules`。如果目前 session 只能修改 Hub repo，請不要假裝完成工作機實作，只更新交接文件並回報。

## 開始前必讀

- `AGENTS.md`
- `Plan/STATUS.md`
- `Plan/CONTRACT-INVENTORY.md`
- 本檔
- `docs/ai/protocols.md`
- `docs/ai/office2016-workstation-contract.md`
- `Models/Dtos.cs`

## 目標

讓公司電腦上的 Outlook Add-in 收到 `fetch_rules` command 後，讀取 Outlook rules，並 POST 到 Hub。

## Hub Contract

Add-in poll 會收到：

```json
{
  "id": "...",
  "type": "fetch_rules",
  "mailsRequest": null,
  "calendarRequest": null
}
```

Add-in 完成後呼叫：

```http
POST /api/outlook/push-rules
Content-Type: application/json
```

## 實作步驟

1. 在工作機 Add-in command handler 加入 `fetch_rules` case。
2. 使用 Outlook Object Model 取得目前 session 的 `Rules` collection。
3. 逐一轉成 `OutlookRuleDto` 相容 JSON。
4. 無法完整解析的 condition/action 先用可讀文字表示。
5. POST `/api/outlook/push-rules`。
6. POST command result 或 admin log。
7. 更新 `Plan/STATUS.md`。

## 注意事項

- 第一版只列出 rules，不修改、新增或刪除。
- 不要把真實 rule name 或條件寫進 log。

## 驗證

1. 啟動 Hub。
2. Web UI 開啟 `Outlook` 分頁。
3. 按 `Fetch Rules`。
4. 確認 rules 顯示。

## 更新 STATUS

- `008-workstation-fetch-rules` 改成 `done`，或標記 `blocked` 並說明沒有工作機 repo。
- 下一個任務改成 `Plan/009-workstation-fetch-calendar.md`。

## 完成時請回報

- 工作機修改的檔案。
- Outlook API 使用方式。
- 匿名化測試結果。

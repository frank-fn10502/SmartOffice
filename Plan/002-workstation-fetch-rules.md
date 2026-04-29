# Task 002：工作機實作 Fetch Rules

## 新 Session 起手

本任務可以在全新 session 單獨執行。請先讀：

1. `AGENTS.md`
2. `Plan/000-session-handoff.md`
3. `docs/ai/protocols.md`
4. `docs/ai/office2016-workstation-contract.md`
5. `Models/Dtos.cs`
6. `Controllers/OutlookController.cs`
7. 本檔

不要假設已完成 `001-read-current-contract.md`。如果工作機 Add-in repo 與 Hub repo 分開，請只修改工作機 Add-in repo，並用本檔 contract 對齊 Hub。

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

Payload：

```json
[
  {
    "name": "Move customer mail",
    "enabled": true,
    "executionOrder": 1,
    "ruleType": "receive",
    "conditions": ["sender contains example.com"],
    "actions": ["move to \\\\Mailbox - User\\Customers"],
    "exceptions": []
  }
]
```

## 建議實作步驟

1. 在工作機 Add-in 的 command handler 加入 `fetch_rules` case。
2. 使用 Outlook Object Model 取得目前 session 的 `Rules` collection。
3. 逐一轉成 `OutlookRuleDto` 相容 JSON。
4. 無法完整解析的 condition/action 先用可讀文字表示。
5. POST `/api/outlook/push-rules`。
6. 發生例外時，POST `/api/outlook/admin/log`，不要把真實 rule 內容完整寫入 log。

## 注意事項

- Outlook Rules object model 只支援 Rules and Alerts Wizard 的部分能力。
- 第一版只需要列出與解釋，不要修改、新增或刪除 rules。
- `conditions`、`actions`、`exceptions` 可以先是 best-effort 字串陣列。

## 驗證

1. 啟動 Hub。
2. Web UI 開啟 `Outlook` 分頁。
3. 按 `Fetch Rules`。
4. 確認 rules 顯示在 Web UI。
5. 確認 Hub admin log 沒有 error。

## 完成回報

請回報工作機修改的檔案、Outlook API 使用方式、匿名化測試結果，以及是否有無法解析的 rule condition/action。

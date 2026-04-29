# Task 005：AddIn 實作 Fetch Rules

## 執行位置

本任務請在工作機的完整 SmartOffice / Outlook AddIn solution 中執行。

不要修改 `SmartOffice.Hub` 程式碼。

## 目標

AddIn 收到 `fetch_rules` 後，讀取 Outlook rules，並 push 到 Hub。

## Hub Contract

完成後呼叫：

```http
POST /api/outlook/push-rules
```

Payload：

```json
[
  {
    "name": "Move sample mail",
    "enabled": true,
    "executionOrder": 1,
    "ruleType": "receive",
    "conditions": ["sender contains example.invalid"],
    "actions": ["move to sample folder"],
    "exceptions": []
  }
]
```

## 實作步驟

1. 在 command handler 加入 `fetch_rules`。
2. 使用 Outlook Object Model 取得 `Rules` collection。
3. 逐一轉成 Hub `OutlookRuleDto` JSON。
4. condition/action/exception 第一版可以是 best-effort 可讀字串。
5. POST `/api/outlook/push-rules`。

## 注意事項

- 第一版只列出 rules，不修改、新增或刪除 rules。
- 不要把真實 rule name 或條件寫進 log。

## 驗證

1. Web UI Outlook 分頁按 Fetch Rules。
2. 確認 rules 顯示。
3. 確認 Hub admin log 沒有 error。

## 完成後更新

更新工作機 `Plan/WORKSTATION-STATUS.md`：

- `005-addin-fetch-rules` 改為 done。
- 下一個任務改為 `006-addin-fetch-calendar.md`。

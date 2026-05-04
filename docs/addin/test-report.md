# Office 2016 工作機測試回報

本文件規範工作機實測資料、差異與錯誤要如何回傳。回報不只用於格式不符合；開發 Web UI、mock、Add-in mapping、檔案寫入策略或 protocol 時，只要需要真實 Office 2016 行為作為依據，都可以回傳已匿名化的實測資料。線上文件入口請看 `docs/addin/outlook-references.md`；目前傳送與接收格式請看 `docs/addin/signalr-contract.md`。

## 何時需要回報

當工作機測試發現 mock、DTO、Web UI、檔案輸出或協定需要真實 Office 2016 行為校準時，請回傳一份 markdown。即使沒有錯誤，也可以回報代表性資料形狀、欄位限制、排序、空值、附件或 folder path 等觀察，讓開發機可以調整 DTO、Add-in mapping、mock、Web UI 呈現、檔案寫入或 protocol。

常見回報情境：

- 格式與目前 contract 不符合，或 Office 2016 只能提供不同欄位。
- Web UI 需要真實資料形狀才能調整列表、明細、搜尋、空值或 fallback 呈現。
- mock data 過於理想化，需要補上真實 folder path、body/bodyHtml、sender、timezone、item type 或 attachment 邊界。
- Add-in 需要說明實際如何讀取、排序、過濾或寫入 Office / filesystem。
- 工作機環境限制會影響可用 API、資料量、編碼、檔名、路徑或權限。

不要只描述「壞掉」或「資料長這樣」；請提供足夠資訊讓開發機知道要調整哪一層，以及哪些資料已被匿名化。

建議檔名：

```text
workstation-report-YYYYMMDD-HHMM-command-type.md
```

## 必填內容

回報包必須包含：

- 測試日期、工作機代號、Office application、Office 版本與 bitness。
- Add-in 類型：VSTO / COM / Office.js / mixed。
- Hub commit 或版本、Hub URL、測試 route。
- 收到的 Hub command JSON。
- Add-in 呼叫的 Office API、物件類型與官方文件連結。
- 轉換前的 Office 實測資料結構摘要。
- 實際送出的 Hub JSON。
- Hub response status code 與 response body。
- 預期格式與實際格式的差異；若沒有差異，請寫明這份資料用於校準哪個開發需求。
- 建議修正：改 Hub DTO、改 Add-in mapping、改 mock、改 Web UI、調整檔案寫入策略、或新增 backward-compatible field。
- 已匿名化的最小 sample。

不得包含：

- 真實 mail body。
- 客戶名稱、內部專案名稱、帳號、token。
- 完整 email thread。
- 未遮蔽的 folder name、mail address 或 business data。

## 回報範本

~~~markdown
# 工作機 Office 2016 測試回報

## Summary

- Date: 2026-04-29 09:35 +08:00
- Workstation: WS-REDACTED-01
- Office app: Outlook 2016
- Office version / build: 16.0.xxxxx.xxxxx
- Office bitness: 32-bit
- Add-in type: VSTO
- Hub commit: abc1234
- Hub URL: http://dev-machine:2805
- Scenario: fetch_mails

## Hub Command

```json
{
  "id": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "fetch_mails",
  "mailsRequest": {
    "folderPath": "\\\\Mailbox - User\\Inbox",
    "range": "1d",
    "maxCount": 10
  }
}
```

## Office API Used

- API: `Application.Session.GetDefaultFolder`, `Folder.Items`, `MailItem.HTMLBody`
- NameSpace doc: https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace
- Folder doc: https://learn.microsoft.com/en-us/office/vba/api/outlook.folder
- MailItem doc: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem

## Observed Office Data

Describe the object shape here. Redact sensitive values.

```json
{
  "folderPathObserved": "\\\\Mailbox - User\\Inbox",
  "itemType": "Outlook.MailItem",
  "subject": "[redacted]",
  "receivedTimeKind": "Local",
  "htmlBodyAvailable": false
}
```

## Hub Payload Sent

```json
[
  {
    "subject": "[redacted]",
    "senderName": "Sample Sender",
    "senderEmail": "sender@example.invalid",
    "receivedTime": "2026-04-29T09:30:00+08:00",
    "body": "[redacted]",
    "bodyHtml": "",
    "folderPath": "\\\\Mailbox - User\\Inbox"
  }
]
```

## Hub Response

- Status code: 200
- Body:

```json
{
  "count": 1
}
```

## Difference From Current Contract

`bodyHtml` may be empty for this mailbox mode. Current mock always sends HTML content, so Web UI should tolerate empty `bodyHtml` and use `body`.

## Development Use

Use this sample to update the Web UI mail detail fallback and add one mock mail with empty `bodyHtml`. No route or JSON field rename is required.

## Suggested Fix

- Keep `bodyHtml` backward-compatible and optional in practice.
- Update Outlook mock to include at least one mail with empty `bodyHtml`.
- Document fallback behavior in `docs/ai/protocols.md`.

## Attachments / Extra Notes

- Add-in exception stack trace, if any.
- Screenshot filename, if needed and scrubbed.
- Any relevant Office Trust Center, Exchange, or account-mode constraint.
~~~

## 開發機收到回報後

工作機回報確認後，開發機應依影響範圍更新：

- `docs/ai/protocols.md`：協定行為或 route 語意改變時更新。
- `docs/addin/signalr-contract.md`：工作機傳送或接收格式改變時更新。
- `Models/Dtos.cs`：新增 backward-compatible field 或補註解。
- `Hubs/OutlookAddinHub.cs` 或 `Services/Stores.cs`：SignalR command/result 或 cache 行為需要調整時更新。
- Web UI：需要呈現真實資料邊界、空值、排序、fallback 或 diagnostics 時更新；不要把 Office-specific mapping 塞進前端。

不建議做法：

- 不要因工作機單次測試就 rename JSON field。
- 不要把真實 mail body、folder name 或客戶資訊 commit 到 repo。
- 不要用 mock 的假資料反向定義 Office 2016 真實格式。
- 不要因回報不是錯誤就忽略；代表性實測資料也可以是調整 Web UI、mock 與檔案輸出行為的依據。
- 不要在沒有實測回報的情況下假設 Office.js 最新文件適用於 Office 2016。

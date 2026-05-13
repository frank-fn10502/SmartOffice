# SmartOffice Solution 整合

## 核心事實

本 repository 是：

```text
SmartOffice.Hub
```

它只負責：

- Hub HTTP API。
- SignalR。
- command routing。
- temporary in-memory state。
- Web UI。
- Hub protocol 文件；Outlook AddIn 實作者文件位於 `..\SmartOffice\docs\outlook-addin\`。

完整 SmartOffice / Outlook AddIn solution 會參考：

```text
..\SmartOffice.Hub\SmartOffice.Hub.csproj
```

因此，Outlook AddIn 功能應在 SmartOffice solution / AddIn project 裡實作；Hub contract、mock、HTTP API 與 Web UI 則在 `SmartOffice.Hub` 內維護。涉及真實 Outlook COM/VSTO 行為時，仍需 Windows / Outlook / Office 環境編譯與實測。

## 責任邊界

Hub：

- 提供 `/api/outlook/request-*` endpoint，讓 Web UI、AI 或 MCP client 發出 request。
- 提供 `/hub/outlook-addin`，讓 AddIn 透過 SignalR 接收 command 並回報結果。
- 提供 `/hub/notifications`，讓 client 接收 status、log、progress 與 data update notification；Web UI 的主要資料路徑仍是 HTTP data endpoint。
- 提供 Web UI 與 admin diagnostics。

SmartOffice / Outlook AddIn：

- 連線到 `/hub/outlook-addin`。
- invoke `RegisterOutlookAddin(info)`。
- listen `OutlookCommand`。
- 執行 Outlook COM/VSTO automation。
- 讀取 folders、mails、rules、calendar。
- 修改 mail read state、flag、categories。
- 建立 folder、移動 mail。
- 透過 `BeginFolderSync`、`PushFolderBatch`、`CompleteFolderSync`、`PushMails`、`PushMail`、`PushRules`、`PushCategories`、`PushCalendar`、`SendChatMessage`、`ReportAddinLog` 與 `ReportCommandResult` 將結果回報 Hub。

## 開始前應讀

1. SmartOffice solution 的入口文件。
2. `..\SmartOffice.Hub\AGENTS.md`
3. `..\SmartOffice.Hub\docs\ai\solution-integration.md`
4. `..\SmartOffice.Hub\docs\ai\protocols.md`
5. `..\SmartOffice\docs\outlook-addin\features-checklist.md`
6. `..\SmartOffice\docs\outlook-addin\signalr-contract.md`

## 敏感資料

mail body、folder name、rule name、calendar subject、attendee、客戶名稱與公司內部資訊都可能是敏感 business data。

回報測試結果時，只能使用匿名化 sample。

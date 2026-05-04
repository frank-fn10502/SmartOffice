# 工作機 SmartOffice Solution 關係

## 核心事實

開發機上的本 repository 是：

```text
SmartOffice.Hub
```

它只負責：

- Hub HTTP API。
- SignalR。
- command routing。
- temporary in-memory state。
- Web UI。
- 工作機 contract 文件。

工作機上才有完整的 SmartOffice / Outlook AddIn solution。該 solution 會參考：

```text
..\SmartOffice.Hub\SmartOffice.Hub.csproj
```

因此，工作機 AI 要實作 Outlook AddIn 功能時，應該在工作機 SmartOffice solution / AddIn project 裡修改程式碼，而不是修改 `SmartOffice.Hub`。

## 責任邊界

Hub：

- 提供 `/api/outlook/request-*` endpoint，讓 Web UI、AI 或 MCP client 發出 request。
- 提供 `/hub/outlook-addin`，讓工作機 AddIn 透過 SignalR 接收 command 並回報結果。
- 提供 `/hub/notifications`，讓 Web UI 接收 cached data、status 與 log update。
- 提供 Web UI 與 admin diagnostics。

工作機 SmartOffice / Outlook AddIn：

- 連線到 `/hub/outlook-addin`。
- invoke `RegisterOutlookAddin(info)`。
- listen `OutlookCommand`。
- 執行 Outlook COM/VSTO automation。
- 讀取 folders、mails、rules、calendar。
- 修改 mail read state、flag、categories。
- 建立 folder、移動單封 mail。
- 透過 `BeginFolderSync`、`PushFolderBatch`、`CompleteFolderSync`、`PushMails`、`PushRules`、`PushCategories`、`PushCalendar`、`SendChatMessage`、`ReportAddinLog` 與 `ReportCommandResult` 將結果回報 Hub。

## Plan 使用方式

`Plan/` 內的 markdown 是給工作機 AI 在完整 SmartOffice solution 中逐步執行的任務指引。

除非任務明確說「Hub contract 需要新增或調整」，否則：

- 不要修改 `SmartOffice.Hub` 的 `Controllers/`。
- 不要修改 `SmartOffice.Hub` 的 `Models/`。
- 不要修改 `SmartOffice.Hub` 的 `Services/`。
- 不要把 AddIn 實作寫在 Hub mock 裡冒充完成。

## 工作機 AI 開始前應讀

1. 工作機 SmartOffice solution 的入口文件。
2. `..\SmartOffice.Hub\AGENTS.md`
3. `..\SmartOffice.Hub\docs\ai\workstation-solution.md`
4. `..\SmartOffice.Hub\docs\ai\protocols.md`
5. `..\SmartOffice.Hub\docs\addin\features-checklist.md`
6. `..\SmartOffice.Hub\docs\addin\signalr-contract.md`
7. 當次 `..\SmartOffice.Hub\Plan\NNN-*.md`

## 敏感資料

mail body、folder name、rule name、calendar subject、attendee、客戶名稱與公司內部資訊都可能是敏感 business data。

工作機回報測試結果時，只能使用匿名化 sample。

# Task 000：工作機讀取 Hub Contract

## 執行位置

本任務請在工作機的完整 SmartOffice / Outlook AddIn solution 中執行。

不要修改 `SmartOffice.Hub` 程式碼。`SmartOffice.Hub` 只作為 API contract 參考。

## 目標

讓工作機 AI 先理解 AddIn 與 Hub 的邊界，建立工作機端實作筆記。

## 必讀 Hub 文件

- `..\SmartOffice.Hub\AGENTS.md`
- `..\SmartOffice.Hub\docs\ai\workstation-solution.md`
- `..\SmartOffice.Hub\docs\ai\project.md`
- `..\SmartOffice.Hub\docs\ai\protocols.md`
- `..\SmartOffice.Hub\docs\ai\office2016-workstation-contract.md`
- `..\SmartOffice.Hub\Models\Dtos.cs`
- `..\SmartOffice.Hub\Controllers\OutlookController.cs`

## 工作機端請盤點

在 AddIn solution 中找出：

- poll Hub command 的程式位置。
- push folders 的程式位置。
- push mails 的程式位置。
- HTTP client / Hub URL 設定位置。
- Outlook COM / VSTO automation 入口。
- log 或 diagnostics 的實作位置。

## 請建立工作機文件

在工作機 AddIn repo 建立：

```text
Plan/WORKSTATION-STATUS.md
```

內容至少包含：

- Hub URL。
- AddIn 專案名稱。
- command handler 位置。
- Outlook automation 主要 class。
- 已完成任務列表。
- 下一個任務：`001-addin-poll-command.md`。

## 驗證

不需要 build。只確認文件建立完成，且沒有真實 mail body、folder name、客戶名稱或公司內部資訊。

# AGENTS.md

這份文件是本 repository 的 AI coding agent 主要入口文件。Codex 可直接讀取本檔；GitHub Copilot 的入口請見 `.github/copilot-instructions.md`，內容會指回本檔。

## 必讀規範

- 使用繁體中文與使用者溝通；技術名詞、API name、file path、command、class name 與 JSON field 可保留英文。
- README、AGENTS、CLAUDE、Dockerfile comment、shell script comment、C# XML summary 與 inline code comment 也必須遵守繁體中文規範。
- 本 repository 是 `SmartOffice.Hub`，只包含 Hub/Web UI/contract/mock，不是工作機完整 SmartOffice / Outlook AddIn solution。
- 工作機上的完整 SmartOffice solution 會以 `..\SmartOffice.Hub\SmartOffice.Hub.csproj` 參考本 Hub 專案；真正的 Outlook AddIn / Office automation 實作必須在工作機 SmartOffice solution 中完成。
- 在本 repository 的 `Plan/` 任務是交給工作機 AI 使用的 AddIn 實作指引；除非使用者明確要求修改 Hub contract，否則不要把 `Plan/` 任務解讀成要修改 Hub 程式碼。
- `Plan/status.md` 是 VS Code Copilot custom agent 的任務佇列狀態檔；切分或執行 `Plan/` 任務時也必須遵守 `docs/ai/plan-splitting.md`。
- 修改時維持 SmartOffice.Hub 的邊界：Add-in 負責 Office automation，Hub 負責 HTTP API、SignalR、command routing 與 temporary state，Web UI 負責檢視、手動 request、chat 與 diagnostics。
- Hub 對外 contract 偏好小而明確且 backward-compatible；但 `docs/addin/` 的 AddIn contract 不維護舊版或未使用功能，工作機 AddIn 只實作目前正式 SignalR contract。
- 修改時留意檔案長度與職責邊界；接近或超過約 800 行要評估切分，超過約 1000 行應優先抽出自然模組。詳見 `docs/ai/coding.md` 與 `docs/ai/frontend.md`。
- Office 2016 與受限企業環境是設計約束。除非任務明確需要，避免引入 database、frontend build system、background job framework 或 AI SDK。
- 請假設 mail body、folder name 與 chat message 都可能含有敏感 business data。

## 細節文件

- `docs/ai/project.md`：專案使命、架構邊界、repository layout 與技術選擇。
- `docs/ai/coding.md`：coding rules、Web UI 規範、security notes 與文件期待。
- `docs/ai/frontend.md`：前端框架選擇、限制與導入原則。
- `docs/ai/protocols.md`：Office AddIn SignalR protocol、route 與 SignalR event。
- `docs/ai/workstation-solution.md`：Hub 與工作機 SmartOffice solution 的關係，以及 AddIn 任務應在哪裡實作。
- `docs/ai/plan-splitting.md`：切分 `Plan/` 任務給工作機 AI 或 VS Code Copilot custom agent 的粒度、必要文件與狀態追蹤規範。
- `docs/addin/README.md`：工作機 Outlook AddIn 實作者文件入口。
- `docs/addin/features-checklist.md`：工作機 AI 對照 AddIn command、完成定義與驗收項目的 checklist。
- `docs/addin/outlook-references.md`：Office 2016 Add-in 線上文件入口。
- `docs/addin/signalr-contract.md`：工作機需要傳送與接收的目前格式。
- `docs/addin/test-report.md`：工作機實測資料、差異與錯誤回報格式。
- `docs/ai/validation.md`：本機驗證、Docker Quick Mode、API 與 Web UI 檢查方式。

## 預設驗證

偏好的驗證模式是：

```bash
./scripts/build-in-container.sh
```

如果本機已安裝 .NET SDK，也可以使用：

```bash
dotnet build
```

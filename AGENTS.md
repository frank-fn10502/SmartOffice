# AGENTS.md

這份文件是本 repository 的 AI coding agent 主要入口文件。Codex 可直接讀取本檔；GitHub Copilot 的入口請見 `.github/copilot-instructions.md`，內容會指回本檔。

## 必讀規範

- 使用繁體中文與使用者溝通；技術名詞、API name、file path、command、class name 與 JSON field 可保留英文。
- README、AGENTS、CLAUDE、Dockerfile comment、shell script comment、C# XML summary 與 inline code comment 也必須遵守繁體中文規範。
- 修改時維持 SmartOffice.Hub 的邊界：Add-in 負責 Office automation，Hub 負責 HTTP API、SignalR、command routing 與 temporary state，Web UI 負責檢視、手動 request、chat 與 diagnostics。
- 偏好小而明確、backward-compatible 的 contract；不要隨意 rename JSON field 或破壞既有 route。
- Office 2016 與受限企業環境是設計約束。除非任務明確需要，避免引入 database、frontend build system、background job framework 或 AI SDK。
- 請假設 mail body、folder name 與 chat message 都可能含有敏感 business data。

## 細節文件

- `docs/ai/project.md`：專案使命、架構邊界、repository layout 與技術選擇。
- `docs/ai/coding.md`：coding rules、Web UI 規範、security notes 與文件期待。
- `docs/ai/frontend.md`：前端框架選擇、限制與導入原則。
- `docs/ai/protocols.md`：Office Add-in polling protocol、route 與 SignalR event。
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

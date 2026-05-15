# AGENTS.md

這份文件是本 repository 的 AI coding agent 主要入口文件。Codex 可直接讀取本檔；GitHub Copilot 的入口請見 `.github/copilot-instructions.md`，內容會指回本檔。

## 必讀規範

- 使用繁體中文與使用者溝通；技術名詞、API name、file path、command、class name 與 JSON field 可保留英文。
- README、AGENTS、CLAUDE、Dockerfile comment、shell script comment、C# XML summary 與 inline code comment 也必須遵守繁體中文規範。
- 本 repository 是 `SmartOffice.Hub`，只包含 Hub/Web UI/contract/mock，不是工作機完整 SmartOffice / Outlook AddIn solution。
- `SmartOffice.Hub.Contracts/` 是 Hub-owned contract project；Hub 與工作機 Outlook Add-in 都應引用此 project/package，不要在 Add-in repo 維護 DTO copy。
- Outlook AddIn / Office automation 實作可在本機同步開發；涉及真實 Outlook COM/VSTO 行為時，仍需 Windows / Outlook / Office 環境編譯與實測。
- Outlook AddIn 實作者文件已移到 sibling solution 的 `../SmartOffice/docs/outlook-addin/`。Hub 文件只保留 Hub protocol、route、endpoint 與 AI/MCP 協作資訊；不要在本 repository 重新建立 Outlook AddIn 規格來源。
- 修改時維持 SmartOffice.Hub 的邊界：Add-in 負責 Office automation，Hub 負責 HTTP API、SignalR、command routing 與 temporary state，Web UI 負責檢視、手動 request、chat 與 diagnostics。
- Web UI、HTTP API caller、MCP/SKILL caller 與人類操作者全部都是「使用者」，不是狀態管理器。它們只能提出使用者意圖、查詢條件或操作命令，不能知道或管理 Hub、AddIn、cache、queue、background sync、in-memory state、command lifecycle、資料來源合併策略或 fallback 細節。
- 使用者級功能必須由 Hub 提供單一語意清楚的 API/DTO，例如 profile、通訊錄關聯、mail detail、group members、search result。Web UI 不得自行掃 mail list、拼 recipients、推斷 group/profile、合併 cache、排序狀態或把多個底層 endpoint 編排成業務結果；若 UI 需要這些結果，先補 Hub-owned API/contract/mock，再讓 UI 呈現結果。
- API contract 也必須是使用者視角。不要把內部狀態欄位、cache hit/miss、Hub/AddIn 分工、queue 狀態、背景載入策略或「下一步要 call 哪個內部 endpoint」暴露成 caller 必須理解的流程。若 raw API response 需要 caller 理解內部機制才知道怎麼使用，代表 contract 設計錯了，應回到 Hub 重設語意。
- Web UI 採 feature folder。Outlook domain 程式應集中在 `webui/src/features/outlook/`；不要把 Outlook-specific api、models、utils、components 或 composables 新增回 `webui/src` 根層泛用資料夾。跨 Office Add-in 共用的 workspace shell、navigation 型別與通用 UI contract 放在 `webui/src/features/office/`；未來新增 Word、Excel、PowerPoint 或其他 VSTO AddIn 操作介面時，請建立自己的 `webui/src/features/<domain>/`，並重用 `features/office/` 的 workspace 基礎，不要複製 Outlook-specific shell。
- 新增或修改任何會被 Outlook Add-in 實作的功能時，Hub/mock 是第一階段驗證目標：實作前先查 Microsoft 官方文件確認 Outlook/Office API 概念可行；再更新 Hub DTO、Controller、SignalR Hub、in-memory store、mock backend、Web UI 與 `docs/ai/protocols.md`；用 Mock 環境確認 HTTP API 與 Web UI 行為都符合文件、使用者操作路徑沒有基本 UI 錯誤，並通過 container build 後，才進入 sibling `../SmartOffice/OutlookAddIn` 的 VSTO 真實實作。Web UI 不是唯一 client；若 raw API response 讓其他 caller 難以理解狀態、下一步、資料含義或錯誤原因，必須修正 API contract。
- 本 repository 是 PoC / prototype，不預設保留舊版相容程式碼；contract、DTO、UI state、mock 與文件都應以目前正式行為為準。
- 修改 contract 或流程時，請刪除未使用的舊欄位、舊模式、相容 shim、fallback branch 與死碼；不要留下「可能以後會用」但目前無法驗證或無法處理的相容垃圾。
- 若不確定某段舊行為是否仍有人依賴，最多先詢問使用者是否要完全刪除；除非使用者明確要求 backward compatibility，否則以乾淨刪除為預設。
- 允許主動執行查看型 `git` 指令，例如 `git status`、`git diff`、`git log`、`git show`。不得主動執行會改變 repository 狀態或遠端狀態的 `git` 指令，例如 `git add`、`git commit`、`git checkout`、`git reset`、`git branch`、`git merge`、`git rebase` 或 `git push`；只有在使用者明確授權提交或指定相關 `git` 指令時，才可執行該次必要操作。
- Git commit message 以繁體中文為主，搭配必要英文專有名詞、API name、class name、file path 或 command；不要使用全英文 commit subject，除非使用者明確指定。
- 檔案行數限制是硬性規則，不是風格建議：hand-written source file 不得超過 800 行，包含 `.cs`、`.ts`、`.vue`、`.js`、`.mjs`、`.css` 與 `.sh`。任何程式碼變更完成前都必須讓 `./scripts/check-source-lines.sh` 或 `./scripts/build-in-container.sh` 通過；Web UI 的 `.ts`/`.vue` 也會由 `npm run check:file-lines` 再檢查一次。檔案接近或頂到 800 行時，代表必須停下來檢查職責是否混在一起，並依真實責任邊界切成自然模組；禁止只為了繞過行數限制建立沒有語意的 `partial class` 或機械式 companion file。若檔案超過 800 行，必須先完成職責切分，不能只在回覆中解釋原因。詳見 `docs/ai/coding.md`、`docs/ai/frontend.md` 與 `docs/ai/validation.md`。
- Office 2016 與受限企業環境是設計約束。除非任務明確需要，避免引入 database、frontend build system、background job framework 或 AI SDK。
- 請假設 mail body、folder name 與 chat message 都可能含有敏感 business data。

## 細節文件

- `docs/ai/project.md`：專案使命、架構邊界、repository layout 與技術選擇。
- `docs/ai/coding.md`：coding rules、Web UI 規範、security notes 與文件期待。
- `docs/ai/backend.md`：後端多 Office AddIn feature boundary、service registration、Swagger、SignalR 與 request/fetch-result 擴充原則。
- `docs/ai/frontend.md`：前端框架選擇、限制與導入原則。
- `docs/ai/protocols.md`：Office AddIn SignalR protocol、route 與 SignalR event。
- `docs/ai/solution-integration.md`：Hub 與 SmartOffice solution 的關係，以及 AddIn 任務應在哪裡實作。
- `docs/ai_plugin/README.md`：MCP 與 Agents SKILL 文件入口。
- `docs/ai_plugin/mcp.md`：MCP adapter 與 tool 設計入口。
- `docs/ai_plugin/agents-skill.md`：Agents SKILL 設計入口。
- `docs/ai_plugin/skills/smartoffice-outlook/SKILL.md`：外部 AI 實際讀取的可安裝 SKILL；修改 HTTP API、request/fetch-result workflow、DTO、route、錯誤語意或安全限制時必須同步更新。
- `docs/ai_plugin/skills/smartoffice-outlook/` 是外部 AI 的使用手冊，不是 changelog 或遷移紀錄。內容只能描述當下正式系統如何操作；不要寫「舊的 endpoint」、「legacy」、「保留相容」、「已移除」、「改成」、「目前只/目前需」這類文件更新或歷史脈絡用語。若需要記錄設計演進，放在內部開發文件或 commit/PR 說明，不要放進可安裝 SKILL。
- `../SmartOffice/docs/outlook-addin/README.md`：Outlook AddIn 實作者文件入口。
- `../SmartOffice/docs/outlook-addin/features-checklist.md`：AddIn command、完成定義與驗收項目的 checklist。
- `../SmartOffice/docs/outlook-addin/outlook-references.md`：Office 2016 Add-in 線上文件入口。
- `../SmartOffice/docs/outlook-addin/signalr-contract.md`：工作機需要傳送與接收的目前格式。
- `../SmartOffice/docs/outlook-addin/test-report.md`：工作機實測資料、差異與錯誤回報格式。
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

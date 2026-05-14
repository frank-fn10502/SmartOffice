# Coding Notes for AI Agents

## 語言規範

AI agent 必須使用繁體中文與使用者溝通；技術專有名詞可保留英文，特別是使用英文更清楚或更符合慣例時。

- `SignalR`、`Swagger`、`Docker`、`devcontainer`、`MCP`、`DTO`、`Controller`、`Add-in`、`Office COM/VSTO` 這類技術名詞保留英文。
- 實作說明、風險說明、設計理由與文件正文使用繁體中文。
- 如果有必要使用整段英文，必須立刻補上繁體中文翻譯或摘要。
- Code、API name、file path、command、JSON field、class name、commit-style identifier 保持原文。
- README、AGENTS、CLAUDE、Dockerfile comment、shell script comment、C# XML summary、inline code comment 都必須遵守這條規則。

## Coding Rules

- 本 repository 是 PoC / prototype，不預設保留舊版相容程式碼。修改 contract、DTO、command payload、UI state 或 mock 行為時，以目前正式行為為準。
- AI agent 可主動執行查看型 `git` 指令，例如 `git status`、`git diff`、`git log`、`git show`。不得主動執行會改變 repository 狀態或遠端狀態的 `git` 指令，例如 `git add`、`git commit`、`git checkout`、`git reset`、`git branch`、`git merge`、`git rebase` 或 `git push`；只有在使用者明確授權提交或指定相關 `git` 指令時，才可執行該次必要操作。
- 不要留下未使用的舊欄位、舊模式、相容 shim、fallback branch、dead code 或「可能以後會用」但目前無法驗證的分支。
- 若不確定某個舊 contract 是否仍有人依賴，最多先詢問使用者是否要完全刪除；除非使用者明確要求 backward compatibility，否則以乾淨刪除為預設。
- 可以依任務需要 rename 或移除 JSON field，但必須同步更新 Hub、Web UI、mock、`../SmartOffice/docs/outlook-addin/` contract 與相關文件。
- 不要為了保留舊行為而優先新增 route；只有在使用者明確要求並行新舊流程時才新增相容 route。
- 除非任務明確要求 time handling migration，否則沿用目前的 `DateTime`。
- 註解只用於說明 architectural intent、protocol boundary 或 security-sensitive decision；不要寫只是重述程式碼的註解。
- API 要保持 narrow 且 predictable，方便未來暴露給 MCP。
- Web UI 使用 `Vue 3 + Vite + Element Plus`，但仍要 dependency-light；不要預設加入 Nuxt、Vue Router、Pinia、Axios、Tailwind 或第二套 UI kit。
- Hand-written source file 不得超過 800 行。這是硬性驗收規則，不是提醒或建議；目前 `./scripts/check-source-lines.sh` 會檢查 `.cs`、`.ts`、`.vue`、`.js`、`.mjs`、`.css` 與 `.sh`，`./scripts/build-in-container.sh` 也會執行同一個 gate。Web UI 的 `.ts`/`.vue` 另由 `npm run check:file-lines` 檢查。若檔案超過 800 行，當次任務必須先處理，不能只在回覆中說明「後續再處理」。
- 600 行以上視為預警區：繼續新增功能前要主動尋找自然切分點，例如 helper、controller/composable、component、資料表或 formatter。接近 800 行時，不應再把新功能塞進同一檔案。
- 800 行上限是職責邊界檢查點：頂到上限時，第一步是重新審視該檔案是否混合了不同 domain、transport、state、formatting、UI 或 orchestration 責任，再依實際責任拆出可獨立理解與測試的模組。禁止只為了讓 gate 通過而建立沒有語意的 `partial class`、薄 wrapper 或機械式 companion file；`partial` 只能用在 framework/designer/source generation 等既有機制確實需要的地方，或已有清楚且可說明的語意邊界。
- 切分檔案時不要走向另一個極端：不要為了降低行數拆出大量只有一兩個 trivial function 的檔案；偏好少量、命名清楚、職責完整的模組。優先抽出無副作用的純函式與資料表，其次才拆 UI component、state composable、domain service 或 transport adapter。
- 每次新增大量邏輯或樣式時，必須檢查受影響檔案行數與職責邊界；最終回報需包含 line-count gate 或 container build 的結果。

## Security Notes

請假設 mail body、folder name 與 chat message 都可能含有敏感 business data。

目前 prototype 行為偏寬鬆：

- CORS 接受任意 origin。
- Swagger 永遠啟用。
- 尚未加入 authentication。
- Cached data 只存在 process-local memory。

如果修改 networking、AI、MCP 或 file export 相關行為，請在 change summary 裡說明 privacy 與 security implication。

## 文件期待

新增 Office Add-in surface 時，請同步文件化：

- route prefix 與 command type。
- request/response contract 的 DTO 說明。
- data 是 cached、streamed 還是 persisted。
- UI 或 tool 需要監聽的 SignalR event。

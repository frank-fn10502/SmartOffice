# Coding Notes for AI Agents

## 語言規範

AI agent 必須使用繁體中文與使用者溝通；技術專有名詞可保留英文，特別是使用英文更清楚或更符合慣例時。

- `SignalR`、`Swagger`、`Docker`、`devcontainer`、`MCP`、`DTO`、`Controller`、`Add-in`、`Office COM/VSTO` 這類技術名詞保留英文。
- 實作說明、風險說明、設計理由與文件正文使用繁體中文。
- 如果有必要使用整段英文，必須立刻補上繁體中文翻譯或摘要。
- Code、API name、file path、command、JSON field、class name、commit-style identifier 保持原文。
- README、AGENTS、CLAUDE、Dockerfile comment、shell script comment、C# XML summary、inline code comment 都必須遵守這條規則。

## Coding Rules

- Hub 對外 DTO change 預設保持 backward-compatible；但 `docs/addin/` 的 AddIn contract 不維護舊版或未使用功能，工作機 AddIn 只實作目前正式 SignalR contract。
- 不要隨意 rename JSON field；Office Add-in 與 Web UI 會依賴它們。
- 優先新增 route，避免破壞既有 route。
- 除非任務明確要求 time handling migration，否則沿用目前的 `DateTime`。
- 註解只用於說明 architectural intent、protocol boundary 或 security-sensitive decision；不要寫只是重述程式碼的註解。
- API 要保持 narrow 且 predictable，方便未來暴露給 MCP。
- Web UI 使用 `Vue 3 + Vite + Element Plus`，但仍要 dependency-light；不要預設加入 Nuxt、Vue Router、Pinia、Axios、Tailwind 或第二套 UI kit。
- 避免單一 source file 過長。修改時若發現檔案接近或超過約 800 行，必須評估是否能依自然職責切分；超過約 1000 行時，除非有明確理由，應優先先抽出純 helper、常數表、型別/normalizer、component 或 CSS section。
- 切分檔案時不要走向另一個極端：不要為了降低行數拆出大量只有一兩個 trivial function 的檔案；偏好少量、命名清楚、職責完整的模組。優先抽出無副作用的純函式與資料表，其次才拆 UI component 或 state composable。
- 每次新增大量邏輯或樣式時，順手檢查受影響檔案行數與職責邊界；若暫時不切分，需在回報中說明原因或下一個合理切分點。

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

# Coding Notes for AI Agents

## 語言規範

AI agent 必須使用繁體中文與使用者溝通；技術專有名詞可保留英文，特別是使用英文更清楚或更符合慣例時。

- `SignalR`、`Swagger`、`Docker`、`devcontainer`、`MCP`、`DTO`、`Controller`、`Add-in`、`Office COM/VSTO` 這類技術名詞保留英文。
- 實作說明、風險說明、設計理由與文件正文使用繁體中文。
- 如果有必要使用整段英文，必須立刻補上繁體中文翻譯或摘要。
- Code、API name、file path、command、JSON field、class name、commit-style identifier 保持原文。
- README、AGENTS、CLAUDE、Dockerfile comment、shell script comment、C# XML summary、inline code comment 都必須遵守這條規則。

## Coding Rules

- DTO change 盡量保持 backward-compatible，因為 Add-in 可能落後 Hub 版本。
- 不要隨意 rename JSON field；Office Add-in 與 Web UI 會依賴它們。
- 優先新增 route，避免破壞既有 route。
- 除非任務明確要求 time handling migration，否則沿用目前的 `DateTime`。
- 註解只用於說明 architectural intent、protocol boundary 或 security-sensitive decision；不要寫只是重述程式碼的註解。
- API 要保持 narrow 且 predictable，方便未來暴露給 MCP。
- Web UI 刻意保持 static 且 dependency-light；除非使用者要求，避免加入 npm、bundler 或大型 client framework。

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

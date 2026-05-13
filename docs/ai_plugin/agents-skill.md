# Agents SKILL 設計文件

本文件是 SmartOffice Outlook Agents SKILL 的設計入口。Agents SKILL 適合做輕量版操作手冊、prompt workflow 或 helper script，不一定需要啟動 MCP server。

## 邊界

- Agents SKILL 一律透過本機 Outlook HTTP API 操作 Outlook request，不描述或依賴 API 背後的實作細節。
- 使用者沒有明確指定 folder、全域搜尋或 folder 操作時，Agents SKILL 預設只以主要 mailbox 的 Inbox 為範圍。
- Agents SKILL 回覆查詢或操作結果時必須告知本次 folder 範圍，避免使用者誤會已查全信箱。
- Agents SKILL 可使用 skill folder 底下的 `tmp/<run>/` 暫存大型 JSON response 方便查找；暫存資料可能含敏感 business data，不得提交或長期保留。
- 每次 `request-*` 後都必須用 `requestId` 查詢 `POST /api/outlook/fetch-result-*`，不要只依賴 HTTP 200。
- 修改郵件前必須先取得 `MailItemDto.id`，不可只用 subject 或 folder name 猜測目標。
- Agents SKILL 應提醒 agent 將 mail body、folder name 與 chat message 視為敏感 business data。

## 維護規則

- 外部 AI 主要讀取可安裝的 skill folder，不會自動讀取 repository 內的 `AGENTS.md`、Hub protocol notes 或 AddIn 實作者文件。
- 修改 Hub HTTP API、request/fetch-result workflow、DTO 欄位、route、錯誤語意、安全限制、folder/path 規則或建議操作順序時，必須同步更新 `docs/ai_plugin/skills/smartoffice-outlook/SKILL.md` 與其 `references/`。
- 若新增 API workflow 或修正 API 可理解性問題，也要更新 `references/http-api.md`、`references/workflows.md`，必要時補 `acceptance-scenarios.md`，讓外部 AI 可以照新 contract 驗證。
- 不要只更新內部文件後假設外部 AI 會知道；SKILL folder 是外部 AI 的唯一操作手冊。
- SKILL folder 是當下系統的使用手冊，不是版本差異、遷移紀錄或開發備忘錄。`SKILL.md`、`references/http-api.md`、`references/workflows.md` 與 installer help 只能描述目前正式操作方式；禁止使用「舊的 endpoint」、「legacy」、「保留相容」、「已移除」、「不再」、「改成」、「目前只/目前需」等歷史脈絡或文件更新語氣。需要保留演進背景時，寫在內部開發文件、issue、commit 或 PR 說明。

## 共用流程

完整呼叫順序、HTTP workflow 與資料安全限制請參考：

- `docs/ai_plugin/mcp-agents-skill-integration.md`
- `docs/ai_plugin/acceptance-scenarios.md`：開發時使用的使用者任務模擬與驗收清單；不會安裝進 Agents SKILL。

## Skill 資料夾

目前可安裝的 Agents SKILL 集中於：

- `docs/ai_plugin/skills/smartoffice-outlook/`

此資料夾包含 `SKILL.md`、`agents/openai.yaml`、`references/` 與安裝 script。安裝 script 只安裝 skill 資料夾，不會產生或修改 `AGENTS.md`、`CODEX.md`、`.github/copilot-instructions.md` 等專案規則檔。

## 目前支援的 AI 工具

不同 AI 工具讀取 SKILL folder 的位置不同；installer 只複製 `smartoffice-outlook` skill folder，不會產生或修改 `AGENTS.md`、`CODEX.md`、`.github/copilot-instructions.md`、`*.instructions.md` 等規則檔。

| 工具 | 支援狀態 | 安裝 / 讀取位置 | 備註 |
| --- | --- | --- | --- |
| Codex / OpenAI Agents skills | 支援 | user: `${CODEX_HOME:-$HOME/.codex}/skills/smartoffice-outlook`；project: `<project>/.codex/skills/smartoffice-outlook` | 會讀取 `SKILL.md`、`agents/openai.yaml` 與 `references/`。 |
| GitHub Copilot Agent Skills | 支援 | user: `$HOME/.copilot/skills/smartoffice-outlook`；project: `<project>/.github/skills/smartoffice-outlook` | 只複製 SKILL folder；不寫入 Copilot custom instructions。 |
| opencode Agent Skills | 支援 | user: `${XDG_CONFIG_HOME:-$HOME/.config}/opencode/skills/smartoffice-outlook`；project: `<project>/.opencode/skills/smartoffice-outlook` | 只複製 SKILL folder；不寫入 `AGENTS.md` 或 `opencode.json`。 |
| 其他 AI 工具 | 未自動支援 | 依各工具規範 | 提供可攜的 Markdown reference 與 helper script。 |

因此，`--project` / `-Project` 會把同一份 skill folder 複製到指定 project 內各工具自己的 skill 位置。

可從 repository root 安裝：

```bash
./install-smartoffice-outlook-skill.sh
```

Windows / PowerShell 環境可用：

```powershell
pwsh -File .\install-smartoffice-outlook-skill.ps1
```

預設會同時安裝到 Codex、Copilot 與 opencode 的 user skill folder；若要安裝到特定專案：

```bash
./install-smartoffice-outlook-skill.sh --project /path/to/project
```

PowerShell：

```powershell
pwsh -File .\install-smartoffice-outlook-skill.ps1 -Project C:\path\to\project
```

root script 只是薄 wrapper；實際安裝邏輯仍在 `docs/ai_plugin/skills/smartoffice-outlook/scripts/`，方便 skill folder 保持可攜。

若只想安裝其中一種工具：

```bash
./install-smartoffice-outlook-skill.sh --tools codex,opencode
./install-smartoffice-outlook-skill.sh --tool copilot --project /path/to/project
```

PowerShell：

```powershell
pwsh -File .\install-smartoffice-outlook-skill.ps1 -Tools codex,opencode
pwsh -File .\install-smartoffice-outlook-skill.ps1 -Tool copilot -Project C:\path\to\project
```

安裝 script 預設採全新重裝：若目標 `smartoffice-outlook` skill folder 已存在，會先移除整個目標 folder 再複製目前版本，避免刪除或更名檔案後留下舊檔污染新版 skill。

為避免覆蓋其他人建立的同名 skill，skill folder 內有 `.smartoffice-skill-id` marker，`SKILL.md` frontmatter 也包含相同 `metadata.skill_id`。安裝 script 只有在既有目標 marker 符合時才會全新重裝；若目標同名資料夾沒有 marker 或 id 不符，會停止。

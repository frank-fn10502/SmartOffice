# 新 Session 任務交接規則

每個 `Plan/*.md` 都必須可以在全新的 AI session 中單獨執行。執行任務前，不要假設模型看過其他對話或記得上一個任務。

## 每次開始都要做

1. 先讀 `AGENTS.md`，遵守繁體中文溝通與 SmartOffice.Hub 邊界。
2. 讀本檔。
3. 讀當次任務 markdown。
4. 讀任務列出的必讀檔案。
5. 用 `rg` 搜尋現有 DTO、route、command type 與 Web UI 呼叫點。
6. 只實作當次任務，不順手做下一個任務。

## Repository 邊界

- Add-in 負責 Office automation。
- Hub 負責 HTTP API、SignalR、command routing 與 temporary state。
- Web UI 負責檢視、手動 request、chat 與 diagnostics。
- 不引入 database、background job framework、frontend build system 或 AI SDK。
- 不 rename 既有 JSON field、route 或 command type，除非任務明確要求。

## 敏感資料規則

- 不要把真實 mail body、folder name、calendar subject、rule name、客戶名稱或公司內部資訊寫進文件、測試資料或 log。
- 測試資料一律用 `Sample`、`Mock`、`example.invalid`、`[redacted]`。
- 錯誤 log 只放匿名化摘要。

## 預設驗證

優先執行：

```bash
./scripts/build-in-container.sh
```

如果工作機沒有 Docker 或無法執行，才改用：

```bash
dotnet build
```

Web UI 任務也要透過 container build，因為它會跑 `vue-tsc` 與 `vite build`。

## 完成回報格式

每個 session 結束時，請回報：

- 完成的任務檔名。
- 修改的檔案。
- 新增或變更的 route、DTO、command type、SignalR event。
- 驗證 command 與結果。
- 未完成事項或需要下一個 session 接手的內容。

## 不要做的事

- 不要實作批次刪除、直接寄信、批次移動大量郵件。
- 不要在沒有 preview/confirm 的情況下讓 AI 直接修改 Outlook 狀態。
- 不要把 mock 寫在 Web UI；mock Add-in 必須走 Hub protocol。

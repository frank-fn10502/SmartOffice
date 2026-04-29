# Task 001：盤點目前 Hub Contract

## 這個任務的定位

這是線性任務的第 1 步。請只做 contract 盤點與文件更新，不實作新功能。

## 開始前必讀

- `AGENTS.md`
- `Plan/STATUS.md`
- 本檔
- `docs/ai/protocols.md`
- `docs/ai/office2016-workstation-contract.md`
- `Models/Dtos.cs`
- `Controllers/OutlookController.cs`
- `Services/Stores.cs`
- `webui/src/App.vue`

## 目標

建立一份目前 Hub contract 的索引，讓後續任務不用重新猜測 route、DTO、command type 與 SignalR event。

## 請建立檔案

新增：

```text
Plan/CONTRACT-INVENTORY.md
```

## 內容要求

請在 `Plan/CONTRACT-INVENTORY.md` 記錄：

1. Web UI / AI 可呼叫的 request endpoint。
2. Outlook Add-in 可呼叫的 poll 與 push endpoint。
3. 目前 `PendingCommand` 支援的欄位。
4. 目前 DTO 清單與用途。
5. 目前 SignalR event 清單。
6. 哪些資料可能含敏感 business data。
7. 目前已存在但尚未有工作機實作的 command。

## 請更新檔案

更新 `Plan/STATUS.md`：

- 將 `001-contract-inventory` 狀態改成 `done`。
- 在「已完成變更」加入本任務摘要。
- 將「下一個任務」改成 `Plan/002-hub-command-result-log.md`。

## 驗證

不需要 build。請確認 `Plan/CONTRACT-INVENTORY.md` 沒有真實 mail body、folder name、rule name 或 calendar subject。

## 完成時請回報

- 新增的檔案。
- 盤點到的 endpoint 數量。
- 下一個任務。

# Task 009：工作機實作 Fetch Calendar

## 這個任務的定位

本任務在公司電腦的 Outlook Add-in 實作 `fetch_calendar`。如果目前 session 只能修改 Hub repo，請不要假裝完成工作機實作，只更新交接文件並回報。

## 開始前必讀

- `AGENTS.md`
- `Plan/STATUS.md`
- `Plan/CONTRACT-INVENTORY.md`
- 本檔
- `docs/ai/protocols.md`
- `docs/ai/office2016-workstation-contract.md`
- `Models/Dtos.cs`

## 目標

讓 Outlook Add-in 收到 `fetch_calendar` command 後，讀取指定天數內的行事曆事件，並 POST 到 Hub。

## Hub Contract

Add-in poll 會收到：

```json
{
  "id": "...",
  "type": "fetch_calendar",
  "calendarRequest": {
    "daysForward": 14
  }
}
```

Add-in 完成後呼叫：

```http
POST /api/outlook/push-calendar
Content-Type: application/json
```

## 實作步驟

1. 在工作機 Add-in command handler 加入 `fetch_calendar` case。
2. 取得 Outlook default Calendar folder。
3. 查 `DateTime.Now` 到 `DateTime.Now.AddDays(daysForward)`。
4. 將 `AppointmentItem` 轉成 `CalendarEventDto` 相容 JSON。
5. recurring meeting 第一版只標記 `isRecurring`。
6. POST `/api/outlook/push-calendar`。
7. POST command result 或 admin log。
8. 更新 `Plan/STATUS.md`。

## 注意事項

- `subject`、`location`、`attendees` 都可能含敏感資料。
- 不要全量掃描多年行事曆。
- 第一版不要建立或修改會議。

## 驗證

1. Web UI 開啟 `Outlook` 分頁。
2. 按 `Fetch Calendar`。
3. 確認近期會議顯示。

## 更新 STATUS

- `009-workstation-fetch-calendar` 改成 `done`，或標記 `blocked` 並說明沒有工作機 repo。
- 下一個任務改成 `Plan/010-workstation-mail-metadata.md`。

## 完成時請回報

- 工作機修改的檔案。
- 查詢日期範圍。
- 匿名化測試結果。

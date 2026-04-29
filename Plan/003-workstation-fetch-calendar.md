# Task 003：工作機實作 Fetch Calendar

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

Payload：

```json
[
  {
    "id": "...",
    "subject": "Meeting subject",
    "start": "2026-04-30T10:00:00+08:00",
    "end": "2026-04-30T11:00:00+08:00",
    "location": "Meeting Room",
    "organizer": "Organizer Name",
    "requiredAttendees": "User A; User B",
    "isRecurring": false,
    "busyStatus": "busy"
  }
]
```

## 建議實作步驟

1. 在工作機 Add-in 的 command handler 加入 `fetch_calendar` case。
2. 取得 Outlook default Calendar folder。
3. 只查 `DateTime.Now` 到 `DateTime.Now.AddDays(daysForward)`。
4. 將 `AppointmentItem` 轉成 `CalendarEventDto` 相容 JSON。
5. recurring meeting 第一版只標記 `isRecurring`，不需要展開所有例外。
6. POST `/api/outlook/push-calendar`。

## 注意事項

- `subject`、`location`、`attendees` 都可能含敏感資料。
- 不要全量掃描多年行事曆。
- 第一版不要建立或修改會議。

## 驗證

1. Web UI 開啟 `Outlook` 分頁。
2. 按 `Fetch Calendar`。
3. 確認近期會議顯示。
4. 用匿名化資料回報是否遇到 recurring meeting 或權限問題。

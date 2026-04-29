# Task 006：AddIn 實作 Fetch Calendar

## 執行位置

本任務請在工作機的完整 SmartOffice / Outlook AddIn solution 中執行。

不要修改 `SmartOffice.Hub` 程式碼。

## 目標

AddIn 收到 `fetch_calendar` 後，讀取近期 Outlook 行事曆事件，並 push 到 Hub。

## Hub Contract

Poll 會收到：

```json
{
  "type": "fetch_calendar",
  "calendarRequest": {
    "daysForward": 14
  }
}
```

完成後呼叫：

```http
POST /api/outlook/push-calendar
```

## Payload Shape

```json
[
  {
    "id": "...",
    "subject": "Sample meeting",
    "start": "2026-04-30T10:00:00+08:00",
    "end": "2026-04-30T11:00:00+08:00",
    "location": "Sample room",
    "organizer": "Sample organizer",
    "requiredAttendees": "Sample attendee",
    "isRecurring": false,
    "busyStatus": "busy"
  }
]
```

## 實作步驟

1. 在 command handler 加入 `fetch_calendar`。
2. 取得 Outlook default Calendar folder。
3. 查 `DateTime.Now` 到 `DateTime.Now.AddDays(daysForward)`。
4. 將 `AppointmentItem` 轉成 Hub `CalendarEventDto` JSON。
5. recurring meeting 第一版只標記 `isRecurring`。
6. POST `/api/outlook/push-calendar`。

## 注意事項

- 不要全量掃描多年行事曆。
- subject、location、attendees 都可能含敏感資料，不可寫入 log。
- 第一版不要建立或修改會議。

## 驗證

1. Web UI Outlook 分頁按 Fetch Calendar。
2. 確認近期會議顯示。

## 完成後更新

更新工作機 `Plan/WORKSTATION-STATUS.md`：

- `006-addin-fetch-calendar` 改為 done。
- 下一個任務改為 `007-addin-mark-mail.md`。

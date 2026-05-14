# Organizing API Reference

讀這份文件處理 address book、calendar、rules、categories、mail/folder mutation、chat 與常用 DTO。共通 request/fetch-result envelope 見 `http-api.md`。

## 目錄

- `Address Book`: 輕量 lookup 與通訊錄同步。
- `Calendar / Rules / Categories`: calendar、rule、category endpoint 與欄位。
- `Mail / Folder Mutations`: update、move、delete 與 create folder。
- `Chat`: chat endpoint。
- `CalendarEventDto`: calendar event 欄位。

## Address Book

通訊錄是 SmartOffice 的關聯視圖。資料來源包含已讀取 mails 的 sender / to / cc / bcc / group members、已讀取 calendar events 的 organizer / attendees，以及 `request-address-book` 從 Outlook Contacts folder / AddressLists / GAL 同步回來的 metadata。它只暴露 metadata、mail ids 與少量 subject sample，不會讀取或回傳完整 mail body。

使用方式：

- 想檢查一個收件者是否和使用者有已知互動：呼叫 `GET /api/outlook/address-book/lookup?email={email}`。
- `state=known` 代表 SmartOffice 找到 mail 或 calendar 關聯；`state=unknown` 只代表目前未知，不代表 Outlook 裡一定沒有。
- 若使用者要求同步真正 Outlook 通訊錄，呼叫 `POST /api/outlook/request-address-book`，再用 `POST /api/outlook/fetch-result-address-book` 讀 `data.contacts`。
- 需要檢查收件者是否能用 group 合併時，呼叫 `POST /api/outlook/address-book/merge-suggestions`，body 為 `{ "recipients": ["frank@example.test", "group1@example.test"] }`。
- `contact.relationKinds` 會指出 `sender`、`to`、`cc`、`bcc`、`organizer`、`attendee` 或 `group_member` 等關聯。
- `contact.isGroup=true` 代表該 entry 是 distribution list 或 group；用 `memberCount`、`memberSmtpAddresses`、`memberGroupSmtpAddresses` 摘要成員與子群組。個人或 group 的 `memberOfGroupSmtpAddresses` 表示它被哪些 group 包含。
- `contact.isLikelySelf=true` 代表該地址看起來是自己的寄件地址；Hub 主要從 Sent folder 的 sender 推斷。

`request-address-book` body：

```json
{
  "includeOutlookContacts": true,
  "includeAddressLists": true,
  "maxContacts": 1000,
  "maxAddressEntriesPerList": 500,
  "maxGroupMembers": 50,
  "maxGroupDepth": 1
}
```

`maxContacts`、`maxAddressEntriesPerList`、`maxGroupMembers` 與 `maxGroupDepth` 是負載上限；不要要求無限制 GAL 枚舉或無限制展開 nested group。`maxGroupMembers=0` 表示只讀 group metadata，不展開成員。讀取結果一律透過 `fetch-result-address-book` 用 `cursor` / `take` 分頁取得。

## Calendar / Rules / Categories

- `POST /api/outlook/request-calendar` with `{ "daysForward": 31, "startDate": null, "endDate": null }` -> `POST /api/outlook/fetch-result-calendar`
- `POST /api/outlook/request-calendar-rooms` -> `POST /api/outlook/fetch-result-calendar-rooms`
- `POST /api/outlook/request-create-calendar-event` with calendar event fields, `requiredAttendees[]`, and optional `resources[]` -> `POST /api/outlook/fetch-result-create-calendar-event`
- `POST /api/outlook/request-update-calendar-event` with `eventId`, `smartOfficeEventId`, calendar event fields, `requiredAttendees[]`, and optional `resources[]` -> `POST /api/outlook/fetch-result-update-calendar-event`
- `POST /api/outlook/request-delete-calendar-event` with `eventId` and `smartOfficeEventId` -> `POST /api/outlook/fetch-result-delete-calendar-event`
- `POST /api/outlook/request-rules` -> `POST /api/outlook/fetch-result-rules`
- `POST /api/outlook/request-manage-rule` with `OutlookRuleCommandRequest` -> `POST /api/outlook/fetch-result-manage-rule`
- `POST /api/outlook/request-categories` -> `POST /api/outlook/fetch-result-categories`
- `POST /api/outlook/request-upsert-category` with category object -> `POST /api/outlook/fetch-result-upsert-category`

`OutlookRuleCommandRequest` 常用欄位：

```json
{
  "operation": "upsert",
  "storeId": "",
  "ruleName": "客戶郵件標記",
  "originalRuleName": "",
  "originalExecutionOrder": null,
  "ruleType": "receive",
  "enabled": true,
  "executionOrder": null,
  "conditions": {
    "subjectContains": ["報價"],
    "bodyContains": [],
    "bodyOrSubjectContains": [],
    "messageHeaderContains": [],
    "senderAddressContains": ["example.com"],
    "recipientAddressContains": [],
    "categories": ["客戶"],
    "hasAttachment": true,
    "importance": "high",
    "toMe": false,
    "toOrCcMe": false,
    "onlyToMe": false,
    "meetingInviteOrUpdate": false
  },
  "actions": {
    "moveToFolderPath": "\\\\主要信箱 - User\\Inbox\\客戶",
    "copyToFolderPath": "",
    "assignCategories": ["客戶"],
    "clearCategories": false,
    "markAsTask": true,
    "markAsTaskInterval": "this_week",
    "delete": false,
    "desktopAlert": true,
    "stopProcessingMoreRules": true
  }
}
```

`importance` 可用 `any`、`low`、`normal`、`high`；`markAsTaskInterval` 可用 `today`、`tomorrow`、`this_week`、`next_week`、`no_date`。`hasAttachment=false` 不支援，因 Outlook Rules object model 只能建立「有附件」條件。

`CategoryCommandRequest`：

```json
{
  "name": "Project",
  "color": "olCategoryColorGreen",
  "colorValue": 5,
  "shortcutKey": ""
}
```

常用 Outlook category color：

| 使用者顏色 | `color` | `colorValue` |
| --- | --- | --- |
| 無色 | `olCategoryColorNone` | `0` |
| 紅色 | `olCategoryColorRed` | `1` |
| 橘色 | `olCategoryColorOrange` | `2` |
| 桃色 | `olCategoryColorPeach` | `3` |
| 黃色 | `olCategoryColorYellow` | `4` |
| 綠色 | `olCategoryColorGreen` | `5` |
| 青色 | `olCategoryColorTeal` | `6` |
| 橄欖 | `olCategoryColorOlive` | `7` |
| 藍色 | `olCategoryColorBlue` | `8` |
| 紫色 | `olCategoryColorPurple` | `9` |
| 栗色 | `olCategoryColorMaroon` | `10` |
| 鋼藍 | `olCategoryColorSteel` | `11` |
| 深鋼藍 | `olCategoryColorDarkSteel` | `12` |
| 灰色 | `olCategoryColorGray` | `13` |
| 深灰 | `olCategoryColorDarkGray` | `14` |
| 黑色 | `olCategoryColorBlack` | `15` |

Agent 處理 category 時應先用 `request-categories` / `fetch-result-categories` 讀取 master category list。Category name 比對大小寫不敏感；若已有同名 category，`request-upsert-category` 視為更新該 category。使用者指定顏色不在上表時，請使用者改選，不要猜 `OlCategoryColor` enum 或 numeric value。套用 category 到 mail 前，必須先從 mail fetch result 定位唯一 `mailId` 與同筆 mail 的 `folderPath`。

## Mail / Folder Mutations

修改、移動或刪除 mail 前，必須使用 `fetch-result-* data.mails` 中的 `id` 與 `folderPath`，並把該 `id` 作為 mutation request 的 `mailId`，確認目標就是使用者指定的 mail。若候選 mail 不唯一，先請使用者確認，不要任選。

### `POST /api/outlook/request-update-mail-properties`

Request:

```json
{
  "mailId": "mail-id",
  "folderPath": "/主要信箱 - User/收件匣",
  "isRead": true,
  "flagInterval": "today",
  "flagRequest": "今天",
  "taskStartDate": null,
  "taskDueDate": null,
  "taskCompletedDate": null,
  "categories": ["Customer"],
  "newCategories": []
}
```

`flagInterval`: `none`、`today`、`tomorrow`、`this_week`、`next_week`、`no_date`、`custom`、`complete`。

### `POST /api/outlook/request-create-folder`

```json
{
  "parentFolderPath": "/主要信箱 - User/Projects",
  "name": "Sample Folder"
}
```

### `POST /api/outlook/request-delete-folder`

```json
{
  "folderPath": "/主要信箱 - User/Projects/Sample Folder"
}
```

語意是將 folder 移到 Outlook default Deleted Items folder，不是永久刪除。HTTP request 不接受目的 folder；SmartOffice 必須用 Outlook default folder identity 定位 Deleted Items，不得依賴 `Deleted Items`、`刪除的郵件` 或其他本地化顯示名稱。完成後告知使用者 folder 已移到刪除資料夾。若目標 folder 已經位於 default Deleted Items folder 或其子層，paired fetch result 會回 `state=failed` / `message=manual_delete_required`；agent 必須停止，並請使用者自行到 Outlook 操作。

### `POST /api/outlook/request-move-mail`

```json
{
  "mailId": "mail-id",
  "sourceFolderPath": "/主要信箱 - User/收件匣",
  "destinationFolderPath": "/主要信箱 - User/Projects"
}
```

### `POST /api/outlook/request-move-mails`

單次最多 500 個 `mailIds`；更多郵件必須由 caller 分批呼叫。

```json
{
  "mailIds": ["mail-id-1", "mail-id-2"],
  "sourceFolderPath": "/主要信箱 - User/收件匣",
  "sourceFolderPaths": ["/主要信箱 - User/收件匣"],
  "destinationFolderPath": "/主要信箱 - User/Projects",
  "continueOnError": true
}
```

超過限制時回 `400 too_many_mail_ids`，並提供 `maxBatchSize` 與 `actualCount`。

### `POST /api/outlook/request-delete-mail`

```json
{
  "mailId": "mail-id",
  "folderPath": "/主要信箱 - User/收件匣"
}
```

語意是移到 Outlook default Deleted Items folder，不是永久刪除。HTTP request 不接受 `destinationFolderPath`；完成後告知使用者 mail 已移到刪除資料夾；若使用者要永久刪除，請使用者自行到 Outlook 操作。

## Chat

- `POST /api/outlook/chat`
- `GET /api/outlook/chat`

Request:

```json
{
  "source": "web",
  "text": "message"
}
```

Chat text 可能含敏感 business data。

## `CalendarEventDto`

`id`, `subject`, `start`, `end`, `location`, `organizer`, `requiredAttendees`, `isRecurring`, `busyStatus`, `smartOfficeOwned`, `smartOfficeEventId`。`organizer` 是 `OutlookRecipientDto`，`requiredAttendees` 是 `OutlookRecipientDto[]`。

Calendar update / delete 只適用 `smartOfficeOwned=true` 且 `smartOfficeEventId` 相符的 event。若 API 回 `not_smartoffice_owned`，停止操作並告知使用者 SmartOffice 只能更新或刪除 SmartOffice 建立的 calendar event。

建立或更新 calendar event 時，出席者放在 `requiredAttendees[]`，會議室或設備放在 `resources[]`。先用 `request-calendar-rooms` 取得 `data.rooms`，讓使用者從下拉選單選會議室或設備 resource。送出 create/update 時把選取項目放進 `resources[]`；SmartOffice 會交由 Outlook 解析。UI 上的「會議室 / Resource」不是地點文字，也不是一般出席者，而是 Outlook resource recipient。

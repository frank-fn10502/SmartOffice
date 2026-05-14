# Organizing API Reference

讀這份文件處理 calendar、rules、categories、mail/folder mutation、chat 與常用 DTO。共通 request/fetch-result envelope 見 `http-api.md`。

## 目錄

- `Calendar / Rules / Categories`: calendar、rule、category endpoint 與欄位。
- `Mail / Folder Mutations`: update、move、delete 與 create folder。
- `Chat`: chat endpoint。
- `CalendarEventDto`: calendar event 欄位。
- `Address Book / Group Relations`: 通訊錄、group members、個人與 group 關聯。

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

## Address Book / Group Relations

AI、MCP client 或人類用 HTTP API 完成一般任務時，優先使用高階查詢：

- `POST /api/outlook/request-address-book-relation`
- `POST /api/outlook/fetch-result-address-book-relation`

Web UI 或需要瀏覽通訊錄目錄的 caller 才使用瀏覽型 endpoint：

- `POST /api/outlook/request-address-book-roots`
- `POST /api/outlook/fetch-result-address-book-roots`
- `POST /api/outlook/request-address-list-entries`
- `POST /api/outlook/fetch-result-address-list-entries`
- `POST /api/outlook/request-address-book-group-members`
- `POST /api/outlook/fetch-result-address-book-group-members`

### 反查 group 或個人關聯

查 group：

```json
{
  "targetKind": "group",
  "groupSmtpAddress": "group@example.com",
  "groupId": "",
  "take": 50
}
```

查個人：

```json
{
  "targetKind": "person",
  "email": "person@example.com",
  "smtpAddress": "",
  "displayName": "",
  "id": "",
  "take": 50
}
```

`targetKind` 可用 `group`、`person` 或空字串；空字串代表自動判斷。Fetch result 的 `data` 常用欄位：

- `state`: `found`、`not_found`、`ambiguous`。
- `target`: 找到的 group 或個人。
- `matches`: 多筆候選；若 `state=ambiguous`，請使用者縮小目標，不要任選。
- `members`: group 的 direct members。
- `memberGroups`: group 直接包含的 nested groups。
- `memberOfGroups`: 個人所屬 groups，或目標 entry 目前已知的上層 groups。
- `containingGroups`: 直接包含目標 entry 的 groups。
- `isLikelySelf`, `isRelatedToSelf`: 可直接判讀的關聯狀態。
- `recipientRelevance`: 只根據收件者路徑計算的參考關聯度。

判讀規則：

- `state=found` 才能回覆確定關聯。
- `state=ambiguous` 時列出必要候選，請使用者確認目標。
- `state=not_found` 代表本次查詢找不到目標。
- group 的 `isRelatedToSelf=true` 代表該 group 與自己有關。
- `members[]` 裡有 `isLikelySelf=true` 的 entry 時，可回覆自己在該 group 中。
- 個人的 `memberOfGroups[]` 表示該個人屬於哪些 groups。
- 若 `state=running`，稍後用同一個 `requestId` 再呼叫 paired fetch result；不要判定為沒有關聯。

### Group 收件者路徑關聯度

當使用者問「這封 mail 是寄給 groupD1，跟我有多相關」時，用 `request-address-book-relation` 查該 group。Fetch result 的 `data.recipientRelevance` 會提供：

- `score`: 0-100，分數越高代表收件者路徑越直接；本人直接收件為 100，往上一到兩層 group 仍維持高分，第三層後用指數衰減快速下降。
- `level`: `direct`、`strong`、`broad`、`weak`、`unknown`。
- `summary`: 可直接摘要給使用者的解釋。
- `routeDepth`: 收件者路徑深度；`0` 代表本人，`1` 代表 direct group，`2` 代表 direct nested group，`3` 代表更間接的已知路徑，`-1` 代表沒有已知路徑。
- `directPersonCount`: 該 group 的 direct person 數量。
- `directGroupCount`: 該 group 的 direct nested group 數量。
- `audienceSize`: direct person + group 的估計受眾大小。
- `includesSelf`: group 是否與自己有關。
- `includesSelfDirectly`: 自己是否為 direct member。
- `reasons[]`: 分數依據，例如 direct membership、受眾大小、過往 mail/calendar 互動。

這個分數只衡量「收件者路徑與受眾範圍」對自己的關聯度，是其中一個排序參考，不判斷 mail 內容本身重要與否。內容點名、任務語氣、deadline、sender 角色等內容訊號可能比收件者路徑更重要；若 `level=broad` 且 `audienceSize` 很大，回覆使用者時可說明「這封信與你有關，但可能是部門/大群組通知；除非內容點名你的工作，否則不一定是高優先個人行動」。

計分模型：

- 先依 `routeDepth` 算基礎分數：`0 -> 100`、`1 -> 95`、`2 -> 90`。
- `routeDepth >= 3` 時使用 `round(90 * exp(-0.4 * (routeDepth - 2)))`，讓第三層後快速遞減。
- `audienceSize` 只作為次要 breadth factor，不用線性扣分；大群組會降低一部分分數並寫入 `reasons[]`，但不取代路徑深度。

### 取得通訊錄來源

Request:

```json
{}
```

呼叫 `POST /api/outlook/request-address-book-roots`，再讀 `POST /api/outlook/fetch-result-address-book-roots`。Fetch result 的 `data.roots[]` 會列出可讀取來源。每個 root 常用欄位：

- `id`
- `name`
- `entryCount`
- `canPageEntries`

### 分頁讀取人員與 group

Request:

```json
{
  "addressListId": "root-id",
  "addressListName": "root-name",
  "offset": 0,
  "pageSize": 100
}
```

`addressListId` 優先取自 `data.roots[].id`。Fetch result 的 `data.contacts[]` 會回傳人員、group、shared mailbox、resource 等 entries；若 `next.hasMore=true` 或 `data.hasMore=true`，用下一個 `offset` 繼續分頁，不要把第一頁當成全部資料。

呼叫 `POST /api/outlook/request-address-list-entries`，再讀 `POST /api/outlook/fetch-result-address-list-entries`。

### 展開 group members

已知 group 的 `smtpAddress` 或 `id` 後：

```json
{
  "groupSmtpAddress": "group@example.com",
  "groupId": "",
  "maxMembers": 5000,
  "forceRefresh": false
}
```

呼叫 `POST /api/outlook/request-address-book-group-members`，再讀 `POST /api/outlook/fetch-result-address-book-group-members`。Fetch result 的 `data.members[]` 是該 group 可取得的 members。若 member 本身也是 group，會以 `isGroup=true` 表示；需要查看下一層時，再對該 member 的 `smtpAddress` 或 `id` 發起一次 `request-address-book-group-members`。不要在沒有任務需求時遞迴展開所有 nested groups。

### Entries 內的顯示欄位

`AddressBookContactDto` 常用欄位：

- `displayName`
- `smtpAddress`
- `isGroup`
- `isLikelySelf`
- `isRelatedToSelf`
- `memberCount`
- `memberSmtpAddresses`
- `memberGroupSmtpAddresses`
- `memberOfGroupSmtpAddresses`

判讀規則：

- `isLikelySelf=true`：此 entry 代表目前使用者自己。
- group 的 `isRelatedToSelf=true`：該 group 與自己有關，通常表示自己是該 group 的 member，或 group 展開後能判定與自己相關。
- 個人的 `memberOfGroupSmtpAddresses[]`：該個人屬於哪些 group。
- group 的 `memberSmtpAddresses[]`：該 group 包含哪些 member SMTP。
- group 的 `memberGroupSmtpAddresses[]`：該 group 包含哪些 nested groups。
- 需要確認完整 direct members 時，呼叫 `request-address-book-group-members`，不要只依賴 entry 摘要欄位。

### 備用任務流程

確認「某個 group 是否包含自己」：

1. 優先用 `request-address-book-relation`，body 放 `groupSmtpAddress` 或 `groupId`。
2. 若 `state=found`，讀 `isRelatedToSelf`、`members[]`、`memberGroups[]` 與 `containingGroups[]`。
3. 若 `state=ambiguous`，請使用者確認 `matches[]` 裡的目標。

確認「某個人屬於哪些 group」：

1. 優先用 `request-address-book-relation`，body 放 `email`、`smtpAddress`、`displayName` 或 `id`。
2. 若 `state=found`，讀 `memberOfGroups[]` 與 `containingGroups[]`。
3. 若 `state=not_found` 或 `memberOfGroups[]` 為空，只能回報本次查詢未顯示所屬 group。

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

語意是將 folder 移到 Outlook default Deleted Items folder，不是永久刪除。HTTP request 不接受目的 folder；Deleted Items 由 SmartOffice API 以 Outlook default folder identity 定位，不依賴 `Deleted Items`、`刪除的郵件` 或其他本地化顯示名稱。完成後告知使用者 folder 已移到刪除資料夾。若目標 folder 已經位於 default Deleted Items folder 或其子層，paired fetch result 會回 `state=failed` / `message=manual_delete_required`；agent 必須停止，並請使用者自行到 Outlook 操作。

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

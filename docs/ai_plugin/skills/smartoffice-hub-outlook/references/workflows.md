# SmartOffice Outlook Workflows

## Shell Setup

```bash
export SMARTOFFICE_OUTLOOK_URL="${SMARTOFFICE_OUTLOOK_URL:-http://localhost:2805}"
```

以下範例使用 `curl`；若環境有 JSON parser，請用結構化 parser 取出 `commandId`，不要用脆弱字串切割。

## Wait For Command

1. 呼叫 `request-*`。
2. 從 response 取出 `commandId`。
3. 每 1-2 秒查：

```bash
curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/command-results/$COMMAND_ID"
```

4. `status` 是 `pending` 時繼續等；`completed` 且 `success=true` 才進下一步。
5. `failed`、`folder_cache_unavailable`、`timeout` 或其他非完成狀態都要回報使用者，不要假裝完成。

## Check API Status

```bash
curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/admin/status"
```

若 `connected=false`，回報目前 Outlook API 無法完成即時 request，並附上 status response 中可用的簡短診斷資訊。

## Load Folders

```bash
curl -sS -X POST "$SMARTOFFICE_OUTLOOK_URL/api/outlook/request-folders"
curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/command-results/$COMMAND_ID"
curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/folders"
```

若需要展開某個 folder 的 children：

```bash
curl -sS -X POST "$SMARTOFFICE_OUTLOOK_URL/api/outlook/request-folder-children" \
  -H "Content-Type: application/json" \
  -d '{"storeId":"store-id","parentEntryId":"entry-id","parentFolderPath":"\\\\Mailbox - User\\Inbox","maxDepth":1,"maxChildren":50}'
```

## Load Recent Mails

未指定 folder 時預設使用主要 mailbox 的 Inbox；不要主動載入或搜尋其他 folder。

```bash
curl -sS -X POST "$SMARTOFFICE_OUTLOOK_URL/api/outlook/request-mails" \
  -H "Content-Type: application/json" \
  -d '{"folderPath":"\\\\Mailbox - User\\Inbox","range":"1m","maxCount":30}'

curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/command-results/$COMMAND_ID"
curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/mails"
```

回覆使用者時優先摘要 `subject`、`senderName`、`receivedTime`、`categories`、`flagInterval` 等 metadata；不要主動輸出完整 body。
回覆中必須說明 folder 範圍，例如「範圍：主要 mailbox 的 Inbox」。

## Read One Mail Body

先從 `GET /api/outlook/mails` 或 `GET /api/outlook/mail-search` 找到目標 `id` 與 `folderPath`。

```bash
curl -sS -X POST "$SMARTOFFICE_OUTLOOK_URL/api/outlook/request-mail-body" \
  -H "Content-Type: application/json" \
  -d '{"mailId":"mail-id","folderPath":"\\\\Mailbox - User\\Inbox"}'

curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/command-results/$COMMAND_ID"
curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/mails"
```

只取任務需要的 body 片段；避免把整封信貼回對話。

## Read Mail Attachments

先從 `GET /api/outlook/mails` 或 `GET /api/outlook/mail-search` 找到目標 `id` 與 `folderPath`。

```bash
curl -sS -X POST "$SMARTOFFICE_OUTLOOK_URL/api/outlook/request-mail-attachments" \
  -H "Content-Type: application/json" \
  -d '{"mailId":"mail-id","folderPath":"\\\\Mailbox - User\\Inbox"}'

curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/command-results/$COMMAND_ID"
curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/mail-attachments?mailId=mail-id"
```

若要匯出附件，使用回傳的 `attachmentId` 或 `index`，不要用附件顯示名稱猜測。

## Search Mail

未指定 folder 時，`scopeFolderPaths` 只放主要 mailbox 的 Inbox。只有使用者明確要求其他 folder、子資料夾或全域搜尋時，才擴大 scope。

```bash
curl -sS -X POST "$SMARTOFFICE_OUTLOOK_URL/api/outlook/request-mail-search" \
  -H "Content-Type: application/json" \
  -d '{"scopeFolderPaths":["\\\\Mailbox - User\\Inbox"],"includeSubFolders":true,"keyword":"customer","textFields":["subject","sender"],"categoryNames":[],"hasAttachments":null,"flagState":"any","readState":"any","receivedFrom":null,"receivedTo":null}'
```

從 response 取得 `commandId` 與 `searchId`。等待時可查：

```bash
curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/mail-search/progress/$SEARCH_ID"
curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/command-results/$COMMAND_ID"
```

完成或需要檢視累積結果時：

```bash
curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/mail-search"
```

若使用者只說「收件夾」或沒有指定 folder，使用主要 mailbox 的 Inbox 作為 `scopeFolderPaths`。不要為了保險而搜尋所有 folder。
回覆中必須說明 folder 範圍；若使用 `scopeFolderPaths=["\\\\Mailbox - User\\Inbox"]`，要讓使用者知道結果只來自該 Inbox scope。

只依照「存在附件」搜尋時，`keyword` 可以是空字串：

```bash
curl -sS -X POST "$SMARTOFFICE_OUTLOOK_URL/api/outlook/request-mail-search" \
  -H "Content-Type: application/json" \
  -d '{"scopeFolderPaths":["\\\\Mailbox - User\\Inbox"],"includeSubFolders":true,"keyword":"","textFields":["subject"],"categoryNames":[],"hasAttachments":true,"flagState":"any","readState":"any","receivedFrom":"2026-05-01T00:00:00+08:00","receivedTo":"2026-06-01T00:00:00+08:00"}'
```

`hasAttachments=false` 可搜尋沒有附件的信；`hasAttachments=null` 代表不限附件。

## Update Mail Properties

先取得 `mailId` 與 `folderPath`，並向使用者確認會修改哪一封 mail。

```bash
curl -sS -X POST "$SMARTOFFICE_OUTLOOK_URL/api/outlook/request-update-mail-properties" \
  -H "Content-Type: application/json" \
  -d '{"mailId":"mail-id","folderPath":"\\\\Mailbox - User\\Inbox","isRead":true,"flagInterval":"today","flagRequest":"今天","taskStartDate":null,"taskDueDate":null,"taskCompletedDate":null,"categories":["Customer"],"newCategories":[]}'

curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/command-results/$COMMAND_ID"
curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/mails"
```

## Move Or Delete Mail

Move：

```bash
curl -sS -X POST "$SMARTOFFICE_OUTLOOK_URL/api/outlook/request-move-mail" \
  -H "Content-Type: application/json" \
  -d '{"mailId":"mail-id","sourceFolderPath":"\\\\Mailbox - User\\Inbox","destinationFolderPath":"\\\\Mailbox - User\\Projects"}'
```

Delete：

```bash
curl -sS -X POST "$SMARTOFFICE_OUTLOOK_URL/api/outlook/request-delete-mail" \
  -H "Content-Type: application/json" \
  -d '{"mailId":"mail-id","folderPath":"\\\\Mailbox - User\\Inbox"}'
```

完成後讀 `mails` 與 `folders`，確認 snapshot 已更新。`delete-mail` 只代表移到 Deleted Items。

## Calendar

```bash
curl -sS -X POST "$SMARTOFFICE_OUTLOOK_URL/api/outlook/request-calendar" \
  -H "Content-Type: application/json" \
  -d '{"daysForward":31,"startDate":null,"endDate":null}'

curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/command-results/$COMMAND_ID"
curl -sS "$SMARTOFFICE_OUTLOOK_URL/api/outlook/calendar"
```

## API Contract Reflection Checklist

若實作 workflow 時卡住，優先檢查：

- request response 是否包含足夠 correlation id，例如 `commandId`、`searchId`。
- cache endpoint 是否能取得 mutation 後的新狀態。
- 欄位是否讓 agent 不必用 display name 猜測目標。
- sensitive fields 是否有必要回傳；能否只查 metadata。
- destructive 操作是否可以用 snapshot 中的 id/path 做明確確認。

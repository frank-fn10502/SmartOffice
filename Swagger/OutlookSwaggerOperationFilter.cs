using Microsoft.OpenApi.Any;
using Microsoft.OpenApi.Models;
using SmartOffice.Hub.Models;
using Swashbuckle.AspNetCore.SwaggerGen;

namespace SmartOffice.Hub.Swagger
{
    /// <summary>
    /// 補齊 Swagger UI 使用者最需要的 Outlook API 分組、流程說明、request 範例與 response schema。
    /// </summary>
    public class OutlookSwaggerOperationFilter : IOperationFilter
    {
        private const string Json = "application/json";

        private static readonly Dictionary<string, OperationDocs> Docs = new(StringComparer.OrdinalIgnoreCase)
        {
            ["POST api/outlook/request-folders"] = new(
                "Outlook Operations",
                "要求 Outlook folder roots",
                "發起載入 stores 與 root folders 的 request。取得 `requestId` 後，用 `paired POST /api/outlook/fetch-result-*` 查狀態與資料。",
                typeof(OutlookRequestResponse)),
            ["POST api/outlook/request-folder-children"] = new(
                "Outlook Operations",
                "要求單一 folder 的 children",
                "建立載入單一 folder children 的 operation。HTTP API 的 folder path 使用 `/主要信箱 - User/Inbox`。`parentEntryId` 優先，`parentFolderPath` 可作為 fallback；SmartOffice API 會限制 `maxDepth` 1-3、`maxChildren` 1-200。",
                typeof(OutlookRequestResponse),
                FolderChildrenExample()),
            ["POST api/outlook/request-mails"] = new(
                "Outlook Operations",
                "要求指定 folder 的郵件列表",
                "`request-mails` 只發起 request，不直接代表最新郵件內容已在 response body。HTTP API 的 folder path 使用 `/主要信箱 - User/Inbox`。取得 `requestId` 後查 `paired POST /api/outlook/fetch-result-*`。",
                typeof(OutlookRequestResponse),
                FetchMailsExample()),
            ["POST api/outlook/request-mail-body"] = new(
                "Outlook Operations",
                "要求單封郵件 body",
                "Mail list 預設只載入 metadata。呼叫此 endpoint 後用 `paired fetch-result-* endpoint` 等待完成，再從 paired fetch-result-* data 讀取同一封 mail 的 `body` / `bodyHtml`。",
                typeof(OutlookRequestResponse),
                MailIdentityExample()),
            ["POST api/outlook/request-mail-attachments"] = new(
                "Attachments",
                "要求單封郵件附件 metadata",
                "建立讀取 attachment metadata 的 operation。完成後，附件 metadata 會更新到 mail / attachment state。",
                typeof(OutlookRequestResponse),
                MailIdentityExample()),
            ["POST api/outlook/request-export-mail-attachment"] = new(
                "Attachments",
                "要求匯出郵件附件",
                "建立匯出 attachment 的 operation。`exportRootPath` 可留空，SmartOffice API 會使用目前 attachment export settings。完成後使用 `exportedAttachmentId` 呼叫 `open-exported-attachment`。",
                typeof(OutlookRequestResponse),
                ExportAttachmentExample()),
            ["POST api/outlook/open-exported-attachment"] = new(
                "Attachments",
                "開啟已匯出的附件",
                "只接受 SmartOffice API 已記錄的 `exportedAttachmentId`，不接受任意檔案路徑，避免 Swagger 使用者誤把這個 endpoint 當成本機檔案 opener。",
                typeof(OpenExportedAttachmentResponse),
                OpenExportedAttachmentExample()),
            ["GET api/outlook/attachment-export-settings"] = new(
                "Attachments",
                "讀取附件匯出根目錄",
                "讀取 SmartOffice API 要求 AddIn 匯出附件時使用的 root path。避免在共用環境暴露敏感本機路徑。",
                typeof(AttachmentExportSettingsDto)),
            ["POST api/outlook/attachment-export-settings"] = new(
                "Attachments",
                "更新附件匯出根目錄",
                "更新 SmartOffice API 要求 AddIn 匯出附件時使用的 root path。避免在共用環境暴露敏感本機路徑。",
                typeof(AttachmentExportSettingsDto),
                AttachmentExportSettingsExample()),
            ["POST api/outlook/request-rules"] = new(
                "Outlook Operations",
                "要求 Outlook rules",
                "發起讀取 Outlook rules 的 request。完成後用 `paired fetch-result-* endpoint` 讀取資料。",
                typeof(OutlookRequestResponse)),
            ["POST api/outlook/request-categories"] = new(
                "Outlook Operations",
                "要求 Outlook master categories",
                "發起讀取 Outlook master categories 的 request。完成後用 `paired fetch-result-* endpoint` 讀取資料。",
                typeof(OutlookRequestResponse)),
            ["POST api/outlook/request-signalr-ping"] = new(
                "Diagnostics",
                "測試 Outlook AddIn SignalR channel",
                "透過正式 AddIn channel 建立 `ping` operation。用於確認 SmartOffice API 能否聯絡目前連線的 Outlook AddIn。",
                typeof(OutlookRequestResponse)),
            ["POST api/outlook/request-calendar"] = new(
                "Outlook Operations",
                "要求 Outlook calendar events",
                "`daysForward` 是相對今天的簡易範圍；若提供 `startDate` / `endDate`，AddIn 可依日期區間查詢。完成後用 `paired fetch-result-* endpoint` 讀取資料。",
                typeof(OutlookRequestResponse),
                CalendarExample()),
            ["POST api/outlook/request-update-mail-properties"] = new(
                "Outlook Operations",
                "更新單封郵件屬性",
                "正式的單封 mail mutation 入口，可一次更新 read state、flag/task 與 categories。舊的 marker-style endpoint 已移除。",
                typeof(OutlookRequestResponse),
                UpdateMailPropertiesExample()),
            ["POST api/outlook/request-upsert-category"] = new(
                "Outlook Operations",
                "新增或更新 Outlook master category",
                "以 category `name` 為識別新增或更新顏色與 shortcut key。完成後用 `paired fetch-result-* endpoint` 讀取資料。",
                typeof(OutlookRequestResponse),
                CategoryExample()),
            ["POST api/outlook/request-create-folder"] = new(
                "Outlook Operations",
                "建立 Outlook folder",
                "在 `parentFolderPath` 底下建立子 folder。完成後用 `paired fetch-result-* endpoint` 讀取資料。",
                typeof(OutlookRequestResponse),
                CreateFolderExample()),
            ["POST api/outlook/request-delete-folder"] = new(
                "Outlook Operations",
                "移動 Outlook folder 到刪除資料夾",
                "要求 AddIn 將指定 folder 移到 Outlook default Deleted Items folder；不得永久刪除。若目標已在 Deleted Items 內，paired fetch result 會回 `state=failed` / `message=manual_delete_required`，請使用者自行到 Outlook 刪除。呼叫前請確認 folder path 來自 folder fetch result。",
                typeof(OutlookRequestResponse),
                DeleteFolderExample()),
            ["POST api/outlook/request-move-mail"] = new(
                "Outlook Operations",
                "移動單封郵件",
                "將單封 mail 從 source folder 移到 destination folder。完成後用 `paired fetch-result-* endpoint` 或重新送出必要 request 確認結果。",
                typeof(OutlookRequestResponse),
                MoveMailExample()),
            ["POST api/outlook/request-move-mails"] = new(
                "Outlook Operations",
                "移動多封郵件",
                "將多封 mail 移到 destination folder。`mailIds` 必須來自 `paired fetch-result-* data` 或目前資料；單次最多 500 封，超過請由 caller 分批呼叫。完成後用 `paired fetch-result-* endpoint` 或重新送出必要 request 確認結果。",
                typeof(OutlookRequestResponse),
                MoveMailsExample()),
            ["POST api/outlook/request-delete-mail"] = new(
                "Outlook Operations",
                "刪除單封郵件",
                "要求 AddIn 將 mail 移到 Outlook default Deleted Items folder，不會永久刪除。若目標已在 Deleted Items 內，paired fetch result 會回 `state=failed` / `message=manual_delete_required`，請使用者自行到 Outlook 刪除。完成後用 `paired fetch-result-* endpoint` 或重新送出必要 request 確認結果。",
                typeof(OutlookRequestResponse),
                DeleteMailExample()),
            ["POST api/outlook/fetch-result-folders"] = FetchResultDocs(
                "Fetch Results",
                "取得 folders request 結果",
                "查詢 `request-folders` 或 `request-folder-children` 的狀態與分頁資料。`data` 包含 `stores` 與 `folders`。"),
            ["POST api/outlook/fetch-result-folder-children"] = FetchResultDocs(
                "Fetch Results",
                "取得 folder children request 結果",
                "查詢 `request-folder-children` 的狀態與分頁資料。`data` 包含 `stores` 與 `folders`。"),
            ["POST api/outlook/fetch-result-mails"] = FetchResultDocs(
                "Fetch Results",
                "取得 mails request 結果",
                "查詢 `request-mails` 的狀態與分頁資料。`data` 包含 `mails`。"),
            ["POST api/outlook/fetch-result-mail-body"] = FetchResultDocs(
                "Fetch Results",
                "取得 mail body request 結果",
                "查詢 `request-mail-body` 的狀態與分頁資料。`data.mails` 內同一封 mail 會帶有 `body` / `bodyHtml`。"),
            ["POST api/outlook/fetch-result-rules"] = FetchResultDocs(
                "Fetch Results",
                "取得 rules request 結果",
                "查詢 `request-rules` 的狀態與分頁資料。`data` 包含 `rules`。"),
            ["POST api/outlook/fetch-result-categories"] = FetchResultDocs(
                "Fetch Results",
                "取得 categories request 結果",
                "查詢 `request-categories` 的狀態與分頁資料。`data` 包含 `categories`。"),
            ["POST api/outlook/fetch-result-signalr-ping"] = FetchResultDocs(
                "Diagnostics",
                "取得 SignalR ping request 結果",
                "查詢 `request-signalr-ping` 的狀態。這是診斷 endpoint，不是一般資料讀取流程。"),
            ["POST api/outlook/fetch-result-calendar"] = FetchResultDocs(
                "Fetch Results",
                "取得 calendar request 結果",
                "查詢 `request-calendar` 的狀態與分頁資料。`data` 包含 `calendarEvents`。"),
            ["POST api/outlook/fetch-result-update-mail-properties"] = FetchResultDocs(
                "Fetch Results",
                "取得 mail properties update 結果",
                "查詢 `request-update-mail-properties` 的狀態與分頁資料。`data` 包含目前 mail snapshot。"),
            ["POST api/outlook/fetch-result-upsert-category"] = FetchResultDocs(
                "Fetch Results",
                "取得 category upsert 結果",
                "查詢 `request-upsert-category` 的狀態與分頁資料。`data` 包含 `categories`。"),
            ["POST api/outlook/fetch-result-create-folder"] = FetchResultDocs(
                "Fetch Results",
                "取得 create folder 結果",
                "查詢 `request-create-folder` 的狀態與分頁資料。`data` 包含 `stores` 與 `folders`。"),
            ["POST api/outlook/fetch-result-delete-folder"] = FetchResultDocs(
                "Fetch Results",
                "取得 delete folder 結果",
                "查詢 `request-delete-folder` 的狀態與分頁資料。`data` 包含 `stores` 與 `folders`。"),
            ["POST api/outlook/fetch-result-move-mail"] = FetchResultDocs(
                "Fetch Results",
                "取得 move mail 結果",
                "查詢 `request-move-mail` 的狀態與分頁資料。`data` 包含目前 mail snapshot。"),
            ["POST api/outlook/fetch-result-move-mails"] = FetchResultDocs(
                "Fetch Results",
                "取得 move mails 結果",
                "查詢 `request-move-mails` 的狀態與分頁資料。`data` 包含目前 mail snapshot。"),
            ["POST api/outlook/fetch-result-delete-mail"] = FetchResultDocs(
                "Fetch Results",
                "取得 delete mail 結果",
                "查詢 `request-delete-mail` 的狀態與分頁資料。`data` 包含目前 mail snapshot。"),
            ["POST api/outlook/fetch-result-mail-attachments"] = FetchResultDocs(
                "Attachments",
                "取得 mail attachments request 結果",
                "查詢 `request-mail-attachments` 的狀態與分頁資料。`data` 包含 `mailId`、`folderPath` 與 `attachments`。"),
            ["POST api/outlook/fetch-result-export-mail-attachment"] = FetchResultDocs(
                "Attachments",
                "取得 export attachment request 結果",
                "`request-export-mail-attachment` 完成狀態查詢。此 fetch result 目前不直接回傳 `exportedAttachmentId`，需要重新讀 attachment metadata。"),
            ["POST api/outlook/request-mail-search"] = new(
                "Mail Search",
                "搜尋 Outlook mails",
                "SmartOffice API 會先確保 folder data，展開 store/folder scope，再分成單 folder slices 送給 AddIn。使用 `paired fetch-result-* data` 查進度，完成或累積結果後讀取 `paired fetch-result-* data`。",
                typeof(OutlookRequestResponse),
                MailSearchExample()),
            ["POST api/outlook/fetch-result-mail-search"] = FetchResultDocs(
                "Mail Search",
                "取得 mail search 結果",
                "查詢 `request-mail-search` 的狀態與分頁資料。`data` 包含 `searchId` 與 `mails`。"),
            ["POST api/outlook/fetch-result-folder-mails"] = FetchResultDocs(
                "Mail Search",
                "取得 folder mails 結果",
                "查詢 `request-folder-mails` 的狀態與分頁資料。`data` 包含 `searchId` 與 `mails`。"),
            ["GET api/outlook/mail-search"] = new(
                "Mail Search",
                "取得 cached mail search results",
                "只讀取最近一次 mail search 累積在 SmartOffice API 的結果，不會觸發 Outlook 搜尋。若需要最新結果，先呼叫 `POST /api/outlook/request-mail-search`。",
                typeof(List<MailItemDto>)),
            ["GET api/outlook/mail-search/progress/{searchId}"] = new(
                "Mail Search",
                "用 searchId 查詢搜尋進度",
                "`searchId` 是 `request-mail-search` request 或 `data.searchId` 中的 correlation id。",
                typeof(MailSearchProgressDto)),
            ["GET api/outlook/mail-search/progress/by-command/{commandId}"] = new(
                "Mail Search",
                "用內部 commandId 查詢搜尋進度",
                "診斷用 endpoint；正式 caller 優先讀 `paired fetch-result-* data`。",
                typeof(MailSearchProgressDto)),
            ["GET api/outlook/mails"] = new(
                "Cached Snapshots",
                "取得 cached mails",
                "只讀取 SmartOffice API 目前記憶體中的 mail snapshot，不會呼叫 Outlook。若需要刷新，先呼叫 `POST /api/outlook/request-mails` 或 `POST /api/outlook/request-mail-search`。",
                typeof(List<MailItemDto>)),
            ["GET api/outlook/mail-attachments"] = new(
                "Attachments",
                "取得單封郵件的 cached attachment metadata",
                "只讀取 SmartOffice API 目前記憶體中的 attachment metadata。若需要刷新，先呼叫 `POST /api/outlook/request-mail-attachments`，等待 command 完成後用 query string `mailId` 查詢。",
                typeof(MailAttachmentsDto)),
            ["GET api/outlook/folders"] = new(
                "Cached Snapshots",
                "取得 cached folders",
                "只讀取 SmartOffice API 目前記憶體中的 folder snapshot。HTTP API 回傳的 folder path 使用 `/主要信箱 - User/Inbox`。若 folder tree 不完整，先呼叫 `request-folders` 或 `request-folder-children`。",
                typeof(FolderSnapshotDto)),
            ["GET api/outlook/rules"] = new(
                "Cached Snapshots",
                "取得 cached Outlook rules",
                "只讀取 SmartOffice API cache；若需要刷新，先呼叫 `POST /api/outlook/request-rules`。",
                typeof(List<OutlookRuleDto>)),
            ["GET api/outlook/categories"] = new(
                "Cached Snapshots",
                "取得 cached Outlook master categories",
                "只讀取 SmartOffice API cache；若需要刷新，先呼叫 `POST /api/outlook/request-categories`。",
                typeof(List<OutlookCategoryDto>)),
            ["GET api/outlook/calendar"] = new(
                "Cached Snapshots",
                "取得 cached calendar events",
                "只讀取 SmartOffice API cache；若需要刷新，先呼叫 `POST /api/outlook/request-calendar`。",
                typeof(List<CalendarEventDto>)),
            ["GET api/outlook/chat"] = new(
                "Chat",
                "取得 chat messages",
                "只讀取 SmartOffice API chat cache。chat text 可能含敏感 business data。",
                typeof(List<ChatMessageDto>)),
            ["POST api/outlook/chat"] = new(
                "Chat",
                "送出 chat message",
                "廣播 chat message 並可能觸發 mock Outlook 回覆。chat text 可能含敏感 business data。",
                typeof(ChatMessageDto),
                ChatExample()),
            ["GET api/outlook/command-results/{commandId}"] = new(
                "Command Results",
                "查詢指定 command 狀態",
                "診斷用 endpoint；正式 client workflow 應使用 paired `POST /api/outlook/fetch-result-*` 查詢 `state` 與分頁資料。",
                typeof(OutlookCommandStatusDto)),
            ["GET api/outlook/command-results"] = new(
                "Command Results",
                "查詢最近 command 狀態",
                "用於 diagnostics 或列出近期 dispatch 狀態；一般等待流程請優先使用 paired `fetch-result-*` endpoint。",
                typeof(List<OutlookCommandStatusDto>)),
            ["GET api/outlook/admin/status"] = new(
                "Diagnostics",
                "取得 Outlook AddIn 連線狀態",
                "讀取 SmartOffice API 對目前 AddIn SignalR 連線與最後 push/poll 的觀測狀態。",
                typeof(AddinStatusDto)),
            ["GET api/outlook/admin/logs"] = new(
                "Diagnostics",
                "取得 AddIn logs",
                "讀取 SmartOffice API 記錄的 AddIn diagnostic logs。",
                typeof(List<AddinLogEntry>)),
            ["POST api/outlook/admin/log"] = new(
                "Diagnostics",
                "寫入 AddIn log",
                "供 AddIn 回推 diagnostic log；一般 Web UI / AI client 不需要呼叫。",
                null,
                AddinLogExample()),
        };

        public void Apply(OpenApiOperation operation, OperationFilterContext context)
        {
            var path = context.ApiDescription.RelativePath ?? string.Empty;
            var method = context.ApiDescription.HttpMethod ?? string.Empty;
            if (!Docs.TryGetValue($"{method} {path}", out var docs))
                return;

            operation.Tags = new List<OpenApiTag> { new() { Name = docs.Tag } };
            operation.Summary = docs.Summary;
            operation.Description = docs.Description;
            operation.OperationId = BuildOperationId(context, path);

            if (docs.RequestExample is not null &&
                operation.RequestBody?.Content.TryGetValue(Json, out var mediaType) == true)
            {
                mediaType.Example = docs.RequestExample;
            }

            if (docs.ResponseType is not null)
                SetResponseSchema(operation, context, docs.ResponseType);
        }

        private static void SetResponseSchema(OpenApiOperation operation, OperationFilterContext context, Type responseType)
        {
            if (!operation.Responses.TryGetValue("200", out var response))
            {
                response = new OpenApiResponse { Description = "OK" };
                operation.Responses["200"] = response;
            }

            response.Content[Json] = new OpenApiMediaType
            {
                Schema = context.SchemaGenerator.GenerateSchema(responseType, context.SchemaRepository),
            };
        }

        private static string BuildOperationId(OperationFilterContext context, string path)
        {
            var method = context.ApiDescription.HttpMethod?.ToLowerInvariant() ?? "get";
            return $"{method}_{path.Replace("api/outlook/", string.Empty).Replace("/", "_").Replace("{", string.Empty).Replace("}", string.Empty).Replace("-", "_")}";
        }

        private static OpenApiObject FolderChildrenExample() => new()
        {
            ["storeId"] = new OpenApiString("store-primary"),
            ["parentEntryId"] = new OpenApiString("00000000A1B2C3D4"),
            ["parentFolderPath"] = new OpenApiString("/主要信箱 - User/Inbox"),
            ["maxDepth"] = new OpenApiInteger(1),
            ["maxChildren"] = new OpenApiInteger(50),
        };

        private static OpenApiObject FetchMailsExample() => new()
        {
            ["folderPath"] = new OpenApiString("/主要信箱 - User/Inbox"),
            ["lookbackHours"] = new OpenApiDouble(168),
            ["maxCount"] = new OpenApiInteger(30),
        };

        private static OpenApiObject MailIdentityExample() => new()
        {
            ["mailId"] = new OpenApiString("mail-20260506-001"),
            ["folderPath"] = new OpenApiString("/主要信箱 - User/Inbox"),
        };

        private static OpenApiObject ExportAttachmentExample() => new()
        {
            ["mailId"] = new OpenApiString("mail-20260506-001"),
            ["folderPath"] = new OpenApiString("/主要信箱 - User/Inbox"),
            ["attachmentId"] = new OpenApiString("attachment-001"),
            ["index"] = new OpenApiInteger(1),
            ["name"] = new OpenApiString("報價單.pdf"),
            ["fileName"] = new OpenApiString("報價單.pdf"),
            ["displayName"] = new OpenApiString("報價單.pdf"),
            ["exportRootPath"] = new OpenApiString(""),
        };

        private static OpenApiObject OpenExportedAttachmentExample() => new()
        {
            ["exportedAttachmentId"] = new OpenApiString("exported-attachment-001"),
        };

        private static OpenApiObject AttachmentExportSettingsExample() => new()
        {
            ["rootPath"] = new OpenApiString(@"C:\SmartOffice\AttachmentExports"),
        };

        private static OpenApiObject CalendarExample() => new()
        {
            ["daysForward"] = new OpenApiInteger(31),
            ["startDate"] = new OpenApiString("2026-05-06T00:00:00+08:00"),
            ["endDate"] = new OpenApiString("2026-06-06T23:59:59+08:00"),
        };

        private static OpenApiObject UpdateMailPropertiesExample() => new()
        {
            ["mailId"] = new OpenApiString("mail-20260506-001"),
            ["folderPath"] = new OpenApiString("/主要信箱 - User/Inbox"),
            ["isRead"] = new OpenApiBoolean(true),
            ["flagInterval"] = new OpenApiString("today"),
            ["flagRequest"] = new OpenApiString("今天"),
            ["taskStartDate"] = new OpenApiString("2026-05-06T00:00:00+08:00"),
            ["taskDueDate"] = new OpenApiString("2026-05-06T23:59:59+08:00"),
            ["taskCompletedDate"] = new OpenApiNull(),
            ["categories"] = new OpenApiArray { new OpenApiString("Customer"), new OpenApiString("Urgent") },
            ["newCategories"] = new OpenApiArray
            {
                new OpenApiObject
                {
                    ["name"] = new OpenApiString("Urgent"),
                    ["color"] = new OpenApiString("olCategoryColorRed"),
                    ["colorValue"] = new OpenApiInteger(1),
                    ["shortcutKey"] = new OpenApiString(""),
                },
            },
        };

        private static OpenApiObject CategoryExample() => new()
        {
            ["name"] = new OpenApiString("Customer"),
            ["color"] = new OpenApiString("olCategoryColorBlue"),
            ["colorValue"] = new OpenApiInteger(8),
            ["shortcutKey"] = new OpenApiString(""),
        };

        private static OpenApiObject CreateFolderExample() => new()
        {
            ["parentFolderPath"] = new OpenApiString("/主要信箱 - User/Inbox"),
            ["name"] = new OpenApiString("客戶追蹤"),
        };

        private static OpenApiObject DeleteFolderExample() => new()
        {
            ["folderPath"] = new OpenApiString("/主要信箱 - User/Inbox/客戶追蹤"),
        };

        private static OpenApiObject MoveMailExample() => new()
        {
            ["mailId"] = new OpenApiString("mail-20260506-001"),
            ["sourceFolderPath"] = new OpenApiString("/主要信箱 - User/Inbox"),
            ["destinationFolderPath"] = new OpenApiString("/主要信箱 - User/Inbox/客戶追蹤"),
        };

        private static OpenApiObject MoveMailsExample() => new()
        {
            ["mailIds"] = new OpenApiArray
            {
                new OpenApiString("mail-20260506-001"),
                new OpenApiString("mail-20260506-002"),
            },
            ["sourceFolderPath"] = new OpenApiString("/主要信箱 - User/Inbox"),
            ["sourceFolderPaths"] = new OpenApiArray
            {
                new OpenApiString("/主要信箱 - User/Inbox"),
            },
            ["destinationFolderPath"] = new OpenApiString("/主要信箱 - User/Inbox/客戶追蹤"),
            ["continueOnError"] = new OpenApiBoolean(true),
        };

        private static OpenApiObject DeleteMailExample() => new()
        {
            ["mailId"] = new OpenApiString("mail-20260506-001"),
            ["folderPath"] = new OpenApiString("/主要信箱 - User/Inbox"),
        };

        private static OperationDocs FetchResultDocs(string tag, string summary, string description) => new(
            tag,
            summary,
            $"{description} Request body 使用 `requestId`、`cursor` 與 `take`；正式 client 應輪詢到 `state=completed`，或遇到 `failed`、`unavailable`、`timeout` 後停止。",
            typeof(FetchResultResponse),
            FetchResultExample());

        private static OpenApiObject FetchResultExample() => new()
        {
            ["requestId"] = new OpenApiString("7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88"),
            ["cursor"] = new OpenApiString(""),
            ["take"] = new OpenApiInteger(100),
        };

        private static OpenApiObject MailSearchExample() => new()
        {
            ["searchId"] = new OpenApiString("6fb66d3a-7f4f-4a6d-9b3f-7e1e8c2f2d84"),
            ["storeId"] = new OpenApiString(""),
            ["scopeFolderPaths"] = new OpenApiArray { new OpenApiString("/主要信箱 - User/Inbox") },
            ["includeSubFolders"] = new OpenApiBoolean(true),
            ["keyword"] = new OpenApiString("客戶"),
            ["textFields"] = new OpenApiArray { new OpenApiString("subject"), new OpenApiString("sender"), new OpenApiString("body") },
            ["categoryNames"] = new OpenApiArray { new OpenApiString("Customer") },
            ["hasAttachments"] = new OpenApiBoolean(true),
            ["flagState"] = new OpenApiString("any"),
            ["readState"] = new OpenApiString("unread"),
            ["receivedFrom"] = new OpenApiString("2026-05-01T00:00:00+08:00"),
            ["receivedTo"] = new OpenApiString("2026-05-04T23:59:59+08:00"),
        };

        private static OpenApiObject ChatExample() => new()
        {
            ["source"] = new OpenApiString("web"),
            ["text"] = new OpenApiString("請摘要目前選取的 mails。"),
        };

        private static OpenApiObject AddinLogExample() => new()
        {
            ["level"] = new OpenApiString("info"),
            ["message"] = new OpenApiString("AddIn connected."),
            ["timestamp"] = new OpenApiString("2026-05-06T09:00:00+08:00"),
        };

        private sealed record OperationDocs(
            string Tag,
            string Summary,
            string Description,
            Type? ResponseType,
            OpenApiObject? RequestExample = null);
    }
}

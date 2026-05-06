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
                "Outlook Commands",
                "要求 Outlook folder roots",
                "Dispatch `fetch_folder_roots` 給 Outlook AddIn，載入 stores 與 root folders。完成後讀取 `GET /api/outlook/folders`。",
                typeof(FolderRequestDispatchResponse)),
            ["POST api/outlook/request-folder-children"] = new(
                "Outlook Commands",
                "要求單一 folder 的 children",
                "Dispatch `fetch_folder_children`。`parentEntryId` 優先，`parentFolderPath` 可作為 fallback；Hub 會限制 `maxDepth` 1-3、`maxChildren` 1-200。完成後讀取 `GET /api/outlook/folders`。",
                typeof(CommandDispatchResponse),
                FolderChildrenExample()),
            ["POST api/outlook/request-mails"] = new(
                "Outlook Commands",
                "要求指定 folder 的郵件列表",
                "`request-mails` 只會觸發 Outlook AddIn 載入資料，不直接代表最新郵件內容已在 response body。取得 `commandId` 後查 `GET /api/outlook/command-results/{commandId}`，完成後讀取 `GET /api/outlook/mails`。",
                typeof(CommandDispatchResponse),
                FetchMailsExample()),
            ["POST api/outlook/request-mail-body"] = new(
                "Outlook Commands",
                "要求單封郵件 body",
                "Mail list 預設只載入 metadata。呼叫此 endpoint 後等待 command 完成，再讀取 `GET /api/outlook/mails` 中同一封 mail 的 `body` / `bodyHtml`。",
                typeof(CommandDispatchResponse),
                MailIdentityExample()),
            ["POST api/outlook/request-mail-attachments"] = new(
                "Attachments",
                "要求單封郵件附件 metadata",
                "Dispatch `fetch_mail_attachments`。完成後，附件 metadata 會更新到 cached mail / attachment state，Web UI 會透過 SignalR 收到更新。",
                typeof(CommandDispatchResponse),
                MailIdentityExample()),
            ["POST api/outlook/request-export-mail-attachment"] = new(
                "Attachments",
                "要求匯出郵件附件",
                "Dispatch `export_mail_attachment` 給 Outlook AddIn。`exportRootPath` 可留空，Hub 會使用目前 attachment export settings。完成後使用 `exportedAttachmentId` 呼叫 `open-exported-attachment`。",
                typeof(CommandDispatchResponse),
                ExportAttachmentExample()),
            ["POST api/outlook/open-exported-attachment"] = new(
                "Attachments",
                "開啟已匯出的附件",
                "只接受 Hub 已記錄的 `exportedAttachmentId`，不接受任意檔案路徑，避免 Swagger 使用者誤把這個 endpoint 當成本機檔案 opener。",
                typeof(OpenExportedAttachmentResponse),
                OpenExportedAttachmentExample()),
            ["GET api/outlook/attachment-export-settings"] = new(
                "Attachments",
                "讀取附件匯出根目錄",
                "讀取 Hub 要求 AddIn 匯出附件時使用的 root path。避免在共用環境暴露敏感本機路徑。",
                typeof(AttachmentExportSettingsDto)),
            ["POST api/outlook/attachment-export-settings"] = new(
                "Attachments",
                "更新附件匯出根目錄",
                "更新 Hub 要求 AddIn 匯出附件時使用的 root path。避免在共用環境暴露敏感本機路徑。",
                typeof(AttachmentExportSettingsDto),
                AttachmentExportSettingsExample()),
            ["POST api/outlook/request-rules"] = new(
                "Outlook Commands",
                "要求 Outlook rules",
                "Dispatch `fetch_rules`。完成後讀取 `GET /api/outlook/rules`。",
                typeof(CommandDispatchResponse)),
            ["POST api/outlook/request-categories"] = new(
                "Outlook Commands",
                "要求 Outlook master categories",
                "Dispatch `fetch_categories`。完成後讀取 `GET /api/outlook/categories`。",
                typeof(CommandDispatchResponse)),
            ["POST api/outlook/request-signalr-ping"] = new(
                "Diagnostics",
                "測試 Outlook AddIn SignalR channel",
                "透過正式 AddIn channel dispatch `ping` command。用於確認 Hub 能否聯絡目前連線的 Outlook AddIn。",
                typeof(CommandDispatchResponse)),
            ["POST api/outlook/request-calendar"] = new(
                "Outlook Commands",
                "要求 Outlook calendar events",
                "`daysForward` 是相對今天的簡易範圍；若提供 `startDate` / `endDate`，AddIn 可依日期區間查詢。完成後讀取 `GET /api/outlook/calendar`。",
                typeof(CommandDispatchResponse),
                CalendarExample()),
            ["POST api/outlook/request-update-mail-properties"] = new(
                "Outlook Commands",
                "更新單封郵件屬性",
                "正式的單封 mail mutation 入口，可一次更新 read state、flag/task 與 categories。舊的 marker-style endpoint 已移除。",
                typeof(CommandDispatchResponse),
                UpdateMailPropertiesExample()),
            ["POST api/outlook/request-upsert-category"] = new(
                "Outlook Commands",
                "新增或更新 Outlook master category",
                "以 category `name` 為識別新增或更新顏色與 shortcut key。完成後讀取 `GET /api/outlook/categories`。",
                typeof(CommandDispatchResponse),
                CategoryExample()),
            ["POST api/outlook/request-create-folder"] = new(
                "Outlook Commands",
                "建立 Outlook folder",
                "在 `parentFolderPath` 底下建立子 folder。完成後讀取 `GET /api/outlook/folders`。",
                typeof(CommandDispatchResponse),
                CreateFolderExample()),
            ["POST api/outlook/request-delete-folder"] = new(
                "Outlook Commands",
                "刪除 Outlook folder",
                "要求 AddIn 刪除指定 folder。這是 destructive operation；呼叫前請確認 folder path 來自 `GET /api/outlook/folders`。",
                typeof(CommandDispatchResponse),
                DeleteFolderExample()),
            ["POST api/outlook/request-move-mail"] = new(
                "Outlook Commands",
                "移動單封郵件",
                "將單封 mail 從 source folder 移到 destination folder。完成後讀取 `GET /api/outlook/mails` 與 `GET /api/outlook/folders`。",
                typeof(CommandDispatchResponse),
                MoveMailExample()),
            ["POST api/outlook/request-delete-mail"] = new(
                "Outlook Commands",
                "刪除單封郵件",
                "要求 AddIn 將 mail 移到 Deleted Items，不應永久刪除。完成後讀取 `GET /api/outlook/mails` 與 `GET /api/outlook/folders`。",
                typeof(CommandDispatchResponse),
                DeleteMailExample()),
            ["POST api/outlook/request-mail-search"] = new(
                "Mail Search",
                "搜尋 Outlook mails",
                "Hub 會先確保 folder cache，展開 store/folder scope，再分成單 folder slices dispatch 給 AddIn。使用 `searchId` 查 `GET /api/outlook/mail-search/progress/{searchId}`，完成或累積結果後讀取 `GET /api/outlook/mail-search`。",
                typeof(MailSearchDispatchResponse),
                MailSearchExample()),
            ["GET api/outlook/mail-search"] = new(
                "Mail Search",
                "取得 cached mail search results",
                "只讀取最近一次 mail search 累積在 Hub 的結果，不會觸發 Outlook 搜尋。若需要最新結果，先呼叫 `POST /api/outlook/request-mail-search`。",
                typeof(List<MailItemDto>)),
            ["GET api/outlook/mail-search/progress/{searchId}"] = new(
                "Mail Search",
                "用 searchId 查詢搜尋進度",
                "`searchId` 是 `request-mail-search` request/response 中的 correlation id。",
                typeof(MailSearchProgressDto)),
            ["GET api/outlook/mail-search/progress/by-command/{commandId}"] = new(
                "Mail Search",
                "用 commandId 查詢搜尋進度",
                "只知道 dispatch response 的 `commandId` 時使用；一般流程可直接用 `searchId`。",
                typeof(MailSearchProgressDto)),
            ["GET api/outlook/mails"] = new(
                "Cached Snapshots",
                "取得 cached mails",
                "只讀取 Hub 目前記憶體中的 mail snapshot，不會呼叫 Outlook。若需要刷新，先呼叫 `POST /api/outlook/request-mails` 或 `POST /api/outlook/request-mail-search`。",
                typeof(List<MailItemDto>)),
            ["GET api/outlook/folders"] = new(
                "Cached Snapshots",
                "取得 cached folders",
                "只讀取 Hub 目前記憶體中的 folder snapshot。若 folder tree 不完整，先呼叫 `request-folders` 或 `request-folder-children`。",
                typeof(FolderSnapshotDto)),
            ["GET api/outlook/rules"] = new(
                "Cached Snapshots",
                "取得 cached Outlook rules",
                "只讀取 Hub cache；若需要刷新，先呼叫 `POST /api/outlook/request-rules`。",
                typeof(List<OutlookRuleDto>)),
            ["GET api/outlook/categories"] = new(
                "Cached Snapshots",
                "取得 cached Outlook master categories",
                "只讀取 Hub cache；若需要刷新，先呼叫 `POST /api/outlook/request-categories`。",
                typeof(List<OutlookCategoryDto>)),
            ["GET api/outlook/calendar"] = new(
                "Cached Snapshots",
                "取得 cached calendar events",
                "只讀取 Hub cache；若需要刷新，先呼叫 `POST /api/outlook/request-calendar`。",
                typeof(List<CalendarEventDto>)),
            ["GET api/outlook/chat"] = new(
                "Chat",
                "取得 chat messages",
                "只讀取 Hub chat cache。chat text 可能含敏感 business data。",
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
                "`request-*` 回傳 `commandId` 後，外部 client 應輪詢此 endpoint，直到 `status` 不是 `pending`。",
                typeof(OutlookCommandStatusDto)),
            ["GET api/outlook/command-results"] = new(
                "Command Results",
                "查詢最近 command 狀態",
                "用於 diagnostics 或列出近期 dispatch 狀態；一般等待流程請優先使用 `command-results/{commandId}`。",
                typeof(List<OutlookCommandStatusDto>)),
            ["GET api/outlook/admin/status"] = new(
                "Diagnostics",
                "取得 Outlook AddIn 連線狀態",
                "讀取 Hub 對目前 AddIn SignalR 連線與最後 push/poll 的觀測狀態。",
                typeof(AddinStatusDto)),
            ["GET api/outlook/admin/logs"] = new(
                "Diagnostics",
                "取得 AddIn logs",
                "讀取 Hub 記錄的 AddIn diagnostic logs。",
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
            ["parentFolderPath"] = new OpenApiString(@"\\主要信箱 - User\Inbox"),
            ["maxDepth"] = new OpenApiInteger(1),
            ["maxChildren"] = new OpenApiInteger(50),
        };

        private static OpenApiObject FetchMailsExample() => new()
        {
            ["folderPath"] = new OpenApiString(@"\\主要信箱 - User\Inbox"),
            ["range"] = new OpenApiString("1w"),
            ["maxCount"] = new OpenApiInteger(30),
        };

        private static OpenApiObject MailIdentityExample() => new()
        {
            ["mailId"] = new OpenApiString("mail-20260506-001"),
            ["folderPath"] = new OpenApiString(@"\\主要信箱 - User\Inbox"),
        };

        private static OpenApiObject ExportAttachmentExample() => new()
        {
            ["mailId"] = new OpenApiString("mail-20260506-001"),
            ["folderPath"] = new OpenApiString(@"\\主要信箱 - User\Inbox"),
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
            ["folderPath"] = new OpenApiString(@"\\主要信箱 - User\Inbox"),
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
            ["parentFolderPath"] = new OpenApiString(@"\\主要信箱 - User\Inbox"),
            ["name"] = new OpenApiString("客戶追蹤"),
        };

        private static OpenApiObject DeleteFolderExample() => new()
        {
            ["folderPath"] = new OpenApiString(@"\\主要信箱 - User\Inbox\客戶追蹤"),
        };

        private static OpenApiObject MoveMailExample() => new()
        {
            ["mailId"] = new OpenApiString("mail-20260506-001"),
            ["sourceFolderPath"] = new OpenApiString(@"\\主要信箱 - User\Inbox"),
            ["destinationFolderPath"] = new OpenApiString(@"\\主要信箱 - User\Inbox\客戶追蹤"),
        };

        private static OpenApiObject DeleteMailExample() => new()
        {
            ["mailId"] = new OpenApiString("mail-20260506-001"),
            ["folderPath"] = new OpenApiString(@"\\主要信箱 - User\Inbox"),
        };

        private static OpenApiObject MailSearchExample() => new()
        {
            ["searchId"] = new OpenApiString("6fb66d3a-7f4f-4a6d-9b3f-7e1e8c2f2d84"),
            ["storeId"] = new OpenApiString(""),
            ["scopeFolderPaths"] = new OpenApiArray { new OpenApiString(@"\\主要信箱 - User\Inbox") },
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

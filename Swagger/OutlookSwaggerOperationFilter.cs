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
        private const string RequestTag = "Outlook Requests";
        private const string FetchTag = "Fetch Results";
        private const string SearchTag = "Mail Search";
        private const string AttachmentTag = "Attachments";
        private const string HelperTag = "Lookup Helpers";
        private const string DiagnosticsTag = "Diagnostics";
        private const string ChatTag = "Chat";

        private static readonly Dictionary<string, OperationDocs> Docs = new(StringComparer.OrdinalIgnoreCase)
        {
            ["POST api/outlook/request-folders"] = new(
                RequestTag,
                "要求 Outlook folder roots",
                "建立載入 stores 與 root folders 的 request。HTTP response 只代表 request 已建立；取得 `requestId` 與 `data.fetchResultEndpoint` 後，請輪詢 paired fetch-result endpoint 查狀態與資料。",
                typeof(OutlookRequestResponse)),
            ["POST api/outlook/request-folder-children"] = new(
                RequestTag,
                "要求單一 folder 的 children",
                "建立載入單一 folder children 的 request。folder path 必須來自 folder fetch result，例如 `/主要信箱 - User/收件匣`。`parentEntryId` 優先，`parentFolderPath` 可作為 fallback；SmartOffice API 會限制 `maxDepth` 1-3、`maxChildren` 1-200。完成後輪詢 `data.fetchResultEndpoint`。",
                typeof(OutlookRequestResponse),
                FolderChildrenExample()),
            ["POST api/outlook/request-find-folder"] = new(
                RequestTag,
                "查找 Outlook folder",
                "封裝 folder discovery：SmartOffice API 會載入 folder roots 與尚未載入的 children，再用 `folderPath` 完全比對或用 `name` 比對候選 folders。完成後用 `paired POST /api/outlook/fetch-result-find-folder` 讀 `data.matchCount`、`data.isAmbiguous` 與 `data.folders`。若回多筆候選，caller 必須請使用者確認。",
                typeof(OutlookRequestResponse),
                FindFolderExample()),
            ["POST api/outlook/request-mails"] = new(
                RequestTag,
                "要求指定 folder 的郵件列表",
                "`request-mails` 只建立 request，不直接回傳郵件資料。folder path 必須來自 folder fetch result，例如 `/主要信箱 - User/收件匣`。取得 `requestId` 與 `data.fetchResultEndpoint` 後，輪詢 paired fetch-result endpoint 讀取 `data.mails`。",
                typeof(OutlookRequestResponse),
                FetchMailsExample()),
            ["POST api/outlook/request-mail-body"] = new(
                RequestTag,
                "要求單封郵件 body",
                "Mail list 預設只載入 metadata。建立 request 後用 `data.fetchResultEndpoint` 等待完成，再從 paired fetch-result data 讀取同一封 mail 的 `body` / `bodyHtml`。",
                typeof(OutlookRequestResponse),
                MailIdentityExample()),
            ["POST api/outlook/request-mail-attachments"] = new(
                AttachmentTag,
                "要求單封郵件附件 metadata",
                "建立讀取 attachment metadata 的 request。完成後用 `data.fetchResultEndpoint` 讀取 `data.attachments`。",
                typeof(OutlookRequestResponse),
                MailIdentityExample()),
            ["POST api/outlook/request-mail-conversation"] = new(
                RequestTag,
                "要求單封郵件對話脈絡",
                "建立讀取同一 conversation 郵件脈絡的 request。完成後用 `data.fetchResultEndpoint` 讀取 `data.mails`。",
                typeof(OutlookRequestResponse),
                MailIdentityExample()),
            ["POST api/outlook/request-export-mail-attachment"] = new(
                AttachmentTag,
                "要求匯出郵件附件",
                "建立匯出 attachment 的 operation。`exportRootPath` 可留空，SmartOffice API 會使用目前 attachment export settings。`fetch-result-export-mail-attachment` 目前不直接回傳 `exportedAttachmentId`；完成後請重新讀同一封 mail 的 attachment metadata，再用已記錄的 `exportedAttachmentId` 呼叫 `open-exported-attachment`。",
                typeof(OutlookRequestResponse),
                ExportAttachmentExample()),
            ["POST api/outlook/open-exported-attachment"] = new(
                AttachmentTag,
                "開啟已匯出的附件",
                "只接受 SmartOffice API 已記錄的 `exportedAttachmentId`，不接受任意檔案路徑，避免 Swagger 使用者誤把這個 endpoint 當成本機檔案 opener。",
                typeof(OpenExportedAttachmentResponse),
                OpenExportedAttachmentExample()),
            ["GET api/outlook/attachment-export-settings"] = new(
                AttachmentTag,
                "讀取附件匯出根目錄",
                "讀取附件匯出時使用的 root path。避免在共用環境暴露敏感本機路徑。",
                typeof(AttachmentExportSettingsDto)),
            ["POST api/outlook/attachment-export-settings"] = new(
                AttachmentTag,
                "更新附件匯出根目錄",
                "更新附件匯出時使用的 root path。避免在共用環境暴露敏感本機路徑。",
                typeof(AttachmentExportSettingsDto),
                AttachmentExportSettingsExample()),
            ["POST api/outlook/request-rules"] = new(
                RequestTag,
                "要求 Outlook rules",
                "建立讀取 Outlook rules 的 request。完成後用 `data.fetchResultEndpoint` 讀取 `data.rules`。",
                typeof(OutlookRequestResponse)),
            ["POST api/outlook/request-manage-rule"] = new(
                RequestTag,
                "建立或更新 Outlook rule",
                "建立 rule mutation request。完成後用 `data.fetchResultEndpoint` 讀取更新後的 `data.rules`；若 request failed，請顯示 paired fetch-result 回傳的 `message`。",
                typeof(OutlookRequestResponse)),
            ["POST api/outlook/request-categories"] = new(
                RequestTag,
                "要求 Outlook master categories",
                "建立讀取 Outlook master categories 的 request。完成後用 `data.fetchResultEndpoint` 讀取 `data.categories`。",
                typeof(OutlookRequestResponse)),
            ["POST api/outlook/request-signalr-ping"] = new(
                DiagnosticsTag,
                "測試 Outlook worker channel",
                "建立 `ping` diagnostics request，用於確認目前 Outlook worker channel 是否可用。這是診斷 endpoint，不是一般資料讀取流程；完成後輪詢 `fetch-result-signalr-ping`。",
                typeof(OutlookRequestResponse)),
            ["POST api/outlook/request-calendar"] = new(
                RequestTag,
                "要求 Outlook calendar events",
                "`daysForward` 是相對今天的簡易範圍；若提供 `startDate` / `endDate`，會依日期區間查詢。完成後用 `data.fetchResultEndpoint` 讀取 `data.calendarEvents`。",
                typeof(OutlookRequestResponse),
                CalendarExample()),
            ["POST api/outlook/request-address-book"] = new(
                RequestTag,
                "要求 Outlook address book",
                "建立通訊錄同步 request。完成後用 `data.fetchResultEndpoint` 讀取 `data.contacts`。",
                typeof(OutlookRequestResponse)),
            ["POST api/outlook/request-update-mail-properties"] = new(
                RequestTag,
                "更新單封郵件屬性",
                "單封 mail mutation 入口，可一次更新 read state、flag/task 與 categories。完成後用 `data.fetchResultEndpoint` 讀取目前 mail snapshot。",
                typeof(OutlookRequestResponse),
                UpdateMailPropertiesExample()),
            ["POST api/outlook/request-upsert-category"] = new(
                RequestTag,
                "新增或更新 Outlook master category",
                "以 category `name` 為識別新增或更新顏色與 shortcut key。完成後用 `data.fetchResultEndpoint` 讀取 `data.categories`。",
                typeof(OutlookRequestResponse),
                CategoryExample()),
            ["POST api/outlook/request-create-folder"] = new(
                RequestTag,
                "建立 Outlook folder",
                "在 `parentFolderPath` 底下建立子 folder。`parentFolderPath` 必須來自 folder fetch result。完成後用 `data.fetchResultEndpoint` 讀取更新後的 folder snapshot。",
                typeof(OutlookRequestResponse),
                CreateFolderExample()),
            ["POST api/outlook/request-delete-folder"] = new(
                RequestTag,
                "移動 Outlook folder 到刪除資料夾",
                "將指定 folder 移到 Outlook default Deleted Items folder；不得永久刪除。若目標已在 default Deleted Items folder 或其子層，paired fetch result 會回 `state=failed` / `message=manual_delete_required`，請使用者自行到 Outlook 操作。呼叫前請確認 folder path 來自 folder fetch result。",
                typeof(OutlookRequestResponse),
                DeleteFolderExample()),
            ["POST api/outlook/request-move-mail"] = new(
                RequestTag,
                "移動單封郵件",
                "將單封 mail 從 source folder 移到 destination folder。完成後用 `data.fetchResultEndpoint` 確認狀態；需要刷新列表時重新送出必要 request。",
                typeof(OutlookRequestResponse),
                MoveMailExample()),
            ["POST api/outlook/request-move-mails"] = new(
                RequestTag,
                "移動多封郵件",
                "將多封 mail 移到 destination folder。`mailIds` 必須來自 paired fetch-result data 或目前資料；單次最多 500 封，超過請由 caller 分批呼叫。完成後用 `data.fetchResultEndpoint` 確認狀態；需要刷新列表時重新送出必要 request。",
                typeof(OutlookRequestResponse),
                MoveMailsExample()),
            ["POST api/outlook/request-delete-mail"] = new(
                RequestTag,
                "刪除單封郵件",
                "將 mail 移到 Outlook default Deleted Items folder，不會永久刪除。完成後告知使用者 mail 已移到刪除資料夾；若要永久刪除，請使用者自行到 Outlook 操作。完成後用 `data.fetchResultEndpoint` 確認狀態；需要刷新列表時重新送出必要 request。",
                typeof(OutlookRequestResponse),
                DeleteMailExample()),
            ["POST api/outlook/fetch-result-folders"] = FetchResultDocs(
                FetchTag,
                "取得 folders request 結果",
                "查詢 `request-folders` 的狀態與分頁資料。`data` 包含 `stores` 與 root `folders`。若要查 `request-folder-children`，請使用 `fetch-result-folder-children`。"),
            ["POST api/outlook/fetch-result-folder-children"] = FetchResultDocs(
                FetchTag,
                "取得 folder children request 結果",
                "查詢 `request-folder-children` 的狀態與分頁資料。`data` 包含 `stores` 與 `folders`。"),
            ["POST api/outlook/fetch-result-find-folder"] = FetchResultDocs(
                FetchTag,
                "取得 find folder 結果",
                "查詢 `request-find-folder` 的狀態與分頁資料。`data` 包含 `query`、`matchCount`、`isAmbiguous`、`discoveryComplete`、`pendingDiscoveryTargets` 與候選 `folders`。"),
            ["POST api/outlook/fetch-result-mails"] = FetchResultDocs(
                FetchTag,
                "取得 mails request 結果",
                "查詢 `request-mails` 的狀態與分頁資料。`data` 包含 `mails`。"),
            ["POST api/outlook/fetch-result-mail-body"] = FetchResultDocs(
                FetchTag,
                "取得 mail body request 結果",
                "查詢 `request-mail-body` 的狀態與分頁資料。`data.mails` 內同一封 mail 會帶有 `body` / `bodyHtml`。"),
            ["POST api/outlook/fetch-result-mail-conversation"] = FetchResultDocs(
                FetchTag,
                "取得 mail conversation request 結果",
                "查詢 `request-mail-conversation` 的狀態與分頁資料。`data` 包含同一 conversation 的 `mails`。"),
            ["POST api/outlook/fetch-result-rules"] = FetchResultDocs(
                FetchTag,
                "取得 rules request 結果",
                "查詢 `request-rules` 的狀態與分頁資料。`data` 包含 `rules`。"),
            ["POST api/outlook/fetch-result-manage-rule"] = FetchResultDocs(
                FetchTag,
                "取得 manage rule 結果",
                "查詢 `request-manage-rule` 的狀態與分頁資料。`data` 包含更新後的 `rules`。"),
            ["POST api/outlook/fetch-result-categories"] = FetchResultDocs(
                FetchTag,
                "取得 categories request 結果",
                "查詢 `request-categories` 的狀態與分頁資料。`data` 包含 `categories`。"),
            ["POST api/outlook/fetch-result-signalr-ping"] = FetchResultDocs(
                DiagnosticsTag,
                "取得 SignalR ping request 結果",
                "查詢 `request-signalr-ping` 的狀態。這是診斷 endpoint，不是一般資料讀取流程。"),
            ["POST api/outlook/fetch-result-calendar"] = FetchResultDocs(
                FetchTag,
                "取得 calendar request 結果",
                "查詢 `request-calendar` 的狀態與分頁資料。`data` 包含 `calendarEvents`。"),
            ["POST api/outlook/fetch-result-address-book"] = FetchResultDocs(
                FetchTag,
                "取得 address book request 結果",
                "查詢 `request-address-book` 的狀態與分頁資料。`data` 包含 `contacts`。"),
            ["POST api/outlook/fetch-result-update-mail-properties"] = FetchResultDocs(
                FetchTag,
                "取得 mail properties update 結果",
                "查詢 `request-update-mail-properties` 的狀態與分頁資料。`data` 包含目前 mail snapshot。"),
            ["POST api/outlook/fetch-result-upsert-category"] = FetchResultDocs(
                FetchTag,
                "取得 category upsert 結果",
                "查詢 `request-upsert-category` 的狀態與分頁資料。`data` 包含 `categories`。"),
            ["POST api/outlook/fetch-result-create-folder"] = FetchResultDocs(
                FetchTag,
                "取得 create folder 結果",
                "查詢 `request-create-folder` 的狀態與分頁資料。`data` 包含 `stores` 與 `folders`。"),
            ["POST api/outlook/fetch-result-delete-folder"] = FetchResultDocs(
                FetchTag,
                "取得 delete folder 結果",
                "查詢 `request-delete-folder` 的狀態與分頁資料。`data` 包含 `stores` 與 `folders`。"),
            ["POST api/outlook/fetch-result-move-mail"] = FetchResultDocs(
                FetchTag,
                "取得 move mail 結果",
                "查詢 `request-move-mail` 的狀態與分頁資料。`data` 包含目前 mail snapshot。"),
            ["POST api/outlook/fetch-result-move-mails"] = FetchResultDocs(
                FetchTag,
                "取得 move mails 結果",
                "查詢 `request-move-mails` 的狀態與分頁資料。`data` 包含目前 mail snapshot。"),
            ["POST api/outlook/fetch-result-delete-mail"] = FetchResultDocs(
                FetchTag,
                "取得 delete mail 結果",
                "查詢 `request-delete-mail` 的狀態與分頁資料。`data` 包含目前 mail snapshot。"),
            ["POST api/outlook/fetch-result-mail-attachments"] = FetchResultDocs(
                AttachmentTag,
                "取得 mail attachments request 結果",
                "查詢 `request-mail-attachments` 的狀態與分頁資料。`data` 包含 `mailId`、`folderPath` 與 `attachments`。"),
            ["POST api/outlook/fetch-result-export-mail-attachment"] = FetchResultDocs(
                AttachmentTag,
                "取得 export attachment request 結果",
                "`request-export-mail-attachment` 完成狀態查詢。此 fetch result 目前不直接回傳 `exportedAttachmentId`，需要重新讀 attachment metadata。"),
            ["POST api/outlook/request-mail-search"] = new(
                SearchTag,
                "搜尋 Outlook mails",
                "搜尋或篩選 Outlook mails。`scopeFolderPaths` 或 `storeId` 必須至少提供一個；只有明確全域搜尋時才設定 `allowGlobalScope=true`。完成後用 paired fetch-result endpoint 讀取 `data.searchId` 與 `data.mails`。",
                typeof(OutlookRequestResponse),
                MailSearchExample()),
            ["POST api/outlook/request-folder-mails"] = new(
                SearchTag,
                "列出指定 folder 範圍的 mails",
                "列出指定 `folderPath` 範圍內的 mail metadata，適合批次搬移或統計前枚舉 ids。這是直接列出 folder mails，不是文字搜尋。`folderPath` 必須來自 folder fetch result。`includeSubFolders` 預設為 `true`；只有使用者明確排除 subfolders 時才設為 `false`。完成後用 paired fetch-result endpoint 讀 `data.folderMailsId` 與 `data.mails`。",
                typeof(OutlookRequestResponse),
                FolderMailsExample()),
            ["POST api/outlook/fetch-result-mail-search"] = FetchResultDocs(
                SearchTag,
                "取得 mail search 結果",
                "查詢 `request-mail-search` 的狀態與分頁資料。`data` 包含 `searchId` 與 `mails`。"),
            ["POST api/outlook/fetch-result-folder-mails"] = FetchResultDocs(
                SearchTag,
                "取得 folder mails 結果",
                "查詢 `request-folder-mails` 的狀態與分頁資料。`data` 包含 `folderMailsId` 與 `mails`。"),
            ["GET api/outlook/mail-search/progress/{searchId}"] = new(
                DiagnosticsTag,
                "用 searchId 查詢搜尋進度",
                "診斷用 endpoint；正式資料取得流程請輪詢 `fetch-result-mail-search`。`searchId` 是 `request-mail-search` request 或 `data.searchId` 中的 correlation id。",
                typeof(MailSearchProgressDto)),
            ["GET api/outlook/mail-search/progress/by-command/{commandId}"] = new(
                DiagnosticsTag,
                "用內部 commandId 查詢搜尋進度",
                "診斷用 endpoint；正式 caller 優先讀 paired fetch-result data。",
                typeof(MailSearchProgressDto)),
            ["GET api/outlook/address-book/lookup"] = new(
                HelperTag,
                "查詢單一 email 的通訊錄線索",
                "輕量查詢 endpoint，用於依 email 檢查目前已知的 mail 或 calendar 關係。需要完整通訊錄資料時，請使用 `request-address-book` 建立 request，再輪詢 `fetch-result-address-book`。",
                typeof(AddressBookLookupResponse)),
            ["GET api/outlook/chat"] = new(
                ChatTag,
                "取得 chat messages",
                "讀取目前已記錄的 chat messages。chat text 可能含敏感 business data。",
                typeof(List<ChatMessageDto>)),
            ["POST api/outlook/chat"] = new(
                ChatTag,
                "送出 chat message",
                "送出 chat message。chat text 可能含敏感 business data。",
                typeof(ChatMessageDto),
                ChatExample()),
            ["GET api/outlook/command-results/{commandId}"] = new(
                DiagnosticsTag,
                "查詢指定 command 狀態",
                "診斷用 endpoint；正式 client workflow 應使用 paired `POST /api/outlook/fetch-result-*` 查詢 `state` 與分頁資料。",
                typeof(OutlookCommandStatusDto)),
            ["GET api/outlook/command-results"] = new(
                DiagnosticsTag,
                "查詢最近 command 狀態",
                "用於 diagnostics 或列出近期 dispatch 狀態；一般等待流程請優先使用 paired `fetch-result-*` endpoint。",
                typeof(List<OutlookCommandStatusDto>)),
            ["GET api/outlook/admin/status"] = new(
                DiagnosticsTag,
                "取得 Outlook worker 連線狀態",
                "讀取 SmartOffice API 對目前 Outlook worker channel 與最後 push/poll 的觀測狀態。",
                typeof(AddinStatusDto)),
            ["GET api/outlook/admin/logs"] = new(
                DiagnosticsTag,
                "取得 Outlook worker logs",
                "讀取 SmartOffice API 記錄的 Outlook worker diagnostic logs。",
                typeof(List<AddinLogEntry>)),
            ["POST api/outlook/admin/log"] = new(
                DiagnosticsTag,
                "寫入 Outlook worker log",
                "供 Outlook worker 回推 diagnostic log；一般 Web UI / AI client 不需要呼叫。",
                null,
                AddinLogExample()),
        };

        public void Apply(OpenApiOperation operation, OperationFilterContext context)
        {
            var path = context.ApiDescription.RelativePath ?? string.Empty;
            var method = context.ApiDescription.HttpMethod ?? string.Empty;
            if (!Docs.TryGetValue($"{method} {path}", out var docs))
                docs = BuildFallbackDocs(method, path);

            if (docs is null)
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

        private static OperationDocs? BuildFallbackDocs(string method, string path)
        {
            if (!path.StartsWith("api/outlook/", StringComparison.OrdinalIgnoreCase))
                return null;

            var route = path["api/outlook/".Length..];
            if (method.Equals("POST", StringComparison.OrdinalIgnoreCase) &&
                route.StartsWith("request-", StringComparison.OrdinalIgnoreCase))
            {
                return new OperationDocs(
                    RequestTag,
                    $"建立 {route} request",
                    "建立 Outlook request。HTTP response 只代表 request 已建立；請使用 response 內的 `requestId` 與 `data.fetchResultEndpoint` 輪詢結果。",
                    typeof(OutlookRequestResponse));
            }

            if (method.Equals("POST", StringComparison.OrdinalIgnoreCase) &&
                route.StartsWith("fetch-result-", StringComparison.OrdinalIgnoreCase))
            {
                return FetchResultDocs(
                    FetchTag,
                    $"取得 {route} 結果",
                    "查詢 paired request 的狀態與分頁資料。");
            }

            if (route.StartsWith("admin/", StringComparison.OrdinalIgnoreCase) ||
                route.StartsWith("command-results", StringComparison.OrdinalIgnoreCase) ||
                route.StartsWith("mail-search/progress", StringComparison.OrdinalIgnoreCase))
            {
                return new OperationDocs(
                    DiagnosticsTag,
                    "Diagnostics endpoint",
                    "診斷用 endpoint；一般資料取得流程請使用 request endpoint 建立工作，再輪詢 paired fetch-result endpoint。",
                    null);
            }

            return null;
        }

        private static OpenApiObject FolderChildrenExample() => new()
        {
            ["storeId"] = new OpenApiString("store-primary"),
            ["parentEntryId"] = new OpenApiString("00000000A1B2C3D4"),
            ["parentFolderPath"] = new OpenApiString("/主要信箱 - User/收件匣"),
            ["maxDepth"] = new OpenApiInteger(1),
            ["maxChildren"] = new OpenApiInteger(50),
        };

        private static OpenApiObject FetchMailsExample() => new()
        {
            ["folderPath"] = new OpenApiString("/主要信箱 - User/收件匣"),
            ["lookbackHours"] = new OpenApiDouble(168),
            ["maxCount"] = new OpenApiInteger(30),
        };

        private static OpenApiObject FindFolderExample() => new()
        {
            ["name"] = new OpenApiString("folderAAA"),
            ["folderPath"] = new OpenApiString(""),
            ["folderType"] = new OpenApiString(""),
            ["storeId"] = new OpenApiString(""),
            ["includeHidden"] = new OpenApiBoolean(false),
            ["maxResults"] = new OpenApiInteger(20),
        };

        private static OpenApiObject MailIdentityExample() => new()
        {
            ["mailId"] = new OpenApiString("mail-20260506-001"),
            ["folderPath"] = new OpenApiString("/主要信箱 - User/收件匣"),
        };

        private static OpenApiObject ExportAttachmentExample() => new()
        {
            ["mailId"] = new OpenApiString("mail-20260506-001"),
            ["folderPath"] = new OpenApiString("/主要信箱 - User/收件匣"),
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
            ["folderPath"] = new OpenApiString("/主要信箱 - User/收件匣"),
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
            ["parentFolderPath"] = new OpenApiString("/主要信箱 - User/收件匣"),
            ["name"] = new OpenApiString("客戶追蹤"),
        };

        private static OpenApiObject DeleteFolderExample() => new()
        {
            ["folderPath"] = new OpenApiString("/主要信箱 - User/收件匣/客戶追蹤"),
        };

        private static OpenApiObject MoveMailExample() => new()
        {
            ["mailId"] = new OpenApiString("mail-20260506-001"),
            ["sourceFolderPath"] = new OpenApiString("/主要信箱 - User/收件匣"),
            ["destinationFolderPath"] = new OpenApiString("/主要信箱 - User/收件匣/客戶追蹤"),
        };

        private static OpenApiObject MoveMailsExample() => new()
        {
            ["mailIds"] = new OpenApiArray
            {
                new OpenApiString("mail-20260506-001"),
                new OpenApiString("mail-20260506-002"),
            },
            ["sourceFolderPath"] = new OpenApiString("/主要信箱 - User/收件匣"),
            ["sourceFolderPaths"] = new OpenApiArray
            {
                new OpenApiString("/主要信箱 - User/收件匣"),
            },
            ["destinationFolderPath"] = new OpenApiString("/主要信箱 - User/收件匣/客戶追蹤"),
            ["continueOnError"] = new OpenApiBoolean(true),
        };

        private static OpenApiObject DeleteMailExample() => new()
        {
            ["mailId"] = new OpenApiString("mail-20260506-001"),
            ["folderPath"] = new OpenApiString("/主要信箱 - User/收件匣"),
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
            ["scopeFolderPaths"] = new OpenApiArray { new OpenApiString("/主要信箱 - User/收件匣") },
            ["allowGlobalScope"] = new OpenApiBoolean(false),
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

        private static OpenApiObject FolderMailsExample() => new()
        {
            ["folderPath"] = new OpenApiString("/主要信箱 - User/Projects/folderA"),
            ["includeSubFolders"] = new OpenApiBoolean(true),
            ["receivedFrom"] = new OpenApiNull(),
            ["receivedTo"] = new OpenApiNull(),
        };

        private static OpenApiObject ChatExample() => new()
        {
            ["source"] = new OpenApiString("web"),
            ["text"] = new OpenApiString("請摘要目前選取的 mails。"),
        };

        private static OpenApiObject AddinLogExample() => new()
        {
            ["level"] = new OpenApiString("info"),
            ["message"] = new OpenApiString("Outlook worker connected."),
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

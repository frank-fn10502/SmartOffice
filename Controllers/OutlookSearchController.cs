using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using SmartOffice.Hub.Hubs;
using SmartOffice.Hub.Models;
using SmartOffice.Hub.Services;

namespace SmartOffice.Hub.Controllers
{
    [ApiController]
    [ApiExplorerSettings(GroupName = "outlook-v1")]
    [Route("api/outlook")]
    public class OutlookSearchController : ControllerBase
    {
        private readonly MailStore _mailStore;
        private readonly CommandResultStore _commandResults;
        private readonly OutlookCommandQueue _commandQueue;
        private readonly OutlookFolderCacheService _folderCache;
        private readonly IHubContext<NotificationHub> _hub;

        public OutlookSearchController(
            MailStore mailStore,
            CommandResultStore commandResults,
            OutlookCommandQueue commandQueue,
            OutlookFolderCacheService folderCache,
            IHubContext<NotificationHub> hub)
        {
            _mailStore = mailStore;
            _commandResults = commandResults;
            _commandQueue = commandQueue;
            _folderCache = folderCache;
            _hub = hub;
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求搜尋 mails；Hub 會展開 folder scope 並分片 dispatch 給 Outlook AddIn。
        /// </summary>
        [HttpPost("request-mail-search")]
        public IActionResult RequestMailSearch([FromBody] SearchMailsRequest req, CancellationToken ct)
        {
            req ??= new SearchMailsRequest();
            if (string.IsNullOrWhiteSpace(req.SearchId))
                req.SearchId = Guid.NewGuid().ToString();
            NormalizeMailSearchRequest(req);
            OutlookFolderPathMapper.NormalizeSearchRequest(req);

            var cmd = new PendingCommand
            {
                Type = "search_mails",
                SearchMailsRequest = req,
            };
            _commandResults.RecordDispatched(cmd);
            _ = Task.Run(() => _commandQueue.ExecuteExclusiveAsync(
                operationCt => RequestMailSearchQueuedAsync(cmd, req, operationCt),
                CancellationToken.None));
            return Ok(OperationAccepted(cmd, new { searchId = req.SearchId, resultKind = "mail_search" }));
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求列出指定 folder 範圍內的 mail metadata。
        /// </summary>
        [HttpPost("request-folder-mails")]
        public IActionResult RequestFolderMails([FromBody] FolderMailsRequest req, CancellationToken ct)
        {
            req ??= new FolderMailsRequest();
            if (string.IsNullOrWhiteSpace(req.FolderPath))
                return BadRequest(new { status = "missing_folder_path", message = "folderPath is required." });

            var search = new SearchMailsRequest
            {
                SearchId = Guid.NewGuid().ToString(),
                ScopeFolderPaths = new List<string> { req.FolderPath },
                IncludeSubFolders = req.IncludeSubFolders,
                Keyword = string.Empty,
                TextFields = new List<string> { "subject" },
                CategoryNames = new List<string>(),
                HasAttachments = null,
                FlagState = "any",
                ReadState = "any",
                ReceivedFrom = req.ReceivedFrom,
                ReceivedTo = req.ReceivedTo,
            };
            NormalizeMailSearchRequest(search);
            OutlookFolderPathMapper.NormalizeSearchRequest(search);

            var cmd = new PendingCommand
            {
                Type = "list_folder_mails",
                SearchMailsRequest = search,
            };
            _mailStore.MarkFolderMailSearch(search.SearchId);
            _commandResults.RecordDispatched(cmd);
            _ = Task.Run(() => _commandQueue.ExecuteExclusiveAsync(
                operationCt => RequestMailSearchQueuedAsync(cmd, search, operationCt, "/api/outlook/folder-mails"),
                CancellationToken.None));
            return Ok(OperationAccepted(cmd, new { searchId = search.SearchId, resultKind = "folder_mails" }));
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 取得 mail search results。
        /// </summary>
        [HttpGet("mail-search")]
        public IActionResult GetMailSearchResults()
        {
            return Ok(OutlookFolderPathMapper.ToApiMails(_mailStore.GetMailSearchResults()));
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 取得上次 folder mail request 的 results。
        /// </summary>
        [HttpGet("folder-mails")]
        public IActionResult GetFolderMails()
        {
            return Ok(OutlookFolderPathMapper.ToApiMails(_mailStore.GetFolderMailResults()));
        }

        [HttpPost("fetch-result-mail-search")]
        public IActionResult FetchResultMailSearch([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "search_mails");
        }

        [HttpPost("fetch-result-folder-mails")]
        public IActionResult FetchResultFolderMails([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "list_folder_mails");
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 查詢 mail search 進度。
        /// </summary>
        [HttpGet("mail-search/progress/{searchId}")]
        public IActionResult GetMailSearchProgress(string searchId)
        {
            var progress = _mailStore.GetMailSearchProgress(searchId);
            return progress is null ? NotFound() : Ok(OutlookFolderPathMapper.ToApiProgress(progress));
        }

        /// <summary>
        /// AI / MCP client 用 command id 查詢對應的 mail search 進度。
        /// </summary>
        [HttpGet("mail-search/progress/by-command/{commandId}")]
        public IActionResult GetMailSearchProgressByCommandId(string commandId)
        {
            var progress = _mailStore.GetMailSearchProgressByCommandId(commandId);
            return progress is null ? NotFound() : Ok(OutlookFolderPathMapper.ToApiProgress(progress));
        }

        private async Task<IActionResult> RequestMailSearchQueuedAsync(
            PendingCommand cmd,
            SearchMailsRequest req,
            CancellationToken ct,
            string resultEndpoint = "/api/outlook/mail-search")
        {
            var folderReady = await _folderCache.EnsureFolderCacheAsync(cmd, ct, loadPendingChildren: true);
            if (!folderReady)
            {
                _commandResults.RecordDispatched(cmd);
                _commandResults.RecordResult(new OutlookCommandResult
                {
                    CommandId = cmd.Id,
                    Success = false,
                    Message = "Hub could not load Outlook folders for mail search.",
                    Timestamp = DateTime.Now,
                });
                var failed = _mailStore.UpdateMailSearchProgress(new MailSearchProgressDto
                {
                    SearchId = req.SearchId,
                    CommandId = cmd.Id,
                    Status = "failed",
                    Phase = "load_folders",
                    Message = "folder_cache_unavailable",
                    Timestamp = DateTime.Now,
                });
                await _hub.Clients.All.SendAsync("MailSearchProgress", failed, ct);
                return Conflict(ResultEnvelope(
                    cmd.Id,
                    cmd.Type,
                    "failed",
                    "Hub could not load Outlook folders for mail search.",
                    new { searchId = req.SearchId }));
            }

            var slices = BuildMailSearchSlices(req, cmd.Id);
            if (slices.Count == 0)
            {
                _commandResults.RecordDispatched(cmd);
                _commandResults.RecordResult(new OutlookCommandResult
                {
                    CommandId = cmd.Id,
                    Success = false,
                    Message = "Hub loaded folders but no searchable folder matched the request.",
                    Timestamp = DateTime.Now,
                });
                var failed = _mailStore.UpdateMailSearchProgress(new MailSearchProgressDto
                {
                    SearchId = req.SearchId,
                    CommandId = cmd.Id,
                    Status = "failed",
                    Phase = "planning",
                    Message = "no_searchable_folder",
                    Timestamp = DateTime.Now,
                });
                await _hub.Clients.All.SendAsync("MailSearchProgress", failed, ct);
                return Conflict(ResultEnvelope(
                    cmd.Id,
                    cmd.Type,
                    "failed",
                    "Hub loaded folders but no searchable folder matched the request.",
                    new { searchId = req.SearchId }));
            }

            req.ScopeFolderPaths = slices.Select(slice => slice.FolderPath).ToList();
            var progress = _mailStore.StartMailSearchProgress(req, cmd.Id);
            await _hub.Clients.All.SendAsync("MailSearchProgress", progress, ct);
            _commandResults.RecordDispatched(cmd);
            var result = await DispatchMailSearchSlicesAsync(cmd, slices, ct);
            var body = new
            {
                requestId = cmd.Id,
                request = RequestName(cmd.Type),
                state = ResultState(result.Status),
                message = result.Message,
                data = new
                {
                    searchId = req.SearchId,
                    sliceCount = slices.Count,
                    resultEndpoint,
                },
            };
            return result.Success ? Ok(body) : StatusCode(result.HttpStatusCode, body);
        }

        private List<MailSearchSliceRequest> BuildMailSearchSlices(SearchMailsRequest req, string parentCommandId)
        {
            var snapshot = _mailStore.GetFolderSnapshot();
            var requestedPaths = req.ScopeFolderPaths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var searchableFolders = snapshot.Folders
                .Where(MailStore.IsSearchableMailFolder)
                .Where(folder => string.IsNullOrWhiteSpace(req.StoreId) || string.Equals(folder.StoreId, req.StoreId, StringComparison.OrdinalIgnoreCase))
                .ToList();

            var scopedFolders = requestedPaths.Count == 0
                ? searchableFolders
                : searchableFolders
                    .Where(folder => requestedPaths.Any(path =>
                        string.Equals(folder.FolderPath, path, StringComparison.OrdinalIgnoreCase)
                        || (req.IncludeSubFolders && folder.FolderPath.StartsWith($"{path}\\", StringComparison.OrdinalIgnoreCase))))
                    .ToList();

            var plannedFolders = scopedFolders
                .Where(folder =>
                    !string.IsNullOrWhiteSpace(folder.EntryId)
                    && !string.IsNullOrWhiteSpace(folder.FolderPath))
                .DistinctBy(folder => folder.FolderPath, StringComparer.OrdinalIgnoreCase)
                .OrderBy(folder => folder.StoreId)
                .ThenBy(folder => folder.FolderPath)
                .ToList();

            return plannedFolders
                .Select((folder, index) => new MailSearchSliceRequest
                {
                    SearchId = req.SearchId,
                    ParentCommandId = parentCommandId,
                    StoreId = folder.StoreId,
                    FolderEntryId = folder.EntryId,
                    FolderPath = folder.FolderPath,
                    Keyword = req.Keyword,
                    TextFields = new List<string>(req.TextFields),
                    CategoryNames = new List<string>(req.CategoryNames),
                    HasAttachments = req.HasAttachments,
                    FlagState = req.FlagState,
                    ReadState = req.ReadState,
                    ReceivedFrom = req.ReceivedFrom,
                    ReceivedTo = req.ReceivedTo,
                    SliceIndex = index,
                    SliceCount = plannedFolders.Count,
                    ResultBatchSize = 5,
                    ResetSearchResults = index == 0,
                    CompleteSearchOnSlice = index == plannedFolders.Count - 1,
                })
                .ToList();
        }

        private async Task<OutlookQueuedCommandResult> DispatchMailSearchSlicesAsync(PendingCommand parentCommand, List<MailSearchSliceRequest> slices, CancellationToken ct)
        {
            const int sliceDelayMs = 100;
            try
            {
                foreach (var slice in slices)
                {
                    var sliceCommand = new PendingCommand
                    {
                        Type = "fetch_mail_search_slice",
                        MailSearchSliceRequest = slice,
                    };
                    slice.CommandId = sliceCommand.Id;
                    var progress = _mailStore.UpdateMailSearchProgress(new MailSearchProgressDto
                    {
                        SearchId = slice.SearchId,
                        CommandId = parentCommand.Id,
                        Status = "running",
                        Phase = "folder",
                        ProcessedFolders = slice.SliceIndex,
                        TotalFolders = slice.SliceCount,
                        CurrentStoreId = slice.StoreId,
                        CurrentFolderPath = slice.FolderPath,
                        ResultCount = _mailStore.GetMailSearchResultCount(slice.SearchId),
                        Message = $"Dispatching mail search slice {slice.SliceIndex + 1}/{slice.SliceCount}.",
                        Timestamp = DateTime.Now,
                    });
                    await _hub.Clients.All.SendAsync("MailSearchProgress", progress, ct);

                    var sliceResult = await _commandQueue.ExecuteQueuedCommandAsync(sliceCommand, ensureReady: true, ct: ct);
                    if (!sliceResult.Success)
                    {
                        await CompleteSearchProgressAsync(parentCommand, sliceResult.Status, sliceResult.Message, ct);
                        var result = new OutlookCommandResult
                        {
                            CommandId = parentCommand.Id,
                            Success = false,
                            Message = sliceResult.Message,
                            Timestamp = DateTime.Now,
                        };
                        _commandResults.RecordResult(result);
                        await _hub.Clients.All.SendAsync("CommandResult", result, ct);
                        return OutlookQueuedCommandResult.Failed(parentCommand.Id, sliceResult.Status, sliceResult.Message);
                    }

                    progress = _mailStore.UpdateMailSearchProgress(new MailSearchProgressDto
                    {
                        SearchId = slice.SearchId,
                        CommandId = parentCommand.Id,
                        Status = "running",
                        Phase = "folder",
                        ProcessedFolders = slice.SliceIndex + 1,
                        TotalFolders = slice.SliceCount,
                        CurrentStoreId = slice.StoreId,
                        CurrentFolderPath = slice.FolderPath,
                        ResultCount = _mailStore.GetMailSearchResultCount(slice.SearchId),
                        Message = $"Completed mail search slice {slice.SliceIndex + 1}/{slice.SliceCount}.",
                        Timestamp = DateTime.Now,
                    });
                    await _hub.Clients.All.SendAsync("MailSearchProgress", progress, ct);
                    await Task.Delay(sliceDelayMs, ct);
                }

                await CompleteSearchProgressAsync(parentCommand, "completed", "Mail search completed by Hub folder slices.", ct);
                var completed = new OutlookCommandResult
                {
                    CommandId = parentCommand.Id,
                    Success = true,
                    Message = "Mail search completed by Hub folder slices.",
                    Timestamp = DateTime.Now,
                };
                _commandResults.RecordResult(completed);
                await _hub.Clients.All.SendAsync("CommandResult", completed, ct);
                return OutlookQueuedCommandResult.Completed(parentCommand.Id, "completed", completed.Message);
            }
            catch (Exception ex)
            {
                var result = new OutlookCommandResult
                {
                    CommandId = parentCommand.Id,
                    Success = false,
                    Message = $"Mail search dispatch failed: {ex.Message}",
                    Timestamp = DateTime.Now,
                };
                _commandResults.RecordResult(result);
                await CompleteSearchProgressAsync(parentCommand, "failed", result.Message, CancellationToken.None);
                await _hub.Clients.All.SendAsync("CommandResult", result, CancellationToken.None);
                return OutlookQueuedCommandResult.Failed(parentCommand.Id, "failed", result.Message);
            }
        }

        private void NormalizeMailSearchRequest(SearchMailsRequest req)
        {
            req.ScopeFolderPaths ??= new List<string>();
            req.TextFields = NormalizeMailSearchTextFields(req.TextFields);
            req.CategoryNames = NormalizeMailSearchCategoryNames(req.CategoryNames);
            req.FlagState = NormalizeMailSearchState(req.FlagState, new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "any", "flagged", "unflagged" });
            req.ReadState = NormalizeMailSearchState(req.ReadState, new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "any", "unread", "read" });
        }

        private IActionResult FetchResult(FetchResultRequest req, string expectedType)
        {
            req ??= new FetchResultRequest();
            if (string.IsNullOrWhiteSpace(req.RequestId))
                return BadRequest(new { state = "failed", message = "requestId is required." });

            var status = _commandResults.Get(req.RequestId);
            if (status is null)
                return NotFound(new { requestId = req.RequestId, request = "", state = "failed", message = "request not found", next = new FetchResultNext(), data = new { } });

            if (!string.Equals(status.Type, expectedType, StringComparison.OrdinalIgnoreCase))
                return BadRequest(new { requestId = req.RequestId, request = RequestName(status.Type), state = "failed", message = "requestId does not match this fetch-result endpoint.", next = new FetchResultNext(), data = new { } });

            var take = Math.Clamp(req.Take <= 0 ? 100 : req.Take, 1, 500);
            var offset = int.TryParse(req.Cursor, out var parsed) && parsed > 0 ? parsed : 0;
            var mails = expectedType == "list_folder_mails"
                ? OutlookFolderPathMapper.ToApiMails(_mailStore.GetFolderMailResults())
                : OutlookFolderPathMapper.ToApiMails(_mailStore.GetMailSearchResults());
            var page = Page(mails, offset, take);
            var progress = _mailStore.GetMailSearchProgressByCommandId(req.RequestId);

            return Ok(new FetchResultResponse
            {
                RequestId = status.CommandId,
                Request = RequestName(status.Type),
                State = ResultState(status.Status),
                Message = status.Message,
                Next = page.Next,
                Data = new
                {
                    searchId = progress?.SearchId ?? string.Empty,
                    mails = page.Items,
                },
            });
        }

        private static (List<T> Items, FetchResultNext Next) Page<T>(List<T> source, int offset, int take)
        {
            var total = source.Count;
            var safeOffset = Math.Clamp(offset, 0, total);
            var items = source.Skip(safeOffset).Take(take).ToList();
            var next = safeOffset + items.Count;
            var hasMore = next < total;
            return (items, new FetchResultNext { Cursor = hasMore ? next.ToString() : string.Empty, HasMore = hasMore });
        }

        private static object OperationAccepted(PendingCommand command, object data)
        {
            return ResultEnvelope(
                command.Id,
                command.Type,
                "accepted",
                "Request accepted. Poll the paired fetch-result-* endpoint for state and data.",
                data);
        }

        private static object ResultEnvelope(string requestId, string commandType, string state, string message, object? data = null)
        {
            return new
            {
                requestId,
                request = RequestName(commandType),
                state,
                message,
                data = data ?? new { },
            };
        }

        private static string ResultState(string status)
        {
            return status switch
            {
                "pending" => "running",
                "completed" or "mocked" => "completed",
                "addin_unavailable" => "unavailable",
                "timeout" => "timeout",
                _ => "failed",
            };
        }

        private static string RequestName(string commandType)
        {
            return commandType switch
            {
                "search_mails" => "request-mail-search",
                "list_folder_mails" => "request-folder-mails",
                _ => commandType,
            };
        }

        private static List<string> NormalizeMailSearchTextFields(List<string>? textFields)
        {
            var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "subject",
                "sender",
                "body",
            };
            var normalized = (textFields ?? new List<string>())
                .Where(field => allowed.Contains(field))
                .Select(field => field.ToLowerInvariant())
                .Distinct()
                .ToList();
            return normalized.Count > 0 ? normalized : new List<string> { "subject" };
        }

        private static List<string> NormalizeMailSearchCategoryNames(List<string>? categoryNames)
        {
            return (categoryNames ?? new List<string>())
                .Select(category => category.Trim())
                .Where(category => !string.IsNullOrWhiteSpace(category))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string NormalizeMailSearchState(string state, HashSet<string> allowed)
        {
            var normalized = (state ?? string.Empty).Trim().ToLowerInvariant();
            return allowed.Contains(normalized) ? normalized : "any";
        }

        private async Task CompleteSearchProgressAsync(PendingCommand cmd, string status, string message, CancellationToken ct)
        {
            if (cmd.SearchMailsRequest is null) return;
            var progress = _mailStore.GetMailSearchProgress(cmd.SearchMailsRequest.SearchId) ?? new MailSearchProgressDto
            {
                SearchId = cmd.SearchMailsRequest.SearchId,
                CommandId = cmd.Id,
            };
            progress.Status = status;
            progress.Phase = status == "completed" ? "completed" : "dispatch";
            progress.ResultCount = _mailStore.GetMailSearchResultCount(progress.SearchId);
            progress.Message = message;
            progress.Timestamp = DateTime.Now;
            progress = _mailStore.UpdateMailSearchProgress(progress);
            await _hub.Clients.All.SendAsync("MailSearchProgress", progress, ct);
        }
    }
}

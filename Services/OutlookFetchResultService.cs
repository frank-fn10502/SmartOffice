using Microsoft.AspNetCore.Mvc;
using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public class OutlookFetchResultService
    {
        private readonly MailStore _mailStore;
        private readonly CommandResultStore _commandResults;

        public OutlookFetchResultService(MailStore mailStore, CommandResultStore commandResults)
        {
            _mailStore = mailStore;
            _commandResults = commandResults;
        }

        public IActionResult FetchResult(FetchResultRequest req, params string[] expectedTypes)
        {
            req ??= new FetchResultRequest();
            if (string.IsNullOrWhiteSpace(req.RequestId))
                return new BadRequestObjectResult(new { state = "failed", message = "requestId is required." });

            var status = _commandResults.Get(req.RequestId);
            if (status is null)
                return new NotFoundObjectResult(new { requestId = req.RequestId, request = "", state = "failed", message = "request not found", next = new FetchResultNext(), data = new { } });

            if (expectedTypes.Length > 0 && !expectedTypes.Contains(status.Type, StringComparer.OrdinalIgnoreCase))
                return new BadRequestObjectResult(new { requestId = req.RequestId, request = RequestName(status.Type), state = "failed", message = "requestId does not match this fetch-result endpoint.", next = new FetchResultNext(), data = new { } });

            var take = Math.Clamp(req.Take <= 0 ? 100 : req.Take, 1, 500);
            var offset = int.TryParse(req.Cursor, out var parsed) && parsed > 0 ? parsed : 0;
            var command = _commandResults.GetRequestCommand(req.RequestId);
            var (data, next) = GetFetchResultData(status.Type, command, offset, take);

            return new OkObjectResult(new FetchResultResponse
            {
                RequestId = status.CommandId,
                Request = RequestName(status.Type),
                State = ResultState(status.Status),
                Message = status.Message,
                Next = next,
                Data = data,
            });
        }

        private (object Data, FetchResultNext Next) GetFetchResultData(string requestType, PendingCommand? command, int offset, int take)
        {
            switch (requestType)
            {
                case "fetch_folder_roots":
                case "fetch_folder_children":
                case "create_folder":
                case "delete_folder":
                {
                    var snapshot = OutlookFolderPathMapper.ToApiSnapshot(_mailStore.GetFolderSnapshot());
                    var page = Page(snapshot.Folders, offset, take);
                    return (new { stores = snapshot.Stores, folders = page.Items }, page.Next);
                }
                case "find_folder":
                {
                    var snapshot = OutlookFolderPathMapper.ToApiSnapshot(_mailStore.GetFolderSnapshot());
                    var matches = FindFolderMatches(snapshot.Stores, snapshot.Folders, command?.FindFolderRequest);
                    var requestedMaxResults = command?.FindFolderRequest?.MaxResults ?? 20;
                    var maxResults = Math.Clamp(requestedMaxResults <= 0 ? 20 : requestedMaxResults, 1, 100);
                    var page = Page(matches.Take(maxResults).ToList(), offset, take);
                    var pendingDiscoveryTargets = _mailStore.GetPendingFolderDiscoveryTargets().Count;
                    return (new
                    {
                        query = new
                        {
                            name = command?.FindFolderRequest?.Name ?? string.Empty,
                            folderPath = OutlookFolderPathMapper.ToApiPath(command?.FindFolderRequest?.FolderPath ?? string.Empty),
                            folderType = command?.FindFolderRequest?.FolderType ?? string.Empty,
                            storeId = command?.FindFolderRequest?.StoreId ?? string.Empty,
                        },
                        matchCount = matches.Count,
                        isAmbiguous = matches.Count > 1,
                        discoveryComplete = pendingDiscoveryTargets == 0,
                        pendingDiscoveryTargets,
                        folders = page.Items,
                    }, page.Next);
                }
                case "fetch_mails":
                case "update_mail_properties":
                case "move_mail":
                case "move_mails":
                case "delete_mail":
                {
                    var page = Page(OutlookFolderPathMapper.ToApiMails(_mailStore.GetMails()), offset, take);
                    return (new { mails = page.Items }, page.Next);
                }
                case "fetch_mail_body":
                {
                    var mails = OutlookFolderPathMapper.ToApiMails(_mailStore.GetMails());
                    var mail = FindRequestedMail(
                        mails,
                        command?.MailBodyRequest?.MailId,
                        command?.MailBodyRequest?.FolderPath);
                    var result = mail is null ? new List<MailItemDto>() : new List<MailItemDto> { mail };
                    return (new { mails = result }, new FetchResultNext());
                }
                case "fetch_mail_attachments":
                {
                    var mailId = command?.MailAttachmentsRequest?.MailId ?? string.Empty;
                    var attachments = string.IsNullOrWhiteSpace(mailId)
                        ? null
                        : _mailStore.GetMailAttachments(mailId);
                    var apiAttachments = attachments is null
                        ? null
                        : OutlookFolderPathMapper.ToApiAttachments(attachments);
                    var page = Page(apiAttachments?.Attachments ?? new List<MailAttachmentDto>(), offset, take);
                    return (new { mailId, folderPath = apiAttachments?.FolderPath ?? string.Empty, attachments = page.Items }, page.Next);
                }
                case "fetch_mail_conversation":
                {
                    var mailId = command?.MailConversationRequest?.MailId ?? string.Empty;
                    var conversation = string.IsNullOrWhiteSpace(mailId)
                        ? null
                        : _mailStore.GetMailConversation(mailId);
                    var apiConversation = conversation is null
                        ? new MailConversationDto { MailId = mailId, FolderPath = command?.MailConversationRequest?.FolderPath ?? string.Empty }
                        : OutlookFolderPathMapper.ToApiConversation(conversation);
                    var page = Page(apiConversation.Mails, offset, take);
                    return (new
                    {
                        mailId = apiConversation.MailId,
                        folderPath = apiConversation.FolderPath,
                        conversationId = apiConversation.ConversationId,
                        conversationTopic = apiConversation.ConversationTopic,
                        mails = page.Items,
                    }, page.Next);
                }
                case "export_mail_attachment":
                    return (new { }, new FetchResultNext());
                case "fetch_rules":
                case "manage_rule":
                {
                    var page = Page(_mailStore.GetRules(), offset, take);
                    return (new { rules = page.Items }, page.Next);
                }
                case "fetch_categories":
                case "upsert_category":
                {
                    var page = Page(_mailStore.GetCategories(), offset, take);
                    return (new { categories = page.Items }, page.Next);
                }
                case "fetch_calendar":
                {
                    var page = Page(_mailStore.GetCalendarEvents(), offset, take);
                    return (new { calendarEvents = page.Items }, page.Next);
                }
                default:
                    return (new { }, new FetchResultNext());
            }
        }

        private static MailItemDto? FindRequestedMail(List<MailItemDto> mails, string? mailId, string? folderPath)
        {
            if (string.IsNullOrWhiteSpace(mailId))
                return null;

            var apiFolderPath = OutlookFolderPathMapper.ToApiPath(folderPath ?? string.Empty);
            return mails.FirstOrDefault(mail =>
                string.Equals(mail.Id, mailId, StringComparison.OrdinalIgnoreCase)
                && (
                    string.IsNullOrWhiteSpace(apiFolderPath)
                    || string.Equals(mail.FolderPath, apiFolderPath, StringComparison.OrdinalIgnoreCase)
                ));
        }

        private static List<FolderDto> FindFolderMatches(List<OutlookStoreDto> stores, List<FolderDto> folders, FindFolderRequest? request)
        {
            request ??= new FindFolderRequest();
            var folderPath = OutlookFolderPathMapper.ToApiPath(request.FolderPath);
            var byPath = !string.IsNullOrWhiteSpace(folderPath);
            var name = request.Name.Trim();
            var folderType = OutlookFolderType.Unknown;
            var byFolderType = !string.IsNullOrWhiteSpace(request.FolderType)
                && Enum.TryParse<OutlookFolderType>(request.FolderType, ignoreCase: true, out folderType);
            var primaryStoreId = stores.FirstOrDefault()?.StoreId ?? string.Empty;
            var storeId = request.StoreId;
            if (byFolderType && string.IsNullOrWhiteSpace(storeId))
                storeId = primaryStoreId;

            return folders
                .Where(folder => request.IncludeHidden || !folder.IsHidden)
                .Where(folder => string.IsNullOrWhiteSpace(storeId)
                    || string.Equals(folder.StoreId, storeId, StringComparison.OrdinalIgnoreCase))
                .Where(folder => byPath
                    ? string.Equals(folder.FolderPath, folderPath, StringComparison.OrdinalIgnoreCase)
                    : byFolderType
                        ? folder.FolderType == folderType
                        : string.Equals(folder.Name, name, StringComparison.OrdinalIgnoreCase))
                .OrderBy(folder => folder.StoreId)
                .ThenBy(folder => folder.FolderPath)
                .ToList();
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
                "fetch_folder_roots" => "request-folders",
                "fetch_folder_children" => "request-folder-children",
                "find_folder" => "request-find-folder",
                "fetch_mails" => "request-mails",
                "fetch_mail_body" => "request-mail-body",
                "fetch_mail_attachments" => "request-mail-attachments",
                "fetch_mail_conversation" => "request-mail-conversation",
                "export_mail_attachment" => "request-export-mail-attachment",
                "fetch_rules" => "request-rules",
                "manage_rule" => "request-manage-rule",
                "fetch_categories" => "request-categories",
                "ping" => "request-signalr-ping",
                "fetch_calendar" => "request-calendar",
                "update_mail_properties" => "request-update-mail-properties",
                "upsert_category" => "request-upsert-category",
                "create_folder" => "request-create-folder",
                "delete_folder" => "request-delete-folder",
                "move_mail" => "request-move-mail",
                "move_mails" => "request-move-mails",
                "delete_mail" => "request-delete-mail",
                _ => commandType,
            };
        }
    }
}

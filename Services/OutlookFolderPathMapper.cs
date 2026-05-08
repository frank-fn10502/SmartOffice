using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    /// <summary>
    /// HTTP API 對外使用 `/Mailbox/Inbox`，Hub 內部與 AddIn SignalR contract 使用 `\\Mailbox\Inbox`。
    /// </summary>
    public static class OutlookFolderPathMapper
    {
        public static string ToAddinPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;

            var trimmed = path.Trim();
            if (trimmed.StartsWith(@"\\", StringComparison.Ordinal))
                return NormalizeAddinSeparators(trimmed);

            if (trimmed.StartsWith("/", StringComparison.Ordinal))
            {
                var segments = trimmed
                    .Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                return segments.Length == 0 ? string.Empty : $@"\\{string.Join("\\", segments)}";
            }

            if (trimmed.Contains('/', StringComparison.Ordinal))
            {
                var segments = trimmed
                    .Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                return segments.Length == 0 ? string.Empty : $@"\\{string.Join("\\", segments)}";
            }

            return NormalizeAddinSeparators(trimmed);
        }

        public static string ToApiPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;

            var trimmed = path.Trim();
            var segments = trimmed
                .Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            return segments.Length == 0 ? string.Empty : $"/{string.Join("/", segments)}";
        }

        public static void NormalizeRequest(PendingCommand command)
        {
            if (command.FolderDiscoveryRequest is not null)
                command.FolderDiscoveryRequest.ParentFolderPath = ToAddinPath(command.FolderDiscoveryRequest.ParentFolderPath);
            if (command.MailsRequest is not null)
                command.MailsRequest.FolderPath = ToAddinPath(command.MailsRequest.FolderPath);
            if (command.SearchMailsRequest is not null)
                NormalizeSearchRequest(command.SearchMailsRequest);
            if (command.MailSearchSliceRequest is not null)
                command.MailSearchSliceRequest.FolderPath = ToAddinPath(command.MailSearchSliceRequest.FolderPath);
            if (command.FolderMailsSliceRequest is not null)
                command.FolderMailsSliceRequest.FolderPath = ToAddinPath(command.FolderMailsSliceRequest.FolderPath);
            if (command.MailBodyRequest is not null)
                command.MailBodyRequest.FolderPath = ToAddinPath(command.MailBodyRequest.FolderPath);
            if (command.MailAttachmentsRequest is not null)
                command.MailAttachmentsRequest.FolderPath = ToAddinPath(command.MailAttachmentsRequest.FolderPath);
            if (command.ExportMailAttachmentRequest is not null)
                command.ExportMailAttachmentRequest.FolderPath = ToAddinPath(command.ExportMailAttachmentRequest.FolderPath);
            if (command.MailPropertiesRequest is not null)
                command.MailPropertiesRequest.FolderPath = ToAddinPath(command.MailPropertiesRequest.FolderPath);
            if (command.CreateFolderRequest is not null)
                command.CreateFolderRequest.ParentFolderPath = ToAddinPath(command.CreateFolderRequest.ParentFolderPath);
            if (command.DeleteFolderRequest is not null)
                command.DeleteFolderRequest.FolderPath = ToAddinPath(command.DeleteFolderRequest.FolderPath);
            if (command.MoveMailRequest is not null)
            {
                command.MoveMailRequest.SourceFolderPath = ToAddinPath(command.MoveMailRequest.SourceFolderPath);
                command.MoveMailRequest.DestinationFolderPath = ToAddinPath(command.MoveMailRequest.DestinationFolderPath);
            }
            if (command.MoveMailsRequest is not null)
            {
                command.MoveMailsRequest.SourceFolderPath = ToAddinPath(command.MoveMailsRequest.SourceFolderPath);
                command.MoveMailsRequest.SourceFolderPaths = command.MoveMailsRequest.SourceFolderPaths
                    .Select(ToAddinPath)
                    .Where(path => !string.IsNullOrWhiteSpace(path))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                command.MoveMailsRequest.DestinationFolderPath = ToAddinPath(command.MoveMailsRequest.DestinationFolderPath);
            }
            if (command.DeleteMailRequest is not null)
                command.DeleteMailRequest.FolderPath = ToAddinPath(command.DeleteMailRequest.FolderPath);
        }

        public static void NormalizeSearchRequest(SearchMailsRequest request)
        {
            request.ScopeFolderPaths = request.ScopeFolderPaths
                .Select(ToAddinPath)
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public static FolderSnapshotDto ToApiSnapshot(FolderSnapshotDto snapshot)
        {
            foreach (var store in snapshot.Stores)
                store.RootFolderPath = ToApiPath(store.RootFolderPath);

            foreach (var folder in snapshot.Folders)
            {
                folder.FolderPath = ToApiPath(folder.FolderPath);
                folder.ParentFolderPath = ToApiPath(folder.ParentFolderPath);
            }

            return snapshot;
        }

        public static List<MailItemDto> ToApiMails(List<MailItemDto> mails)
        {
            foreach (var mail in mails)
                mail.FolderPath = ToApiPath(mail.FolderPath);
            return mails;
        }

        public static MailAttachmentsDto ToApiAttachments(MailAttachmentsDto attachments)
        {
            attachments.FolderPath = ToApiPath(attachments.FolderPath);
            return attachments;
        }

        public static MailSearchProgressDto ToApiProgress(MailSearchProgressDto progress)
        {
            progress.CurrentFolderPath = ToApiPath(progress.CurrentFolderPath);
            return progress;
        }

        private static string NormalizeAddinSeparators(string path)
        {
            var segments = path
                .Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            return segments.Length == 0 ? string.Empty : $@"\\{string.Join("\\", segments)}";
        }
    }
}

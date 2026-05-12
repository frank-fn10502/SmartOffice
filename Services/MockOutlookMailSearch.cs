using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public static class MockOutlookMailSearch
    {
        public static List<MailItemDto> FilterMails(
            List<MailItemDto> mails,
            string defaultFolderPath,
            string folderPath,
            int maxCount,
            DateTime? receivedFrom = null,
            DateTime? receivedTo = null)
        {
            var target = string.IsNullOrWhiteSpace(folderPath) ? defaultFolderPath : folderPath;
            return mails
                .Where(mail =>
                    string.Equals(mail.FolderPath, target, StringComparison.OrdinalIgnoreCase)
                    && InReceivedTime(mail, receivedFrom, receivedTo))
                .OrderByDescending(mail => mail.ReceivedTime)
                .Take(Math.Max(1, maxCount))
                .Select(CloneMailMetadata)
                .ToList();
        }

        public static List<MailItemDto> FetchMailSearchSlice(List<MailItemDto> mails, MailSearchSliceRequest request)
        {
            var query = mails
                .Where(mail => string.Equals(mail.FolderPath, request.FolderPath, StringComparison.OrdinalIgnoreCase))
                .Where(mail => InReceivedTime(mail, request.ReceivedFrom, request.ReceivedTo))
                .Where(mail => MatchesSearchFilters(mail, request))
                .Where(mail => MatchesOutlookIndexSearch(mail, request))
                .OrderByDescending(mail => mail.ReceivedTime);

            return query.Select(CloneMailMetadata).ToList();
        }

        public static List<MailItemDto> FetchFolderMailsSlice(List<MailItemDto> mails, FolderMailsSliceRequest request)
        {
            return mails
                .Where(mail => string.Equals(mail.FolderPath, request.FolderPath, StringComparison.OrdinalIgnoreCase))
                .Where(mail => InReceivedTime(mail, request.ReceivedFrom, request.ReceivedTo))
                .OrderByDescending(mail => mail.ReceivedTime)
                .Select(CloneMailMetadata)
                .ToList();
        }

        private static bool MatchesOutlookIndexSearch(MailItemDto mail, MailSearchSliceRequest request)
        {
            var keyword = request.Keyword.Trim();
            if (string.IsNullOrWhiteSpace(keyword)) return true;
            return SearchTextValues(mail, request.TextFields)
                .Any(value => value.Contains(keyword, StringComparison.OrdinalIgnoreCase));
        }

        private static bool MatchesSearchFilters(MailItemDto mail, MailSearchSliceRequest request)
        {
            if (request.HasAttachments is true && mail.AttachmentCount <= 0) return false;
            if (request.HasAttachments is false && mail.AttachmentCount > 0) return false;
            if (request.FlagState == "flagged" && !mail.IsMarkedAsTask) return false;
            if (request.FlagState == "unflagged" && mail.IsMarkedAsTask) return false;
            if (request.ReadState == "unread" && mail.IsRead) return false;
            if (request.ReadState == "read" && !mail.IsRead) return false;
            if (request.CategoryNames.Count > 0)
            {
                var categories = SplitCategories(mail.Categories);
                if (!request.CategoryNames.Any(category => categories.Contains(category))) return false;
            }
            return true;
        }

        private static HashSet<string> SplitCategories(string categories)
        {
            return categories
                .Split(new[] { ',', '、', ';' }, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        private static IEnumerable<string> SearchTextValues(MailItemDto mail, List<string> textFields)
        {
            if (textFields.Any(field => string.Equals(field, "subject", StringComparison.OrdinalIgnoreCase))) yield return mail.Subject;
            if (textFields.Any(field => string.Equals(field, "sender", StringComparison.OrdinalIgnoreCase)))
            {
                yield return mail.Sender.DisplayName;
                yield return mail.Sender.SmtpAddress;
                yield return mail.Sender.RawAddress;
            }
            if (textFields.Any(field => string.Equals(field, "body", StringComparison.OrdinalIgnoreCase)))
            {
                yield return mail.Body;
                yield return mail.BodyHtml;
            }
        }

        private static bool InReceivedTime(MailItemDto mail, DateTime? receivedFrom, DateTime? receivedTo)
        {
            if (receivedFrom is not null && mail.ReceivedTime < receivedFrom.Value) return false;
            if (receivedTo is not null && mail.ReceivedTime > receivedTo.Value) return false;
            return true;
        }

        private static MailItemDto CloneMailMetadata(MailItemDto mail)
        {
            var clone = CloneMail(mail);
            clone.Body = string.Empty;
            clone.BodyHtml = string.Empty;
            return clone;
        }

        private static MailItemDto CloneMail(MailItemDto mail)
        {
            return new MailItemDto
            {
                Id = mail.Id,
                Subject = mail.Subject,
                Sender = CloneRecipient(mail.Sender),
                ToRecipients = CloneRecipients(mail.ToRecipients),
                CcRecipients = CloneRecipients(mail.CcRecipients),
                BccRecipients = CloneRecipients(mail.BccRecipients),
                ReceivedTime = mail.ReceivedTime,
                Body = mail.Body,
                BodyHtml = mail.BodyHtml,
                FolderPath = mail.FolderPath,
                ConversationId = mail.ConversationId,
                ConversationTopic = mail.ConversationTopic,
                ConversationIndex = mail.ConversationIndex,
                Categories = mail.Categories,
                IsRead = mail.IsRead,
                IsMarkedAsTask = mail.IsMarkedAsTask,
                AttachmentCount = mail.AttachmentCount,
                AttachmentNames = mail.AttachmentNames,
                FlagRequest = mail.FlagRequest,
                FlagInterval = mail.FlagInterval,
                TaskStartDate = mail.TaskStartDate,
                TaskDueDate = mail.TaskDueDate,
                TaskCompletedDate = mail.TaskCompletedDate,
                Importance = mail.Importance,
                Sensitivity = mail.Sensitivity,
            };
        }

        private static List<OutlookRecipientDto> CloneRecipients(List<OutlookRecipientDto> recipients)
        {
            return recipients.Select(CloneRecipient).ToList();
        }

        private static OutlookRecipientDto CloneRecipient(OutlookRecipientDto recipient)
        {
            return new OutlookRecipientDto
            {
                RecipientKind = recipient.RecipientKind,
                DisplayName = recipient.DisplayName,
                SmtpAddress = recipient.SmtpAddress,
                RawAddress = recipient.RawAddress,
                AddressType = recipient.AddressType,
                EntryUserType = recipient.EntryUserType,
                IsGroup = recipient.IsGroup,
                IsResolved = recipient.IsResolved,
                Members = CloneRecipients(recipient.Members),
            };
        }
    }
}

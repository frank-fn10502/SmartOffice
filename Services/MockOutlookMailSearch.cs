using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public static class MockOutlookMailSearch
    {
        public static List<MailItemDto> FilterMails(List<MailItemDto> mails, string defaultFolderPath, string folderPath, string range, int maxCount)
        {
            var target = string.IsNullOrWhiteSpace(folderPath) ? defaultFolderPath : folderPath;
            var since = RangeStart(range);
            return mails
                .Where(mail => mail.FolderPath == target && mail.ReceivedTime >= since)
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
                yield return mail.SenderName;
                yield return mail.SenderEmail;
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

        private static DateTime RangeStart(string range)
        {
            var now = DateTime.Now;
            return range switch
            {
                "1w" => now.AddDays(-7),
                "1m" => now.AddMonths(-1),
                _ => now.AddDays(-1),
            };
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
                SenderName = mail.SenderName,
                SenderEmail = mail.SenderEmail,
                ReceivedTime = mail.ReceivedTime,
                Body = mail.Body,
                BodyHtml = mail.BodyHtml,
                FolderPath = mail.FolderPath,
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
    }
}

using System.Globalization;
using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public static class MockOutlookMailSearch
    {
        public static List<MailItemDto> FilterMails(
            List<MailItemDto> mails,
            string defaultFolderPath,
            string folderPath,
            string range,
            int maxCount,
            string receivedFrom = "",
            string receivedTo = "")
        {
            var target = string.IsNullOrWhiteSpace(folderPath) ? defaultFolderPath : folderPath;
            var dateRange = MailListDateRange(range, receivedFrom, receivedTo);
            return mails
                .Where(mail =>
                    string.Equals(mail.FolderPath, target, StringComparison.OrdinalIgnoreCase)
                    && InReceivedTime(mail, dateRange.From, dateRange.To))
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

        private static (DateTime? From, DateTime? To) MailListDateRange(string range, string receivedFrom, string receivedTo)
        {
            var from = ParseFlexibleDateTime(receivedFrom, isEnd: false);
            var to = ParseFlexibleDateTime(receivedTo, isEnd: true);
            if (from is not null || to is not null) return (from, to);

            var rangeParts = SplitDateRange(range);
            if (rangeParts is not null)
            {
                return (
                    ParseFlexibleDateTime(rangeParts.Value.From, isEnd: false),
                    ParseFlexibleDateTime(rangeParts.Value.To, isEnd: true));
            }

            return (RangeStart(range), null);
        }

        private static (string From, string To)? SplitDateRange(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            var separators = new[] { "~", "～", "..", " 至 ", " 到 ", " - ", " – ", " — ", " to " };
            foreach (var separator in separators)
            {
                var parts = value.Split(separator, 2, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2) return (parts[0], parts[1]);
            }
            return null;
        }

        private static DateTime? ParseFlexibleDateTime(string value, bool isEnd)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            var trimmed = value.Trim();
            var styles = DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal;
            var cultures = new[] { CultureInfo.CurrentCulture, CultureInfo.GetCultureInfo("zh-TW"), CultureInfo.InvariantCulture };

            foreach (var culture in cultures)
            {
                if (DateTime.TryParse(trimmed, culture, styles, out var parsed))
                    return NormalizeDateBoundary(parsed, trimmed, isEnd);
            }

            return null;
        }

        private static DateTime NormalizeDateBoundary(DateTime value, string source, bool isEnd)
        {
            var hasTime = source.Contains(':', StringComparison.Ordinal) || source.Contains('T', StringComparison.OrdinalIgnoreCase);
            if (hasTime) return value;
            return isEnd ? value.Date.AddDays(1).AddTicks(-1) : value.Date;
        }

        private static DateTime RangeStart(string range)
        {
            var now = DateTime.Now;
            return range switch
            {
                "1d" => now.AddDays(-1),
                "1w" => now.AddDays(-7),
                "1m" or "30d" => now.AddDays(-30),
                "60d" => now.AddDays(-60),
                "90d" => now.AddDays(-90),
                _ => now.AddDays(-30),
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
                Sender = CloneRecipient(mail.Sender),
                ToRecipients = CloneRecipients(mail.ToRecipients),
                CcRecipients = CloneRecipients(mail.CcRecipients),
                BccRecipients = CloneRecipients(mail.BccRecipients),
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

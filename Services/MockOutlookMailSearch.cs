using SmartOffice.Hub.Models;
using System.Text.RegularExpressions;

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
            var maxCount = Math.Clamp(request.MaxCount <= 0 ? 50 : request.MaxCount, 1, 200);
            return mails
                .Where(mail => string.Equals(mail.FolderPath, request.FolderPath, StringComparison.OrdinalIgnoreCase))
                .Where(mail => InReceivedTime(mail, request.ReceivedFrom, request.ReceivedTo))
                .OrderByDescending(mail => mail.ReceivedTime)
                .Take(maxCount)
                .Select(mail => request.IncludeBody ? CloneMail(mail) : CloneMailMetadata(mail))
                .ToList();
        }

        private static bool InReceivedTime(MailItemDto mail, DateTime? receivedFrom, DateTime? receivedTo)
        {
            if (receivedFrom is not null && mail.ReceivedTime < receivedFrom.Value) return false;
            if (receivedTo is not null && mail.ReceivedTime > receivedTo.Value) return false;
            return true;
        }

        private static bool MatchesKeyword(MailItemDto mail, List<string> fields, string keyword, string matchMode)
        {
            if (string.IsNullOrWhiteSpace(keyword)) return true;
            var haystack = SearchHaystack(mail, fields);
            if (string.Equals(matchMode, "exact", StringComparison.OrdinalIgnoreCase))
                return haystack.Any(value => string.Equals(value, keyword, StringComparison.OrdinalIgnoreCase));
            if (string.Equals(matchMode, "regex", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    return haystack.Any(value => Regex.IsMatch(value, keyword, RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(100)));
                }
                catch
                {
                    return false;
                }
            }
            if (string.Equals(matchMode, "fuzzy", StringComparison.OrdinalIgnoreCase))
                return haystack.Any(value => FuzzyMatches(value, keyword));
            return haystack.Any(value => value.Contains(keyword, StringComparison.OrdinalIgnoreCase));
        }

        private static List<string> SearchHaystack(MailItemDto mail, List<string> fields)
        {
            var selected = fields.Count == 0 ? new List<string> { "subject" } : fields;
            var values = new List<string>();
            if (HasSearchField(selected, "subject")) values.Add(mail.Subject);
            if (HasSearchField(selected, "sender"))
            {
                values.Add(mail.SenderName);
                values.Add(mail.SenderEmail);
            }
            if (HasSearchField(selected, "categories")) values.Add(mail.Categories);
            if (HasSearchField(selected, "body"))
            {
                values.Add(mail.Body);
                values.Add(mail.BodyHtml);
            }
            return values;
        }

        private static bool HasSearchField(List<string> fields, string field)
        {
            return fields.Any(item => string.Equals(item, field, StringComparison.OrdinalIgnoreCase));
        }

        private static bool FuzzyMatches(string value, string keyword)
        {
            var normalizedValue = NormalizeFuzzyText(value);
            var normalizedKeyword = NormalizeFuzzyText(keyword);
            if (string.IsNullOrWhiteSpace(normalizedValue) || string.IsNullOrWhiteSpace(normalizedKeyword)) return false;
            if (normalizedValue.Contains(normalizedKeyword, StringComparison.OrdinalIgnoreCase)) return true;
            if (IsSubsequence(normalizedKeyword, normalizedValue)) return true;

            var keywordTokens = SplitFuzzyTokens(keyword);
            if (keywordTokens.Count == 0) return false;
            var valueTokens = SplitFuzzyTokens(value);
            return keywordTokens.All(term => valueTokens.Any(token => WithinFuzzyDistance(token, term)));
        }

        private static string NormalizeFuzzyText(string value)
        {
            return new string(value
                .Where(char.IsLetterOrDigit)
                .Select(char.ToLowerInvariant)
                .ToArray());
        }

        private static List<string> SplitFuzzyTokens(string value)
        {
            return Regex.Split(value.ToLowerInvariant(), @"[^\p{L}\p{Nd}]+")
                .Where(token => !string.IsNullOrWhiteSpace(token))
                .ToList();
        }

        private static bool WithinFuzzyDistance(string value, string keyword)
        {
            if (value.Contains(keyword, StringComparison.OrdinalIgnoreCase)) return true;
            var threshold = keyword.Length <= 4 ? 1 : Math.Max(1, keyword.Length / 4);
            return LevenshteinDistance(value, keyword) <= threshold;
        }

        private static bool IsSubsequence(string needle, string haystack)
        {
            var index = 0;
            foreach (var item in haystack)
            {
                if (index < needle.Length && item == needle[index]) index++;
                if (index == needle.Length) return true;
            }
            return false;
        }

        private static int LevenshteinDistance(string left, string right)
        {
            if (left.Length == 0) return right.Length;
            if (right.Length == 0) return left.Length;

            var previous = Enumerable.Range(0, right.Length + 1).ToArray();
            var current = new int[right.Length + 1];
            for (var i = 1; i <= left.Length; i++)
            {
                current[0] = i;
                for (var j = 1; j <= right.Length; j++)
                {
                    var cost = left[i - 1] == right[j - 1] ? 0 : 1;
                    current[j] = Math.Min(
                        Math.Min(current[j - 1] + 1, previous[j] + 1),
                        previous[j - 1] + cost);
                }
                (previous, current) = (current, previous);
            }
            return previous[right.Length];
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

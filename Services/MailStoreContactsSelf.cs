namespace SmartOffice.Hub.Services
{
    public partial class MailStore
    {
        private HashSet<string> InferSelfAddresses()
        {
            var sentFolders = _folders
                .Where(folder => folder.FolderType == OutlookFolderType.Sent)
                .Select(folder => folder.FolderPath)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var addresses = _stores
                .SelectMany(store => ExtractEmailLikeValues(store.AccountSmtpAddress)
                    .Concat(ExtractEmailLikeValues(store.StoreFilePath))
                    .Concat(ExtractEmailLikeValues(store.DisplayName))
                    .Concat(ExtractEmailLikeValues(store.RootFolderPath)))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var sender in _mails
                .Concat(_mailSearchResults)
                .Concat(_folderMailResults)
                .Concat(_mailSearchResultsBySearchId.Values.SelectMany(items => items))
                .Concat(_folderMailResultsByRequestId.Values.SelectMany(items => items))
                .Where(mail => sentFolders.Contains(mail.FolderPath))
                .Select(mail => NormalizeEmail(mail.Sender.SmtpAddress))
                .Where(email => !string.IsNullOrWhiteSpace(email)))
            {
                addresses.Add(sender);
            }

            return addresses;
        }

        private static IEnumerable<string> ExtractEmailLikeValues(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) yield break;
            var matches = System.Text.RegularExpressions.Regex.Matches(
                value,
                "[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,}",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            foreach (System.Text.RegularExpressions.Match match in matches)
                yield return NormalizeEmail(match.Value);
        }
    }
}

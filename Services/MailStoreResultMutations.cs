using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public partial class MailStore
    {
        public void RemoveMailFromCachedResults(string? mailId)
        {
            if (string.IsNullOrWhiteSpace(mailId)) return;

            lock (_lock)
            {
                RemoveMail(_mails, mailId);
                RemoveMail(_mailSearchResults, mailId);
                RemoveMail(_folderMailResults, mailId);
                foreach (var results in _mailSearchResultsBySearchId.Values)
                    RemoveMail(results, mailId);
                foreach (var results in _folderMailResultsByRequestId.Values)
                    RemoveMail(results, mailId);
                _attachments.Remove(mailId);
                _conversations.Remove(mailId);
            }
        }

        private static void RemoveMail(List<MailItemDto> mails, string mailId)
        {
            mails.RemoveAll(mail => string.Equals(mail.Id, mailId, StringComparison.OrdinalIgnoreCase));
        }
    }
}

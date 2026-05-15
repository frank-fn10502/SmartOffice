using SmartOffice.Hub.Contracts;

namespace SmartOffice.Hub.Services
{
    public partial class MailStore
    {
        private static void ApplySelfGroupRelations(List<AddressBookContactDto> contacts)
        {
            var selfGroups = contacts
                .Where(contact => contact.IsGroup && contact.IsRelatedToSelf)
                .Select(contact => NormalizeEmail(contact.SmtpAddress))
                .Concat(contacts.Where(contact => contact.IsLikelySelf).SelectMany(contact => contact.MemberOfGroupSmtpAddresses).Select(NormalizeEmail))
                .Where(email => !string.IsNullOrWhiteSpace(email))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (selfGroups.Count == 0) return;

            var sameGroupContacts = contacts.Where(contact => !contact.IsGroup && !contact.IsLikelySelf
                && contact.MemberOfGroupSmtpAddresses.Select(NormalizeEmail).Any(group => selfGroups.Contains(group)));
            foreach (var contact in sameGroupContacts) { contact.IsRelatedToSelf = true; contact.RelationScore += 30; }
        }
    }
}

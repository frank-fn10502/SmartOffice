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

        private static int AddressBookContactPriority(AddressBookContactDto contact)
        {
            var score = 0;
            if (contact.IsLikelySelf) score += 100000;
            if (contact.IsRelatedToSelf) score += 70000;
            if (HasUserVisibleOutlookContext(contact)) score += 30000;
            if (contact.MemberOfGroupSmtpAddresses.Count > 0) score += 15000;
            if (contact.IsGroup && (contact.MemberCount > 0 || contact.MemberSmtpAddresses.Count > 0)) score += 5000;
            return score;
        }

        private static bool HasUserVisibleOutlookContext(AddressBookContactDto contact)
        {
            if (contact.MailCount > 0 || contact.CalendarCount > 0) return true;
            if (contact.Sources.Any(source => source.Equals("mail", StringComparison.OrdinalIgnoreCase)
                || source.Equals("calendar", StringComparison.OrdinalIgnoreCase)
                || source.Equals("mail_recipient", StringComparison.OrdinalIgnoreCase))) return true;
            return contact.RelationKinds.Any(kind => !kind.Equals("address_book", StringComparison.OrdinalIgnoreCase));
        }
    }
}

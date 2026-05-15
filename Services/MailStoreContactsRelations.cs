using SmartOffice.Hub.Contracts;

namespace SmartOffice.Hub.Services
{
    public partial class MailStore
    {
        private static void ApplySelfGroupRelations(List<AddressBookContactDto> contacts)
        {
            var selfKeys = contacts
                .Where(contact => contact.IsLikelySelf)
                .SelectMany(ContactKeys)
                .Where(key => !string.IsNullOrWhiteSpace(key))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (selfKeys.Count == 0) return;

            var groupContacts = contacts
                .Where(contact => contact.IsGroup)
                .Where(contact => !string.IsNullOrWhiteSpace(NormalizeEmail(contact.SmtpAddress)))
                .ToList();
            var selfGroups = contacts
                .Where(contact => contact.IsLikelySelf)
                .SelectMany(contact => contact.MemberOfGroupSmtpAddresses)
                .Select(NormalizeEmail)
                .Concat(groupContacts
                    .Where(group => group.IsRelatedToSelf
                        || group.MemberSmtpAddresses.Select(NormalizeEmail).Any(member => selfKeys.Contains(member)))
                    .Select(group => NormalizeEmail(group.SmtpAddress)))
                .Where(email => !string.IsNullOrWhiteSpace(email))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (selfGroups.Count == 0) return;

            var changed = true;
            while (changed)
            {
                changed = false;
                foreach (var group in groupContacts)
                {
                    var groupEmail = NormalizeEmail(group.SmtpAddress);
                    if (string.IsNullOrWhiteSpace(groupEmail) || selfGroups.Contains(groupEmail)) continue;
                    var includesSelfGroup = group.MemberGroupSmtpAddresses
                        .Concat(group.MemberSmtpAddresses)
                        .Select(NormalizeEmail)
                        .Any(member => selfGroups.Contains(member));
                    if (includesSelfGroup) changed = selfGroups.Add(groupEmail);
                }
            }

            foreach (var group in groupContacts.Where(group => selfGroups.Contains(NormalizeEmail(group.SmtpAddress))))
            {
                group.IsRelatedToSelf = true;
                group.RelationScore += 60;
            }

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

        private static void AddRecipientGroupMembership(
            AddressBookAccumulator groupContact,
            AddressBookAccumulator? memberContact,
            OutlookRecipientDto member)
        {
            if (memberContact is null || !groupContact.IsGroup) return;
            var groupEmail = NormalizeEmail(groupContact.SmtpAddress);
            var memberEmail = NormalizeEmail(PreferEmail(memberContact.SmtpAddress, member.SmtpAddress, member.RawAddress));
            if (string.IsNullOrWhiteSpace(groupEmail) || string.IsNullOrWhiteSpace(memberEmail)) return;

            memberContact.AddMemberOfGroup(groupEmail);
            if (!member.IsGroup) return;
            memberContact.MarkAsGroup();
            groupContact.AddMemberGroup(memberEmail);
        }

        private static AddressBookContactDto MergeCachedAddressBookContact(AddressBookContactDto current, AddressBookContactDto next)
        {
            var merged = new AddressBookAccumulator();
            merged.MergeAddressBookContact(current);
            merged.MergeAddressBookContact(next);
            var dto = merged.ToDto(new HashSet<string>(StringComparer.OrdinalIgnoreCase));
            dto.GroupMembersLoaded = current.GroupMembersLoaded || next.GroupMembersLoaded;
            dto.GroupMembersLoading = current.GroupMembersLoading || next.GroupMembersLoading;
            dto.GroupMembersRequestId = PreferLonger(current.GroupMembersRequestId, next.GroupMembersRequestId);
            dto.GroupMembersUpdatedAt = new[] { current.GroupMembersUpdatedAt, next.GroupMembersUpdatedAt }.Max();
            return dto;
        }
    }
}

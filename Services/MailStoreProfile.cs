using SmartOffice.Hub.Contracts;

namespace SmartOffice.Hub.Services
{
    public partial class MailStore
    {
        public OutlookProfileDto GetOutlookProfile()
        {
            lock (_lock)
            {
                var contacts = BuildAddressBookContacts();
                var stores = CloneStores(_stores);
                var selfEmails = InferSelfAddresses();
                var self = contacts
                    .Where(contact => contact.IsLikelySelf || selfEmails.Contains(NormalizeEmail(contact.SmtpAddress)))
                    .OrderByDescending(contact => contact.RelationScore)
                    .FirstOrDefault();
                var primaryStore = stores.FirstOrDefault(store => StoreKind(store) == "ost") ?? stores.FirstOrDefault();
                var seedGroups = contacts
                    .Where(contact => contact.IsGroup && (contact.IsRelatedToSelf || contact.Sources.Contains("mail_recipient") || contact.RecipientCount > 0))
                    .OrderByDescending(contact => contact.IsRelatedToSelf)
                    .ThenByDescending(contact => contact.RelationScore)
                    .ThenBy(contact => contact.DisplayName)
                    .ToList();
                var groups = ExpandProfileGroups(seedGroups, contacts)
                    .Take(48)
                    .Select(CloneAddressBookContact)
                    .ToList();
                var groupMembers = groups
                    .SelectMany(group => group.MemberSmtpAddresses.Select(email => (Email: NormalizeEmail(email), Group: group.SmtpAddress)))
                    .Where(item => !string.IsNullOrWhiteSpace(item.Email))
                    .ToList();
                var contactsByKey = contacts
                    .SelectMany(contact => ContactKeys(contact).Select(key => (Key: key, Contact: contact)))
                    .Where(item => !string.IsNullOrWhiteSpace(item.Key))
                    .GroupBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.First().Contact, StringComparer.OrdinalIgnoreCase);
                var sameGroupPeople = groupMembers
                    .Select(item => contactsByKey.TryGetValue(item.Email, out var contact) ? CloneAddressBookContact(contact) : ContactShell(item.Email, false))
                    .Where(contact => !contact.IsGroup && !contact.IsLikelySelf)
                    .GroupBy(contact => AddressBookContactKey(contact), StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.First())
                    .OrderByDescending(contact => contact.IsRelatedToSelf)
                    .ThenByDescending(contact => contact.RelationScore)
                    .ThenBy(contact => contact.DisplayName)
                    .Take(24)
                    .ToList();

                return new OutlookProfileDto
                {
                    State = "ready",
                    Message = $"Profile is ready from {_mails.Count} loaded mail item(s).",
                    MailboxName = primaryStore?.DisplayName ?? self?.DisplayName ?? string.Empty,
                    SmtpAddress = self?.SmtpAddress ?? StoreEmail(primaryStore) ?? string.Empty,
                    SelfContact = self is null ? null : CloneAddressBookContact(self),
                    GroupTree = BuildProfileGroupTree(groups),
                    Groups = groups,
                    SameGroupPeople = sameGroupPeople,
                    Stores = stores,
                    OstStores = stores.Where(store => StoreKind(store) == "ost").ToList(),
                    PstStores = stores.Where(store => StoreKind(store) == "pst").ToList(),
                    OtherStores = stores.Where(store => StoreKind(store) != "ost" && StoreKind(store) != "pst").ToList(),
                    MailStats = new OutlookProfileMailStatsDto
                    {
                        LoadedCount = _mails.Count,
                        UnreadCount = _mails.Count(mail => !mail.IsRead),
                        AttachmentMailCount = _mails.Count(mail => mail.AttachmentCount > 0),
                    },
                };
            }
        }

        private static string StoreKind(OutlookStoreDto store) => (store.StoreKind ?? string.Empty).Trim().ToLowerInvariant();

        private static List<AddressBookContactDto> ExpandProfileGroups(List<AddressBookContactDto> seedGroups, List<AddressBookContactDto> contacts)
        {
            var contactsByKey = contacts
                .SelectMany(contact => ContactKeys(contact).Select(key => (Key: key, Contact: contact)))
                .Where(item => !string.IsNullOrWhiteSpace(item.Key))
                .GroupBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().Contact, StringComparer.OrdinalIgnoreCase);
            var expanded = new List<AddressBookContactDto>();
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var queue = new Queue<AddressBookContactDto>(seedGroups);

            while (queue.Count > 0)
            {
                var group = queue.Dequeue();
                var key = AddressBookContactKey(group);
                if (string.IsNullOrWhiteSpace(key) || !visited.Add(key)) continue;
                expanded.Add(group);
                foreach (var childKey in group.MemberGroupSmtpAddresses.Select(NormalizeEmail).Where(childKey => !string.IsNullOrWhiteSpace(childKey)))
                {
                    if (contactsByKey.TryGetValue(childKey, out var child) && child.IsGroup) queue.Enqueue(child);
                }
            }

            return expanded
                .OrderByDescending(contact => contact.IsRelatedToSelf)
                .ThenByDescending(contact => contact.RelationScore)
                .ThenBy(contact => contact.DisplayName)
                .ToList();
        }

        private static List<OutlookProfileGroupNodeDto> BuildProfileGroupTree(List<AddressBookContactDto> groups)
        {
            var groupsByKey = groups
                .SelectMany(group => ContactKeys(group).Select(key => (Key: key, Group: group)))
                .Where(item => !string.IsNullOrWhiteSpace(item.Key))
                .GroupBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().Group, StringComparer.OrdinalIgnoreCase);
            var childKeys = groups
                .SelectMany(group => group.MemberGroupSmtpAddresses)
                .Select(NormalizeEmail)
                .Where(key => groupsByKey.ContainsKey(key))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var roots = groups
                .Where(group => !childKeys.Contains(NormalizeEmail(group.SmtpAddress)))
                .DefaultIfEmpty()
                .Where(group => group is not null)
                .Cast<AddressBookContactDto>()
                .ToList();
            return roots.Select(root => BuildProfileGroupNode(root, groupsByKey, new HashSet<string>(StringComparer.OrdinalIgnoreCase))).ToList();
        }

        private static OutlookProfileGroupNodeDto BuildProfileGroupNode(AddressBookContactDto group, Dictionary<string, AddressBookContactDto> groupsByKey, HashSet<string> path)
        {
            var key = NormalizeEmail(group.SmtpAddress);
            if (!string.IsNullOrWhiteSpace(key)) path.Add(key);
            var childGroups = group.MemberGroupSmtpAddresses
                .Select(NormalizeEmail)
                .Where(childKey => !string.IsNullOrWhiteSpace(childKey) && !path.Contains(childKey))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(childKey => groupsByKey.TryGetValue(childKey, out var child) ? child : ContactShell(childKey, true));
            return new OutlookProfileGroupNodeDto
            {
                Contact = CloneAddressBookContact(group),
                Children = childGroups.Select(child => BuildProfileGroupNode(child, groupsByKey, new HashSet<string>(path, StringComparer.OrdinalIgnoreCase))).ToList(),
            };
        }

        private static string? StoreEmail(OutlookStoreDto? store)
        {
            if (store is null) return null;
            return ExtractEmailLikeValues(store.AccountSmtpAddress)
                .Concat(ExtractEmailLikeValues(store.DisplayName))
                .Concat(ExtractEmailLikeValues(store.StoreFilePath))
                .FirstOrDefault();
        }
    }
}

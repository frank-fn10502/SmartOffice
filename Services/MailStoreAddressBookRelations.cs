namespace SmartOffice.Hub.Services
{
    public partial class MailStore
    {
        public AddressBookRelationLookupResponse GetAddressBookRelationLookup(AddressBookRelationLookupRequest request)
        {
            request ??= new AddressBookRelationLookupRequest();
            lock (_lock)
            {
                var take = Math.Clamp(request.Take <= 0 ? 50 : request.Take, 1, 500);
                var contacts = BuildAddressBookContacts();
                var target = FindRelationLookupTarget(contacts, request);
                var matches = target is null ? FindRelationLookupMatches(contacts, request, take) : new List<AddressBookContactDto>();

                if (target is null)
                {
                    return new AddressBookRelationLookupResponse
                    {
                        Query = RelationLookupQuery(request),
                        TargetKind = NormalizeRelationTargetKind(request.TargetKind),
                        State = matches.Count == 0 ? "not_found" : "ambiguous",
                        Message = matches.Count == 0 ? "No matching address book entry was found." : "Multiple matching address book entries were found.",
                        Matches = matches.Select(CloneAddressBookContact).ToList(),
                    };
                }

                var contactsByKey = contacts
                    .SelectMany(contact => ContactKeys(contact).Select(key => (Key: key, Contact: contact)))
                    .Where(item => !string.IsNullOrWhiteSpace(item.Key))
                    .GroupBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.First().Contact, StringComparer.OrdinalIgnoreCase);
                var targetKeys = ContactKeys(target).ToHashSet(StringComparer.OrdinalIgnoreCase);
                var containingGroups = contacts
                    .Where(contact => contact.IsGroup && GroupContainsAnyKey(contact, targetKeys))
                    .OrderBy(contact => contact.DisplayName)
                    .ThenBy(contact => contact.SmtpAddress)
                    .Take(take)
                    .Select(CloneAddressBookContact)
                    .ToList();
                var memberOfGroups = target.MemberOfGroupSmtpAddresses
                    .Select(NormalizeEmail)
                    .Where(key => !string.IsNullOrWhiteSpace(key))
                    .Concat(containingGroups.Select(group => NormalizeEmail(group.SmtpAddress)))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Select(key => contactsByKey.TryGetValue(key, out var group) ? CloneAddressBookContact(group) : ContactShell(key, true))
                    .Take(take)
                    .ToList();
                var members = target.IsGroup
                    ? ResolveRelationContacts(target.MemberSmtpAddresses, contactsByKey, take)
                    : new List<AddressBookContactDto>();
                var memberGroups = target.IsGroup
                    ? ResolveRelationContacts(target.MemberGroupSmtpAddresses, contactsByKey, take)
                    : new List<AddressBookContactDto>();

                return new AddressBookRelationLookupResponse
                {
                    Query = RelationLookupQuery(request),
                    TargetKind = target.IsGroup ? "group" : "person",
                    State = "found",
                    Message = "Address book relation result is ready.",
                    Target = CloneAddressBookContact(target),
                    Members = members,
                    MemberGroups = memberGroups,
                    MemberOfGroups = memberOfGroups,
                    ContainingGroups = containingGroups,
                    IsGroup = target.IsGroup,
                    IsLikelySelf = target.IsLikelySelf,
                    IsRelatedToSelf = target.IsRelatedToSelf,
                    GroupMembersLoaded = target.GroupMembersLoaded,
                    GroupMembersLoading = target.GroupMembersLoading,
                };
            }
        }

        private static AddressBookContactDto? FindRelationLookupTarget(List<AddressBookContactDto> contacts, AddressBookRelationLookupRequest request)
        {
            var kind = NormalizeRelationTargetKind(request.TargetKind);
            var exactKeys = RelationLookupExactKeys(request).ToList();
            foreach (var key in exactKeys)
            {
                var matches = contacts
                    .Where(contact => RelationKindMatches(contact, kind))
                    .Where(contact => ContactKeys(contact).Contains(key, StringComparer.OrdinalIgnoreCase))
                    .ToList();
                if (matches.Count == 1) return matches[0];
                if (matches.Count > 1)
                    return matches.OrderByDescending(contact => contact.RelationScore).ThenBy(contact => contact.DisplayName).First();
            }

            var query = NormalizeEmail(RelationLookupQuery(request));
            if (string.IsNullOrWhiteSpace(query)) return null;

            var queryMatches = contacts
                .Where(contact => RelationKindMatches(contact, kind))
                .Where(contact =>
                    string.Equals(NormalizeEmail(contact.DisplayName), query, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(NormalizeEmail(contact.SmtpAddress), query, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(NormalizeEmail(contact.RawAddress), query, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(NormalizeEmail(contact.Id), query, StringComparison.OrdinalIgnoreCase))
                .ToList();
            return queryMatches.Count == 1 ? queryMatches[0] : null;
        }

        private static List<AddressBookContactDto> FindRelationLookupMatches(List<AddressBookContactDto> contacts, AddressBookRelationLookupRequest request, int take)
        {
            var kind = NormalizeRelationTargetKind(request.TargetKind);
            var query = RelationLookupQuery(request);
            if (string.IsNullOrWhiteSpace(query)) return new List<AddressBookContactDto>();
            return contacts
                .Where(contact => RelationKindMatches(contact, kind))
                .Where(contact => ContactMatches(contact, query)
                    || ContactKeys(contact).Any(key => key.Contains(NormalizeEmail(query), StringComparison.OrdinalIgnoreCase)))
                .OrderByDescending(contact => contact.RelationScore)
                .ThenBy(contact => contact.DisplayName)
                .Take(take)
                .Select(CloneAddressBookContact)
                .ToList();
        }

        private static IEnumerable<string> RelationLookupExactKeys(AddressBookRelationLookupRequest request)
        {
            var values = new[]
            {
                request.GroupSmtpAddress,
                request.GroupId,
                request.SmtpAddress,
                request.Email,
                request.Id,
            };
            foreach (var value in values.Select(NormalizeEmail).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase))
                yield return value;
        }

        private static string RelationLookupQuery(AddressBookRelationLookupRequest request)
        {
            return new[]
            {
                request.GroupSmtpAddress,
                request.GroupId,
                request.SmtpAddress,
                request.Email,
                request.Id,
                request.DisplayName,
                request.Query,
            }.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? string.Empty;
        }

        private static string NormalizeRelationTargetKind(string kind)
        {
            kind = (kind ?? string.Empty).Trim().ToLowerInvariant();
            return kind is "group" or "person" ? kind : "auto";
        }

        private static bool RelationKindMatches(AddressBookContactDto contact, string kind)
        {
            return kind switch
            {
                "group" => contact.IsGroup,
                "person" => !contact.IsGroup,
                _ => true,
            };
        }

        private static bool GroupContainsAnyKey(AddressBookContactDto group, HashSet<string> targetKeys)
        {
            if (!group.IsGroup || targetKeys.Count == 0) return false;
            return group.MemberSmtpAddresses
                .Concat(group.MemberGroupSmtpAddresses)
                .Select(NormalizeEmail)
                .Any(key => targetKeys.Contains(key));
        }

        private static List<AddressBookContactDto> ResolveRelationContacts(
            IEnumerable<string> keys,
            Dictionary<string, AddressBookContactDto> contactsByKey,
            int take)
        {
            return keys
                .Select(NormalizeEmail)
                .Where(key => !string.IsNullOrWhiteSpace(key))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(key => contactsByKey.TryGetValue(key, out var contact) ? CloneAddressBookContact(contact) : ContactShell(key, false))
                .Take(take)
                .ToList();
        }

        private static AddressBookContactDto ContactShell(string smtpAddress, bool isGroup)
        {
            return new AddressBookContactDto
            {
                DisplayName = smtpAddress,
                SmtpAddress = smtpAddress,
                RawAddress = smtpAddress,
                IsGroup = isGroup,
                IsKnown = false,
            };
        }
    }
}

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
                var recipientRelevance = BuildRecipientRelevance(target, members, memberGroups);

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
                    RecipientRelevance = recipientRelevance,
                };
            }
        }

        private static AddressBookRecipientRelevanceDto BuildRecipientRelevance(
            AddressBookContactDto target,
            List<AddressBookContactDto> members,
            List<AddressBookContactDto> memberGroups)
        {
            if (target is null) return new AddressBookRecipientRelevanceDto();

            var personCount = members.Count(member => !member.IsGroup);
            var groupCount = memberGroups
                .Select(contact => ContactKeys(contact).FirstOrDefault(key => !string.IsNullOrWhiteSpace(key)) ?? string.Empty)
                .Where(key => !string.IsNullOrWhiteSpace(key))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count();
            var audienceSize = Math.Max(target.MemberCount, personCount + groupCount);
            var includesSelfDirectly = target.IsLikelySelf || members.Any(member => member.IsLikelySelf);
            var includesSelfThroughDirectGroup = memberGroups.Any(group => group.IsRelatedToSelf);
            var includesSelf = includesSelfDirectly || target.IsRelatedToSelf || includesSelfThroughDirectGroup;
            var routeDepth = RecipientRouteDepth(target, includesSelfDirectly, includesSelfThroughDirectGroup, includesSelf);
            var reasons = new List<string>();

            var score = target.IsGroup
                ? RecipientRouteDepthScore(routeDepth)
                : target.IsLikelySelf ? 100 : 0;

            if (target.IsGroup && score > 0)
            {
                score = (int)Math.Round(score * AudienceBreadthFactor(audienceSize));
                score += Math.Min(15, target.RelationScore / 5);
            }

            score = Math.Clamp(score, 0, 100);
            if (includesSelfDirectly) reasons.Add("You are a direct recipient or direct member.");
            else if (includesSelfThroughDirectGroup) reasons.Add("You are related through a direct nested group.");
            else if (includesSelf) reasons.Add("The recipient group is related to you through a known membership path.");
            if (audienceSize > 0) reasons.Add($"Audience size is about {audienceSize} direct entries.");
            if (groupCount > 0) reasons.Add($"It contains {groupCount} direct group entries.");
            if (target.MailCount > 0) reasons.Add($"This entry appears in {target.MailCount} mail interaction(s).");
            if (target.CalendarCount > 0) reasons.Add($"This entry appears in {target.CalendarCount} calendar interaction(s).");
            if (reasons.Count == 0) reasons.Add("No strong recipient-path evidence is available for this entry.");

            return new AddressBookRecipientRelevanceDto
            {
                Score = score,
                Level = RecipientRelevanceLevel(score),
                Summary = RecipientRelevanceSummary(target, score, audienceSize, includesSelf),
                RouteDepth = routeDepth,
                DirectPersonCount = personCount,
                DirectGroupCount = groupCount,
                AudienceSize = audienceSize,
                IncludesSelf = includesSelf,
                IncludesSelfDirectly = includesSelfDirectly,
                Reasons = reasons,
            };
        }

        private static int RecipientRouteDepth(
            AddressBookContactDto target,
            bool includesSelfDirectly,
            bool includesSelfThroughDirectGroup,
            bool includesSelf)
        {
            if (!target.IsGroup) return target.IsLikelySelf ? 0 : -1;
            if (includesSelfDirectly) return 1;
            if (includesSelfThroughDirectGroup) return 2;
            if (includesSelf) return 3;
            return -1;
        }

        private static int RecipientRouteDepthScore(int routeDepth)
        {
            if (routeDepth < 0) return 0;
            if (routeDepth <= 2) return 100 - routeDepth * 5;
            return (int)Math.Round(90 * Math.Exp(-0.4 * (routeDepth - 2)));
        }

        private static double AudienceBreadthFactor(int audienceSize)
        {
            if (audienceSize <= 10) return 1.0;
            if (audienceSize <= 50) return 0.92;
            if (audienceSize <= 200) return 0.82;
            return 0.72;
        }

        private static string RecipientRelevanceLevel(int score)
        {
            if (score >= 80) return "direct";
            if (score >= 60) return "strong";
            if (score >= 35) return "broad";
            if (score > 0) return "weak";
            return "unknown";
        }

        private static string RecipientRelevanceSummary(AddressBookContactDto target, int score, int audienceSize, bool includesSelf)
        {
            if (!target.IsGroup)
                return score >= 80 ? "This person is a direct recipient-path match." : "This person has limited recipient-path evidence.";
            if (!includesSelf)
                return "This group is not known to include you.";
            if (audienceSize >= 50)
                return "This recipient group appears broad; content signals may matter more than recipient path alone.";
            if (audienceSize >= 10)
                return "This recipient group appears team or department sized; it is related to you, but not necessarily personally actionable.";
            return "This recipient path is small or direct, so it is likely personally relevant.";
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

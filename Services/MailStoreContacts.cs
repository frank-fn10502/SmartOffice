using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public partial class MailStore
    {
        private List<AddressBookContactDto> _addressBookContacts = new();
        private readonly Dictionary<string, AddressBookGroupExpansionState> _addressBookGroupExpansions = new(StringComparer.OrdinalIgnoreCase);
        private List<AddressBookRootDto> _addressBookRoots = new();
        private readonly Dictionary<string, AddressBookListEntriesPageDto> _addressBookListEntriesByRequestId = new(StringComparer.OrdinalIgnoreCase);

        public void SetAddressBookContacts(List<AddressBookContactDto> contacts)
        {
            lock (_lock)
            {
                _addressBookContacts = contacts
                    .Where(contact => !string.IsNullOrWhiteSpace(contact.SmtpAddress) || !string.IsNullOrWhiteSpace(contact.DisplayName))
                    .Select(CloneAddressBookContact)
                    .ToList();
            }
        }

        public void ApplyAddressBookBatch(AddressBookBatchDto batch)
        {
            batch ??= new AddressBookBatchDto();
            lock (_lock)
            {
                if (batch.Reset)
                {
                    _addressBookContacts = new List<AddressBookContactDto>();
                    _addressBookGroupExpansions.Clear();
                }
                foreach (var contact in batch.Contacts.Where(contact =>
                    !string.IsNullOrWhiteSpace(contact.SmtpAddress) || !string.IsNullOrWhiteSpace(contact.DisplayName)))
                {
                    var key = AddressBookContactKey(contact);
                    if (string.IsNullOrWhiteSpace(key)) continue;
                    var index = _addressBookContacts.FindIndex(existing =>
                        string.Equals(AddressBookContactKey(existing), key, StringComparison.OrdinalIgnoreCase));
                    var clone = CloneAddressBookContact(contact);
                    if (index >= 0) _addressBookContacts[index] = clone;
                    else _addressBookContacts.Add(clone);
                }
            }
        }

        public AddressBookGroupMembersResponse GetAddressBookGroupMembers(AddressBookGroupMembersRequest request)
        {
            lock (_lock)
            {
                return BuildGroupMembersResponse(GroupKey(request));
            }
        }

        public AddressBookGroupMembersResponse BeginAddressBookGroupExpansion(AddressBookGroupMembersRequest request, string requestId)
        {
            lock (_lock)
            {
                var key = GroupKey(request);
                if (string.IsNullOrWhiteSpace(key))
                    return new AddressBookGroupMembersResponse { State = "failed", Message = "groupId or groupSmtpAddress is required." };

                if (_addressBookGroupExpansions.TryGetValue(key, out var current))
                {
                    if (!request.ForceRefresh && current.State == "completed")
                        return BuildGroupMembersResponse(key);
                    if (current.State == "loading")
                        return BuildGroupMembersResponse(key);
                }

                _addressBookGroupExpansions[key] = new AddressBookGroupExpansionState
                {
                    GroupKey = key,
                    GroupSmtpAddress = request.GroupSmtpAddress,
                    RequestId = requestId,
                    State = "loading",
                };
                return BuildGroupMembersResponse(key);
            }
        }

        public void ApplyAddressBookGroupMembersBatch(AddressBookGroupMembersBatchDto batch)
        {
            batch ??= new AddressBookGroupMembersBatchDto();
            lock (_lock)
            {
                var key = GroupKey(batch.GroupSmtpAddress, batch.GroupId);
                if (string.IsNullOrWhiteSpace(key)) return;

                if (!_addressBookGroupExpansions.TryGetValue(key, out var state))
                {
                    state = new AddressBookGroupExpansionState { GroupKey = key };
                    _addressBookGroupExpansions[key] = state;
                }

                state.GroupSmtpAddress = PreferNonEmpty(state.GroupSmtpAddress, batch.GroupSmtpAddress);
                state.State = batch.IsFinal ? "completed" : "loading";
                state.TotalCount = Math.Max(state.TotalCount, batch.TotalCount);
                if (batch.Reset) state.Members.Clear();

                foreach (var member in batch.Members.Where(contact =>
                    !string.IsNullOrWhiteSpace(contact.SmtpAddress) || !string.IsNullOrWhiteSpace(contact.DisplayName)))
                {
                    var memberKey = AddressBookContactKey(member);
                    if (string.IsNullOrWhiteSpace(memberKey)) continue;
                    state.Members[memberKey] = CloneAddressBookContact(member);
                    UpsertAddressBookContact(member);
                }

                state.TotalCount = Math.Max(state.TotalCount, state.Members.Count);
                if (batch.IsFinal) state.UpdatedAt = DateTime.UtcNow;
                MergeGroupExpansionIntoContact(key, state);
            }
        }

        public List<AddressBookContactDto> GetAddressBookContacts(string query = "", int take = 200)
        {
            lock (_lock)
            {
                var contacts = BuildAddressBookContacts();
                var normalizedQuery = (query ?? string.Empty).Trim();
                if (!string.IsNullOrWhiteSpace(normalizedQuery))
                {
                    contacts = contacts
                        .Where(contact => ContactMatches(contact, normalizedQuery))
                        .ToList();
                }

                var takeLimit = take <= 0 ? int.MaxValue : Math.Clamp(take, 1, 5000);
                return contacts
                    .OrderByDescending(AddressBookContactPriority)
                    .ThenByDescending(contact => contact.RelationScore)
                    .ThenByDescending(contact => contact.LastSeen ?? DateTime.MinValue)
                    .ThenBy(contact => contact.DisplayName)
                    .Take(takeLimit)
                    .Select(CloneAddressBookContact)
                    .ToList();
            }
        }

        public bool IsAddressBookGroupExpansionCompleted(AddressBookGroupMembersRequest? request)
        {
            if (request is null) return false;
            lock (_lock)
            {
                return _addressBookGroupExpansions.TryGetValue(GroupKey(request), out var state)
                    && state.State == "completed";
            }
        }

        public AddressBookContactDto? FindAddressBookContact(string email)
        {
            if (string.IsNullOrWhiteSpace(email)) return null;
            lock (_lock)
            {
                var normalized = NormalizeEmail(email);
                return BuildAddressBookContacts()
                    .FirstOrDefault(contact => string.Equals(NormalizeEmail(contact.SmtpAddress), normalized, StringComparison.OrdinalIgnoreCase)) is { } contact
                        ? CloneAddressBookContact(contact)
                        : null;
            }
        }

        public List<AddressBookMergeSuggestionDto> GetAddressBookMergeSuggestions(IEnumerable<string> recipients)
        {
            var keys = recipients
                .Select(NormalizeEmail)
                .Where(key => !string.IsNullOrWhiteSpace(key))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (keys.Count < 2) return new List<AddressBookMergeSuggestionDto>();

            lock (_lock)
            {
                var contacts = BuildAddressBookContacts();
                var contactsByKey = contacts
                    .SelectMany(contact => ContactKeys(contact).Select(key => (Key: key, Contact: contact)))
                    .GroupBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.First().Contact, StringComparer.OrdinalIgnoreCase);

                var selected = keys
                    .Where(contactsByKey.ContainsKey)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                return contacts
                    .Where(contact => contact.IsGroup && selected.Contains(NormalizeEmail(contact.SmtpAddress)))
                    .Select(group => BuildMergeSuggestion(group, selected, contactsByKey))
                    .Where(suggestion => suggestion.CoveredContacts.Count > 0)
                    .ToList();
            }
        }

        private List<AddressBookContactDto> BuildAddressBookContacts()
        {
            var contacts = new Dictionary<string, AddressBookAccumulator>(StringComparer.OrdinalIgnoreCase);
            var selfAddresses = InferSelfAddresses();

            foreach (var contact in _addressBookContacts)
                AddCachedAddressBookContact(contacts, contact);

            AddCachedGroupMemberships(contacts, _addressBookContacts);

            var knownMails = _mails
                .Concat(_mailSearchResults)
                .Concat(_folderMailResults)
                .Concat(_mailSearchResultsBySearchId.Values.SelectMany(items => items))
                .Concat(_folderMailResultsByRequestId.Values.SelectMany(items => items))
                .GroupBy(mail => $"{mail.Id}\n{mail.FolderPath}", StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First());

            foreach (var mail in knownMails)
            {
                AddMailRecipient(contacts, mail.Sender, "sender", mail);
                AddMailRecipients(contacts, mail.ToRecipients, "to", mail);
                AddMailRecipients(contacts, mail.CcRecipients, "cc", mail);
                AddMailRecipients(contacts, mail.BccRecipients, "bcc", mail);
            }

            foreach (var calendarEvent in _calendarEvents
                .GroupBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First()))
            {
                AddCalendarRecipient(contacts, calendarEvent.Organizer, "organizer", calendarEvent);
                AddCalendarRecipients(contacts, calendarEvent.RequiredAttendees, "attendee", calendarEvent);
            }

            var result = contacts.Values
                .Select(item => item.ToDto(selfAddresses))
                .Select(ApplyGroupExpansionStatus)
                .Where(contact => !string.IsNullOrWhiteSpace(contact.SmtpAddress) || !string.IsNullOrWhiteSpace(contact.DisplayName))
                .ToList();
            ApplySelfGroupRelations(result);
            return result;
        }

        private void UpsertAddressBookContact(AddressBookContactDto contact)
        {
            var key = AddressBookContactKey(contact);
            if (string.IsNullOrWhiteSpace(key)) return;
            var index = _addressBookContacts.FindIndex(existing =>
                string.Equals(AddressBookContactKey(existing), key, StringComparison.OrdinalIgnoreCase));
            var clone = CloneAddressBookContact(contact);
            if (index >= 0) _addressBookContacts[index] = clone;
            else _addressBookContacts.Add(clone);
        }

        private void MergeGroupExpansionIntoContact(string groupKey, AddressBookGroupExpansionState state)
        {
            var index = _addressBookContacts.FindIndex(existing =>
                string.Equals(AddressBookContactKey(existing), groupKey, StringComparison.OrdinalIgnoreCase)
                || string.Equals(NormalizeEmail(existing.SmtpAddress), groupKey, StringComparison.OrdinalIgnoreCase));
            if (index < 0) return;

            var group = CloneAddressBookContact(_addressBookContacts[index]);
            group.IsGroup = true;
            group.MemberCount = Math.Max(group.MemberCount, state.TotalCount);
            group.MemberSmtpAddresses = state.Members.Values
                .Select(member => member.SmtpAddress)
                .Where(email => !string.IsNullOrWhiteSpace(email))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(50)
                .ToList();
            group.MemberGroupSmtpAddresses = state.Members.Values
                .Where(member => member.IsGroup)
                .Select(member => member.SmtpAddress)
                .Where(email => !string.IsNullOrWhiteSpace(email))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(50)
                .ToList();
            _addressBookContacts[index] = ApplyGroupExpansionStatus(group);
        }

        private static void AddCachedAddressBookContact(
            Dictionary<string, AddressBookAccumulator> contacts,
            AddressBookContactDto source)
        {
            var key = NormalizeEmail(source.SmtpAddress);
            if (string.IsNullOrWhiteSpace(key))
                key = source.DisplayName.Trim().ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(key)) return;

            if (!contacts.TryGetValue(key, out var contact))
            {
                contact = new AddressBookAccumulator();
                contacts[key] = contact;
            }

            contact.MergeAddressBookContact(source);
        }

        private static void AddCachedGroupMemberships(
            Dictionary<string, AddressBookAccumulator> contacts,
            IEnumerable<AddressBookContactDto> sources)
        {
            var groupsByEmail = sources
                .Where(source => source.IsGroup && !string.IsNullOrWhiteSpace(source.SmtpAddress))
                .GroupBy(source => NormalizeEmail(source.SmtpAddress), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            foreach (var group in groupsByEmail.Values)
            {
                var groupEmail = NormalizeEmail(group.SmtpAddress);
                if (string.IsNullOrWhiteSpace(groupEmail)) continue;

                foreach (var memberEmail in group.MemberSmtpAddresses.Select(NormalizeEmail).Where(email => !string.IsNullOrWhiteSpace(email)))
                {
                    if (!contacts.TryGetValue(memberEmail, out var member))
                    {
                        member = new AddressBookAccumulator { SmtpAddress = memberEmail, RawAddress = memberEmail };
                        contacts[memberEmail] = member;
                    }

                    member.AddMemberOfGroup(groupEmail);
                    if (groupsByEmail.ContainsKey(memberEmail))
                    {
                        member.MarkAsGroup();
                        member.AddMemberOfGroup(groupEmail);
                        if (contacts.TryGetValue(groupEmail, out var groupContact))
                            groupContact.AddMemberGroup(memberEmail);
                    }
                }
            }
        }

        private static void AddMailRecipients(
            Dictionary<string, AddressBookAccumulator> contacts,
            IEnumerable<OutlookRecipientDto> recipients,
            string relationKind,
            MailItemDto mail)
        {
            foreach (var recipient in recipients)
                AddMailRecipient(contacts, recipient, relationKind, mail);
        }

        private static void AddMailRecipient(
            Dictionary<string, AddressBookAccumulator> contacts,
            OutlookRecipientDto recipient,
            string relationKind,
            MailItemDto mail)
        {
            var contact = GetOrCreateContact(contacts, recipient);
            if (contact is null) return;

            contact.AddMail(relationKind, mail);
            foreach (var member in recipient.Members)
            {
                var memberContact = GetOrCreateContact(contacts, member);
                memberContact?.AddMail("group_member", mail);
            }
        }

        private static void AddCalendarRecipients(
            Dictionary<string, AddressBookAccumulator> contacts,
            IEnumerable<OutlookRecipientDto> recipients,
            string relationKind,
            CalendarEventDto calendarEvent)
        {
            foreach (var recipient in recipients)
                AddCalendarRecipient(contacts, recipient, relationKind, calendarEvent);
        }

        private static void AddCalendarRecipient(
            Dictionary<string, AddressBookAccumulator> contacts,
            OutlookRecipientDto recipient,
            string relationKind,
            CalendarEventDto calendarEvent)
        {
            var contact = GetOrCreateContact(contacts, recipient);
            if (contact is null) return;

            contact.AddCalendar(relationKind, calendarEvent);
            foreach (var member in recipient.Members)
            {
                var memberContact = GetOrCreateContact(contacts, member);
                memberContact?.AddCalendar("group_member", calendarEvent);
            }
        }

        private static AddressBookAccumulator? GetOrCreateContact(
            Dictionary<string, AddressBookAccumulator> contacts,
            OutlookRecipientDto recipient)
        {
            var key = NormalizeEmail(recipient.SmtpAddress);
            if (string.IsNullOrWhiteSpace(key))
                key = NormalizeEmail(recipient.RawAddress);
            if (string.IsNullOrWhiteSpace(key))
                key = recipient.DisplayName.Trim().ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(key)) return null;

            if (!contacts.TryGetValue(key, out var contact))
            {
                contact = new AddressBookAccumulator();
                contacts[key] = contact;
            }

            contact.DisplayName = PreferLonger(contact.DisplayName, recipient.DisplayName);
            contact.SmtpAddress = PreferEmail(contact.SmtpAddress, recipient.SmtpAddress, recipient.RawAddress);
            if (recipient.IsGroup || recipient.Members.Count > 0)
                contact.MergeAddressBookContact(new AddressBookContactDto
                {
                    DisplayName = recipient.DisplayName,
                    SmtpAddress = contact.SmtpAddress,
                    RawAddress = recipient.RawAddress,
                    EntryUserType = recipient.EntryUserType,
                    IsGroup = true,
                    MemberSmtpAddresses = recipient.Members.Select(member => PreferEmail(member.SmtpAddress, member.RawAddress)).Where(email => !string.IsNullOrWhiteSpace(email)).ToList(),
                    MemberGroupSmtpAddresses = recipient.Members.Where(member => member.IsGroup).Select(member => PreferEmail(member.SmtpAddress, member.RawAddress)).Where(email => !string.IsNullOrWhiteSpace(email)).ToList(),
                    Source = "mail_recipient",
                });
            return contact;
        }

        private static bool ContactMatches(AddressBookContactDto contact, string query)
        {
            return Contains(contact.DisplayName, query)
                || Contains(contact.SmtpAddress, query)
                || Contains(contact.Domain, query);
        }

        private static AddressBookMergeSuggestionDto BuildMergeSuggestion(
            AddressBookContactDto group,
            HashSet<string> selected,
            Dictionary<string, AddressBookContactDto> contactsByKey)
        {
            var groupKey = NormalizeEmail(group.SmtpAddress);
            var coveredKeys = group.MemberSmtpAddresses
                .Concat(group.MemberGroupSmtpAddresses)
                .Select(NormalizeEmail)
                .Where(key => !string.IsNullOrWhiteSpace(key) && selected.Contains(key) && key != groupKey)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var coveredContacts = coveredKeys
                .Select(key => contactsByKey.TryGetValue(key, out var contact) ? CloneAddressBookContact(contact) : null)
                .Where(contact => contact is not null)
                .Cast<AddressBookContactDto>()
                .ToList();

            return new AddressBookMergeSuggestionDto
            {
                GroupSmtpAddress = group.SmtpAddress,
                GroupDisplayName = string.IsNullOrWhiteSpace(group.DisplayName) ? group.SmtpAddress : group.DisplayName,
                CoveredContacts = coveredContacts,
                CoveredRecipientKeys = coveredKeys,
                Message = coveredContacts.Count == 0
                    ? string.Empty
                    : $"{(string.IsNullOrWhiteSpace(group.DisplayName) ? group.SmtpAddress : group.DisplayName)} 已包含 {string.Join(", ", coveredContacts.Select(ContactLabel))}",
            };
        }

        private static string ContactLabel(AddressBookContactDto contact)
        {
            return string.IsNullOrWhiteSpace(contact.DisplayName) ? contact.SmtpAddress : contact.DisplayName;
        }

        private static IEnumerable<string> ContactKeys(AddressBookContactDto contact)
        {
            if (!string.IsNullOrWhiteSpace(contact.SmtpAddress)) yield return NormalizeEmail(contact.SmtpAddress);
            if (!string.IsNullOrWhiteSpace(contact.RawAddress)) yield return NormalizeEmail(contact.RawAddress);
            if (!string.IsNullOrWhiteSpace(contact.DisplayName)) yield return NormalizeEmail(contact.DisplayName);
            if (!string.IsNullOrWhiteSpace(contact.Id)) yield return NormalizeEmail(contact.Id);
        }

        private static string AddressBookContactKey(AddressBookContactDto contact)
        {
            return ContactKeys(contact).FirstOrDefault(key => !string.IsNullOrWhiteSpace(key)) ?? string.Empty;
        }

        private static bool Contains(string value, string query)
        {
            return value.Contains(query, StringComparison.OrdinalIgnoreCase);
        }

        private static string PreferEmail(string current, params string[] candidates)
        {
            if (!string.IsNullOrWhiteSpace(current) && current.Contains('@')) return current;
            return candidates.FirstOrDefault(candidate => !string.IsNullOrWhiteSpace(candidate) && candidate.Contains('@'))?.Trim()
                ?? current
                ?? string.Empty;
        }

        private static string PreferLonger(string current, string candidate)
        {
            current = current?.Trim() ?? string.Empty;
            candidate = candidate?.Trim() ?? string.Empty;
            return candidate.Length > current.Length ? candidate : current;
        }

        private static string NormalizeEmail(string email)
        {
            return (email ?? string.Empty).Trim().Trim('<', '>').ToLowerInvariant();
        }

        private static AddressBookContactDto CloneAddressBookContact(AddressBookContactDto contact)
        {
            return new AddressBookContactDto
            {
                Id = contact.Id,
                DisplayName = contact.DisplayName,
                SmtpAddress = contact.SmtpAddress,
                RawAddress = contact.RawAddress,
                AddressType = contact.AddressType,
                EntryUserType = contact.EntryUserType,
                Source = contact.Source,
                CompanyName = contact.CompanyName,
                JobTitle = contact.JobTitle,
                Department = contact.Department,
                OfficeLocation = contact.OfficeLocation,
                BusinessTelephoneNumber = contact.BusinessTelephoneNumber,
                MobileTelephoneNumber = contact.MobileTelephoneNumber,
                Domain = contact.Domain,
                IsKnown = contact.IsKnown,
                IsLikelySelf = contact.IsLikelySelf,
                IsRelatedToSelf = contact.IsRelatedToSelf,
                IsGroup = contact.IsGroup,
                MemberCount = contact.MemberCount,
                GroupMembersLoaded = contact.GroupMembersLoaded,
                GroupMembersLoading = contact.GroupMembersLoading,
                GroupMembersRequestId = contact.GroupMembersRequestId,
                GroupMembersUpdatedAt = contact.GroupMembersUpdatedAt,
                RelationScore = contact.RelationScore,
                MailCount = contact.MailCount,
                CalendarCount = contact.CalendarCount,
                SenderCount = contact.SenderCount,
                RecipientCount = contact.RecipientCount,
                OrganizerCount = contact.OrganizerCount,
                AttendeeCount = contact.AttendeeCount,
                GroupMemberCount = contact.GroupMemberCount,
                FirstSeen = contact.FirstSeen,
                LastSeen = contact.LastSeen,
                RelationKinds = new List<string>(contact.RelationKinds),
                Sources = new List<string>(contact.Sources),
                MemberSmtpAddresses = new List<string>(contact.MemberSmtpAddresses),
                MemberGroupSmtpAddresses = new List<string>(contact.MemberGroupSmtpAddresses),
                MemberOfGroupSmtpAddresses = new List<string>(contact.MemberOfGroupSmtpAddresses),
                FolderPaths = new List<string>(contact.FolderPaths),
                RecentMailIds = new List<string>(contact.RecentMailIds),
                SampleSubjects = new List<string>(contact.SampleSubjects),
            };
        }

        private AddressBookContactDto ApplyGroupExpansionStatus(AddressBookContactDto contact)
        {
            if (!contact.IsGroup) return contact;
            var key = GroupKey(contact.SmtpAddress, contact.Id);
            if (string.IsNullOrWhiteSpace(key) || !_addressBookGroupExpansions.TryGetValue(key, out var state))
                return contact;

            contact.GroupMembersLoaded = state.State == "completed";
            contact.GroupMembersLoading = state.State == "loading";
            contact.GroupMembersRequestId = state.RequestId;
            contact.GroupMembersUpdatedAt = state.UpdatedAt;
            contact.MemberCount = Math.Max(contact.MemberCount, state.TotalCount);
            return contact;
        }

        private AddressBookGroupMembersResponse BuildGroupMembersResponse(string groupKey)
        {
            if (string.IsNullOrWhiteSpace(groupKey))
                return new AddressBookGroupMembersResponse { State = "failed", Message = "groupId or groupSmtpAddress is required." };

            if (!_addressBookGroupExpansions.TryGetValue(groupKey, out var state))
            {
                var group = _addressBookContacts.FirstOrDefault(contact =>
                    string.Equals(AddressBookContactKey(contact), groupKey, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(NormalizeEmail(contact.SmtpAddress), groupKey, StringComparison.OrdinalIgnoreCase));
                return new AddressBookGroupMembersResponse
                {
                    State = "not_loaded",
                    Message = "Group members have not been expanded.",
                    GroupKey = groupKey,
                    GroupSmtpAddress = group?.SmtpAddress ?? groupKey,
                    TotalCount = group?.MemberCount ?? 0,
                };
            }

            var contactsByKey = BuildAddressBookContacts()
                .SelectMany(contact => ContactKeys(contact).Select(key => (Key: key, Contact: contact)))
                .Where(item => !string.IsNullOrWhiteSpace(item.Key))
                .GroupBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().Contact, StringComparer.OrdinalIgnoreCase);

            return new AddressBookGroupMembersResponse
            {
                State = state.State,
                Message = state.State == "completed" ? "Group members loaded from Hub cache." : "Group members are loading.",
                GroupKey = state.GroupKey,
                GroupSmtpAddress = state.GroupSmtpAddress,
                RequestId = state.RequestId,
                TotalCount = state.TotalCount,
                UpdatedAt = state.UpdatedAt,
                Members = state.Members.Values
                    .OrderBy(member => member.DisplayName)
                    .ThenBy(member => member.SmtpAddress)
                    .Select(member => TryGetEnrichedAddressBookContact(member, contactsByKey))
                    .ToList(),
            };
        }

        private static string GroupKey(AddressBookGroupMembersRequest request)
        {
            return request is null ? string.Empty : GroupKey(request.GroupSmtpAddress, request.GroupId);
        }

        private static string GroupKey(string groupSmtpAddress, string groupId)
        {
            var smtp = NormalizeEmail(groupSmtpAddress);
            return !string.IsNullOrWhiteSpace(smtp) ? smtp : (groupId ?? string.Empty).Trim().ToLowerInvariant();
        }

        private static string PreferNonEmpty(string current, string candidate)
        {
            return string.IsNullOrWhiteSpace(current) ? candidate ?? string.Empty : current;
        }

        private sealed class AddressBookGroupExpansionState
        {
            public string GroupKey { get; set; } = string.Empty;
            public string GroupSmtpAddress { get; set; } = string.Empty;
            public string RequestId { get; set; } = string.Empty;
            public string State { get; set; } = "not_loaded";
            public int TotalCount { get; set; }
            public DateTime? UpdatedAt { get; set; }
            public Dictionary<string, AddressBookContactDto> Members { get; } = new(StringComparer.OrdinalIgnoreCase);
        }

        private sealed class AddressBookAccumulator
        {
            private readonly HashSet<string> _relationKinds = new(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _sources = new(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _memberSmtpAddresses = new(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _memberGroupSmtpAddresses = new(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _memberOfGroupSmtpAddresses = new(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _folderPaths = new(StringComparer.OrdinalIgnoreCase);
            private readonly List<(string Id, DateTime SeenAt)> _recentMailIds = new();
            private readonly List<(string Subject, DateTime SeenAt)> _sampleSubjects = new();

            public string DisplayName { get; set; } = string.Empty;
            public string Id { get; set; } = string.Empty;
            public string SmtpAddress { get; set; } = string.Empty;
            public string RawAddress { get; set; } = string.Empty;
            public string AddressType { get; set; } = string.Empty;
            public string EntryUserType { get; set; } = string.Empty;
            public string Source { get; set; } = string.Empty;
            public string CompanyName { get; set; } = string.Empty;
            public string JobTitle { get; set; } = string.Empty;
            public string Department { get; set; } = string.Empty;
            public string OfficeLocation { get; set; } = string.Empty;
            public string BusinessTelephoneNumber { get; set; } = string.Empty;
            public string MobileTelephoneNumber { get; set; } = string.Empty;
            public bool IsGroup { get; private set; }
            public int MemberCount { get; private set; }
            public int MailCount { get; private set; }
            public int CalendarCount { get; private set; }
            public int SenderCount { get; private set; }
            public int RecipientCount { get; private set; }
            public int OrganizerCount { get; private set; }
            public int AttendeeCount { get; private set; }
            public int GroupMemberCount { get; private set; }
            public DateTime? FirstSeen { get; private set; }
            public DateTime? LastSeen { get; private set; }

            public void AddMail(string relationKind, MailItemDto mail)
            {
                MailCount++;
                AddCommon(relationKind, "mail", mail.ReceivedTime);
                if (relationKind == "sender") SenderCount++;
                else if (relationKind == "group_member") GroupMemberCount++;
                else RecipientCount++;

                if (!string.IsNullOrWhiteSpace(mail.FolderPath)) _folderPaths.Add(OutlookFolderPathMapper.ToApiPath(mail.FolderPath));
                if (!string.IsNullOrWhiteSpace(mail.Id)) _recentMailIds.Add((mail.Id, mail.ReceivedTime));
                if (!string.IsNullOrWhiteSpace(mail.Subject)) _sampleSubjects.Add((mail.Subject, mail.ReceivedTime));
            }

            public void AddCalendar(string relationKind, CalendarEventDto calendarEvent)
            {
                CalendarCount++;
                AddCommon(relationKind, "calendar", calendarEvent.Start);
                if (relationKind == "organizer") OrganizerCount++;
                else if (relationKind == "group_member") GroupMemberCount++;
                else AttendeeCount++;
            }

            public void MergeAddressBookContact(AddressBookContactDto source)
            {
                Id = PreferLonger(Id, source.Id);
                DisplayName = PreferLonger(DisplayName, source.DisplayName);
                SmtpAddress = PreferEmail(SmtpAddress, source.SmtpAddress, source.RawAddress);
                RawAddress = PreferLonger(RawAddress, source.RawAddress);
                AddressType = PreferLonger(AddressType, source.AddressType);
                EntryUserType = PreferLonger(EntryUserType, source.EntryUserType);
                Source = PreferLonger(Source, source.Source);
                CompanyName = PreferLonger(CompanyName, source.CompanyName);
                JobTitle = PreferLonger(JobTitle, source.JobTitle);
                Department = PreferLonger(Department, source.Department);
                OfficeLocation = PreferLonger(OfficeLocation, source.OfficeLocation);
                BusinessTelephoneNumber = PreferLonger(BusinessTelephoneNumber, source.BusinessTelephoneNumber);
                MobileTelephoneNumber = PreferLonger(MobileTelephoneNumber, source.MobileTelephoneNumber);
                IsGroup = IsGroup || source.IsGroup || IsGroupEntryUserType(source.EntryUserType);
                MemberCount = Math.Max(MemberCount, source.MemberCount);
                foreach (var member in source.MemberSmtpAddresses.Where(member => !string.IsNullOrWhiteSpace(member)))
                    _memberSmtpAddresses.Add(member.Trim());
                foreach (var group in source.MemberGroupSmtpAddresses.Where(group => !string.IsNullOrWhiteSpace(group)))
                    _memberGroupSmtpAddresses.Add(group.Trim());
                foreach (var group in source.MemberOfGroupSmtpAddresses.Where(group => !string.IsNullOrWhiteSpace(group)))
                    _memberOfGroupSmtpAddresses.Add(group.Trim());
                _sources.Add(string.IsNullOrWhiteSpace(source.Source) ? "address_book" : source.Source);
                _relationKinds.Add("address_book");
            }

            public void AddMemberOfGroup(string groupSmtpAddress)
            {
                if (!string.IsNullOrWhiteSpace(groupSmtpAddress))
                    _memberOfGroupSmtpAddresses.Add(groupSmtpAddress.Trim());
            }

            public void AddMemberGroup(string groupSmtpAddress)
            {
                if (!string.IsNullOrWhiteSpace(groupSmtpAddress))
                    _memberGroupSmtpAddresses.Add(groupSmtpAddress.Trim());
            }

            public void MarkAsGroup()
            {
                IsGroup = true;
            }

            public AddressBookContactDto ToDto(HashSet<string> selfAddresses)
            {
                var normalizedEmail = NormalizeEmail(SmtpAddress);
                return new AddressBookContactDto
                {
                    Id = Id,
                    DisplayName = DisplayName,
                    SmtpAddress = SmtpAddress,
                    RawAddress = RawAddress,
                    AddressType = AddressType,
                    EntryUserType = EntryUserType,
                    Source = Source,
                    CompanyName = CompanyName,
                    JobTitle = JobTitle,
                    Department = Department,
                    OfficeLocation = OfficeLocation,
                    BusinessTelephoneNumber = BusinessTelephoneNumber,
                    MobileTelephoneNumber = MobileTelephoneNumber,
                    Domain = EmailDomain(SmtpAddress),
                    IsKnown = MailCount > 0 || CalendarCount > 0 || _sources.Count > 0,
                    IsLikelySelf = !string.IsNullOrWhiteSpace(normalizedEmail) && selfAddresses.Contains(normalizedEmail),
                    IsRelatedToSelf = IsGroup
                        ? _memberSmtpAddresses.Select(NormalizeEmail).Any(email => !string.IsNullOrWhiteSpace(email) && selfAddresses.Contains(email))
                        : !string.IsNullOrWhiteSpace(normalizedEmail) && selfAddresses.Contains(normalizedEmail),
                    IsGroup = IsGroup,
                    MemberCount = Math.Max(MemberCount, _memberSmtpAddresses.Count),
                    RelationScore = SenderCount * 4 + RecipientCount * 3 + OrganizerCount * 4 + AttendeeCount * 2 + GroupMemberCount + MailCount + CalendarCount + (_sources.Count > 0 ? 5 : 0),
                    MailCount = MailCount,
                    CalendarCount = CalendarCount,
                    SenderCount = SenderCount,
                    RecipientCount = RecipientCount,
                    OrganizerCount = OrganizerCount,
                    AttendeeCount = AttendeeCount,
                    GroupMemberCount = GroupMemberCount,
                    FirstSeen = FirstSeen,
                    LastSeen = LastSeen,
                    RelationKinds = _relationKinds.OrderBy(item => item).ToList(),
                    Sources = _sources.OrderBy(item => item).ToList(),
                    MemberSmtpAddresses = _memberSmtpAddresses.OrderBy(item => item).Take(50).ToList(),
                    MemberGroupSmtpAddresses = _memberGroupSmtpAddresses.OrderBy(item => item).Take(50).ToList(),
                    MemberOfGroupSmtpAddresses = _memberOfGroupSmtpAddresses.OrderBy(item => item).Take(50).ToList(),
                    FolderPaths = _folderPaths.OrderBy(item => item).Take(10).ToList(),
                    RecentMailIds = _recentMailIds.OrderByDescending(item => item.SeenAt).Select(item => item.Id).Distinct().Take(5).ToList(),
                    SampleSubjects = _sampleSubjects.OrderByDescending(item => item.SeenAt).Select(item => item.Subject).Distinct().Take(3).ToList(),
                };
            }

            private void AddCommon(string relationKind, string source, DateTime seenAt)
            {
                _relationKinds.Add(relationKind);
                _sources.Add(source);
                FirstSeen = FirstSeen is null || seenAt < FirstSeen ? seenAt : FirstSeen;
                LastSeen = LastSeen is null || seenAt > LastSeen ? seenAt : LastSeen;
            }

            private static string EmailDomain(string email)
            {
                var at = email.LastIndexOf('@');
                return at >= 0 && at < email.Length - 1 ? email[(at + 1)..].ToLowerInvariant() : string.Empty;
            }

            private static bool IsGroupEntryUserType(string entryUserType)
            {
                return entryUserType.Contains("DistributionList", StringComparison.OrdinalIgnoreCase)
                    || entryUserType.Contains("PublicGroup", StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}

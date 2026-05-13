using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public partial class MailStore
    {
        private List<AddressBookContactDto> _addressBookContacts = new();

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

                return contacts
                    .OrderByDescending(contact => contact.RelationScore)
                    .ThenByDescending(contact => contact.LastSeen ?? DateTime.MinValue)
                    .ThenBy(contact => contact.DisplayName)
                    .Take(Math.Clamp(take <= 0 ? 200 : take, 1, 500))
                    .Select(CloneAddressBookContact)
                    .ToList();
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

        private List<AddressBookContactDto> BuildAddressBookContacts()
        {
            var contacts = new Dictionary<string, AddressBookAccumulator>(StringComparer.OrdinalIgnoreCase);
            var selfAddresses = InferSelfAddresses();

            foreach (var contact in _addressBookContacts)
                AddCachedAddressBookContact(contacts, contact);

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

            return contacts.Values
                .Select(item => item.ToDto(selfAddresses))
                .Where(contact => !string.IsNullOrWhiteSpace(contact.SmtpAddress) || !string.IsNullOrWhiteSpace(contact.DisplayName))
                .ToList();
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

        private HashSet<string> InferSelfAddresses()
        {
            var sentFolders = _folders
                .Where(folder => folder.FolderType == OutlookFolderType.Sent)
                .Select(folder => folder.FolderPath)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            return _mails
                .Where(mail => sentFolders.Contains(mail.FolderPath))
                .Select(mail => NormalizeEmail(mail.Sender.SmtpAddress))
                .Where(email => !string.IsNullOrWhiteSpace(email))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
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
            return contact;
        }

        private static bool ContactMatches(AddressBookContactDto contact, string query)
        {
            return Contains(contact.DisplayName, query)
                || Contains(contact.SmtpAddress, query)
                || Contains(contact.Domain, query);
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
                FolderPaths = new List<string>(contact.FolderPaths),
                RecentMailIds = new List<string>(contact.RecentMailIds),
                SampleSubjects = new List<string>(contact.SampleSubjects),
            };
        }

        private sealed class AddressBookAccumulator
        {
            private readonly HashSet<string> _relationKinds = new(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _sources = new(StringComparer.OrdinalIgnoreCase);
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
                _sources.Add(string.IsNullOrWhiteSpace(source.Source) ? "address_book" : source.Source);
                _relationKinds.Add("address_book");
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
        }
    }
}

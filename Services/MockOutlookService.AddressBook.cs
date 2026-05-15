using Microsoft.AspNetCore.SignalR;

namespace SmartOffice.Hub.Services
{
    public partial class MockOutlookService
    {
        private async Task<bool> TryDispatchAddressBookAsync(PendingCommand command, CancellationToken ct)
        {
            List<AddressBookContactDto> contacts;
            lock (_lock)
            {
                EnsureSeeded();
                contacts = BuildMockAddressBook(command.AddressBookRequest);
                _addinStatus.RecordMockDispatch(command.Type);
            }

            var batchSize = MockAddressBookBatchSize();
            var delayMs = MockAddressBookDelayMs();
            var batchId = Guid.NewGuid().ToString("N");
            var sequence = 1;
            var chunks = contacts.Chunk(batchSize).Select(chunk => chunk.ToList()).ToList();
            if (chunks.Count == 0) chunks.Add(new List<AddressBookContactDto>());

            var resetBatch = new AddressBookBatchDto
            {
                BatchId = batchId,
                Sequence = 0,
                Reset = true,
                IsFinal = false,
                TotalCount = contacts.Count,
                Contacts = new List<AddressBookContactDto>(),
            };
            _mailStore.ApplyAddressBookBatch(resetBatch);
            await _notifications.Clients.All.SendAsync("AddressBookBatchUpdated", resetBatch, ct);

            foreach (var chunk in chunks)
            {
                if (delayMs > 0) await Task.Delay(delayMs, ct);
                var batch = new AddressBookBatchDto
                {
                    BatchId = batchId,
                    Sequence = sequence++,
                    Reset = false,
                    IsFinal = sequence > chunks.Count,
                    TotalCount = contacts.Count,
                    Contacts = chunk,
                };
                _mailStore.ApplyAddressBookBatch(batch);
                _addinStatus.RecordPush("mock address book batch", batch.Contacts.Count);
                await _notifications.Clients.All.SendAsync("AddressBookBatchUpdated", batch, ct);
                await _notifications.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus(), ct);
                await _notifications.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), ct);
            }

            await _notifications.Clients.All.SendAsync("CommandResult", new OutlookCommandResult
            {
                CommandId = command.Id,
                Success = true,
                Message = $"{command.Type} completed by mock backend",
                Timestamp = DateTime.Now,
            }, ct);
            return true;
        }

        private async Task<bool> TryDispatchAddressBookRootsAsync(PendingCommand command, CancellationToken ct)
        {
            List<AddressBookContactDto> contacts;
            lock (_lock)
            {
                EnsureSeeded();
                contacts = BuildMockAddressBook(new AddressBookSyncRequest());
                _addinStatus.RecordMockDispatch(command.Type);
            }

            await DelayMockAddressBookAsync(ct);

            var roots = MockAddressBookRoots()
                .Select(root => new AddressBookRootDto
                {
                    Id = root.Id,
                    Name = root.Name,
                    AddressListType = root.AddressListType,
                    Source = root.Source,
                    EntryCount = contacts.Count(contact => contact.Source == root.Source),
                    CanPageEntries = true,
                })
                .ToList();
            var batch = new AddressBookRootsBatchDto { RequestId = command.Id, Roots = roots };
            _mailStore.SetAddressBookRoots(batch);
            _addinStatus.RecordPush("mock address book roots", roots.Count);
            await _notifications.Clients.All.SendAsync("AddressBookRootsUpdated", batch, ct);
            return true;
        }

        private static List<MockAddressBookRoot> MockAddressBookRoots()
        {
            return new List<MockAddressBookRoot>
            {
                new MockAddressBookRoot
                {
                    Id = "mock-global-address-list",
                    Name = "Global Address List",
                    AddressListType = "olExchangeGlobalAddressList",
                    Source = "global_address_list",
                },
                new MockAddressBookRoot
                {
                    Id = "mock-offline-address-book",
                    Name = "Offline Address Book",
                    AddressListType = "olCustomAddressList",
                    Source = "offline_address_book",
                },
                new MockAddressBookRoot
                {
                    Id = "mock-official-contacts",
                    Name = "Contacts",
                    AddressListType = "olOutlookAddressList",
                    Source = "official_contacts",
                },
                new MockAddressBookRoot
                {
                    Id = "mock-project-directory",
                    Name = "Project Directory",
                    AddressListType = "olCustomAddressList",
                    Source = "project_directory",
                },
                new MockAddressBookRoot
                {
                    Id = "mock-room-resources",
                    Name = "Rooms and Resources",
                    AddressListType = "olCustomAddressList",
                    Source = "room_resources",
                },
            };
        }

        private async Task<bool> TryDispatchAddressListEntriesAsync(PendingCommand command, CancellationToken ct)
        {
            var request = command.AddressBookListEntriesRequest ?? new AddressBookListEntriesRequest();
            List<AddressBookContactDto> contacts;
            lock (_lock)
            {
                EnsureSeeded();
                contacts = BuildMockAddressBook(new AddressBookSyncRequest());
                _addinStatus.RecordMockDispatch(command.Type);
            }

            var source = string.Equals(request.AddressListId, "mock-offline-address-book", StringComparison.OrdinalIgnoreCase)
                ? "offline_address_book"
                : string.Equals(request.AddressListId, "mock-official-contacts", StringComparison.OrdinalIgnoreCase)
                    ? "official_contacts"
                : string.Equals(request.AddressListId, "mock-project-directory", StringComparison.OrdinalIgnoreCase)
                    ? "project_directory"
                : string.Equals(request.AddressListId, "mock-room-resources", StringComparison.OrdinalIgnoreCase)
                    ? "room_resources"
                : "global_address_list";
            var all = contacts
                .Where(contact => string.Equals(contact.Source, source, StringComparison.OrdinalIgnoreCase))
                .OrderBy(contact => contact.DisplayName)
                .ThenBy(contact => contact.SmtpAddress)
                .ToList();
            var pageSize = Math.Clamp(request.PageSize <= 0 ? 100 : request.PageSize, 1, 500);
            var offset = Math.Clamp(request.Offset, 0, all.Count);
            var page = new AddressBookListEntriesPageDto
            {
                RequestId = command.Id,
                AddressListId = request.AddressListId,
                AddressListName = request.AddressListName,
                Offset = offset,
                PageSize = pageSize,
                TotalCount = all.Count,
                HasMore = offset + pageSize < all.Count,
                Contacts = all.Skip(offset).Take(pageSize).ToList(),
            };
            await DelayMockAddressBookAsync(ct);
            _mailStore.SetAddressBookListEntriesPage(page);
            _addinStatus.RecordPush("mock address list entries", page.Contacts.Count);
            await _notifications.Clients.All.SendAsync("AddressBookListEntriesUpdated", page, ct);
            return true;
        }

        private async Task<bool> TryDispatchAddressBookGroupMembersAsync(PendingCommand command, CancellationToken ct)
        {
            var request = command.AddressBookGroupMembersRequest ?? new AddressBookGroupMembersRequest();
            List<AddressBookContactDto> contacts;
            lock (_lock)
            {
                EnsureSeeded();
                contacts = BuildMockAddressBook(new AddressBookSyncRequest());
            }

            var groupKey = (request.GroupSmtpAddress ?? string.Empty).Trim().ToLowerInvariant();
            var group = contacts.FirstOrDefault(contact =>
                contact.IsGroup
                && (string.Equals(contact.SmtpAddress, request.GroupSmtpAddress, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(contact.Id, request.GroupId, StringComparison.OrdinalIgnoreCase)));
            var memberKeys = (group?.MemberSmtpAddresses ?? new List<string>())
                .Select(key => key.Trim())
                .Where(key => !string.IsNullOrWhiteSpace(key))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var max = request.MaxMembers <= 0 ? memberKeys.Count : Math.Min(request.MaxMembers, memberKeys.Count);
            var members = memberKeys
                .Take(max)
                .Select(key => contacts.FirstOrDefault(contact => string.Equals(contact.SmtpAddress, key, StringComparison.OrdinalIgnoreCase))
                    ?? MockContact("mock-member-" + key, key, key, "Unknown", "Member", "group_member"))
                .ToList();

            var batchSize = MockAddressBookGroupMemberBatchSize();
            var delayMs = MockAddressBookDelayMs();
            var batchId = Guid.NewGuid().ToString("N");
            var chunks = members.Chunk(batchSize).Select(chunk => chunk.ToList()).ToList();
            if (chunks.Count == 0) chunks.Add(new List<AddressBookContactDto>());
            for (var index = 0; index < chunks.Count; index++)
            {
                if (delayMs > 0) await Task.Delay(delayMs, ct);
                var batch = new AddressBookGroupMembersBatchDto
                {
                    GroupId = request.GroupId,
                    GroupSmtpAddress = request.GroupSmtpAddress,
                    BatchId = batchId,
                    Sequence = index + 1,
                    Reset = index == 0,
                    IsFinal = index == chunks.Count - 1,
                    TotalCount = memberKeys.Count,
                    Members = chunks[index],
                };
                _mailStore.ApplyAddressBookGroupMembersBatch(batch);
                _addinStatus.RecordPush("mock address book group members", batch.Members.Count);
                await _notifications.Clients.All.SendAsync("AddressBookGroupMembersBatchUpdated", batch, ct);
            }
            if (groupKey.Length > 0)
                _addinStatus.AddLog("info", $"Mock expanded address book group: {groupKey}");
            return true;
        }

        private async Task<bool> TryDispatchAddressBookRelationLookupAsync(PendingCommand command, CancellationToken ct)
        {
            var request = command.AddressBookRelationLookupRequest ?? new AddressBookRelationLookupRequest();
            List<AddressBookContactDto> contacts;
            lock (_lock)
            {
                EnsureSeeded();
                contacts = BuildMockAddressBook(new AddressBookSyncRequest());
                _addinStatus.RecordMockDispatch(command.Type);
            }

            await DelayMockAddressBookAsync(ct);
            var target = FindMockRelationTarget(contacts, request);
            if (target == null) return true;

            var related = new List<AddressBookContactDto> { target };
            if (target.IsGroup)
            {
                related.AddRange(MockGroupMembers(contacts, target, request.Take));
                related.AddRange(MockContainingGroups(contacts, target.SmtpAddress, request.Take));
            }
            else
            {
                related.AddRange(MockContainingGroups(contacts, target.SmtpAddress, request.Take));
            }

            var batch = new AddressBookBatchDto
            {
                BatchId = Guid.NewGuid().ToString("N"),
                Sequence = 1,
                Reset = false,
                IsFinal = true,
                TotalCount = related.Count,
                Contacts = related
                    .Where(contact => contact != null)
                    .GroupBy(contact => NormalizeMockKey(contact.SmtpAddress, contact.Id, contact.DisplayName), StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.First())
                    .ToList(),
            };
            _mailStore.ApplyAddressBookBatch(batch);
            _addinStatus.RecordPush("mock address book relation", batch.Contacts.Count);
            await _notifications.Clients.All.SendAsync("AddressBookBatchUpdated", batch, ct);

            if (target.IsGroup)
            {
                var members = MockGroupMembers(contacts, target, request.Take).ToList();
                var memberBatch = new AddressBookGroupMembersBatchDto
                {
                    GroupId = target.Id,
                    GroupSmtpAddress = target.SmtpAddress,
                    BatchId = Guid.NewGuid().ToString("N"),
                    Sequence = 1,
                    Reset = true,
                    IsFinal = true,
                    TotalCount = members.Count,
                    Members = members,
                };
                _mailStore.ApplyAddressBookGroupMembersBatch(memberBatch);
                await _notifications.Clients.All.SendAsync("AddressBookGroupMembersBatchUpdated", memberBatch, ct);
            }

            await _notifications.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus(), ct);
            await _notifications.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), ct);
            return true;
        }

        private static Task DelayMockAddressBookAsync(CancellationToken ct)
        {
            var delayMs = MockAddressBookDelayMs();
            return delayMs > 0 ? Task.Delay(delayMs, ct) : Task.CompletedTask;
        }

        private static int MockAddressBookDelayMs()
        {
            var raw = Environment.GetEnvironmentVariable("SMARTOFFICE_MOCK_ADDRESS_BOOK_DELAY_MS");
            return int.TryParse(raw, out var delayMs)
                ? Math.Clamp(delayMs, 0, 120000)
                : 750;
        }

        private static int MockAddressBookBatchSize()
        {
            var raw = Environment.GetEnvironmentVariable("SMARTOFFICE_MOCK_ADDRESS_BOOK_BATCH_SIZE");
            return int.TryParse(raw, out var batchSize)
                ? Math.Clamp(batchSize, 1, 100)
                : 25;
        }

        private static int MockAddressBookGroupMemberBatchSize()
        {
            var raw = Environment.GetEnvironmentVariable("SMARTOFFICE_MOCK_ADDRESS_BOOK_GROUP_BATCH_SIZE");
            return int.TryParse(raw, out var batchSize)
                ? Math.Clamp(batchSize, 1, 100)
                : 20;
        }

        private static AddressBookContactDto? FindMockRelationTarget(List<AddressBookContactDto> contacts, AddressBookRelationLookupRequest request)
        {
            var groupKey = NormalizeMockKey(request.GroupSmtpAddress);
            if (!string.IsNullOrWhiteSpace(groupKey))
            {
                var group = contacts.FirstOrDefault(contact => contact.IsGroup && NormalizeMockKey(contact.SmtpAddress) == groupKey);
                if (group != null) return group;
            }

            var id = NormalizeMockKey(request.GroupId, request.Id);
            if (!string.IsNullOrWhiteSpace(id))
            {
                var byId = contacts.FirstOrDefault(contact => NormalizeMockKey(contact.Id) == id);
                if (byId != null) return byId;
            }

            var email = NormalizeMockKey(request.SmtpAddress, request.Email);
            if (!string.IsNullOrWhiteSpace(email))
            {
                var byEmail = contacts.FirstOrDefault(contact => NormalizeMockKey(contact.SmtpAddress) == email);
                if (byEmail != null) return byEmail;
            }

            var query = NormalizeMockKey(request.Query, request.DisplayName);
            if (string.IsNullOrWhiteSpace(query)) return null;
            return contacts.FirstOrDefault(contact =>
                NormalizeMockKey(contact.SmtpAddress) == query
                || NormalizeMockKey(contact.DisplayName) == query)
                ?? contacts.FirstOrDefault(contact =>
                    NormalizeMockKey(contact.SmtpAddress).Contains(query)
                    || NormalizeMockKey(contact.DisplayName).Contains(query));
        }

        private static IEnumerable<AddressBookContactDto> MockGroupMembers(List<AddressBookContactDto> contacts, AddressBookContactDto group, int take)
        {
            var limit = Math.Clamp(take <= 0 ? 50 : take, 1, 500);
            return (group.MemberSmtpAddresses ?? new List<string>())
                .Take(limit)
                .Select(member => contacts.FirstOrDefault(contact => string.Equals(contact.SmtpAddress, member, StringComparison.OrdinalIgnoreCase))
                    ?? MockContact("mock-member-" + member, member, member, "Unknown", "Member", "group_member"));
        }

        private static IEnumerable<AddressBookContactDto> MockContainingGroups(List<AddressBookContactDto> contacts, string smtpAddress, int take)
        {
            var key = NormalizeMockKey(smtpAddress);
            if (string.IsNullOrWhiteSpace(key)) return Enumerable.Empty<AddressBookContactDto>();
            var limit = Math.Clamp(take <= 0 ? 50 : take, 1, 500);
            return contacts
                .Where(contact => contact.IsGroup
                    && (contact.MemberSmtpAddresses ?? new List<string>()).Any(member => NormalizeMockKey(member) == key))
                .Take(limit);
        }

        private static string NormalizeMockKey(params string[] values)
        {
            return values
                .Select(value => (value ?? string.Empty).Trim().Trim('<', '>').ToLowerInvariant())
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;
        }

        private static int MockAddressBookContactTarget()
        {
            var raw = Environment.GetEnvironmentVariable("SMARTOFFICE_MOCK_ADDRESS_BOOK_CONTACT_TARGET");
            return int.TryParse(raw, out var target)
                ? Math.Clamp(target, 50, 5000)
                : 900;
        }

        private List<AddressBookContactDto> BuildMockAddressBook(AddressBookSyncRequest? request)
        {
            request ??= new AddressBookSyncRequest();
            var max = Math.Clamp(request.MaxContacts <= 0 ? MockAddressBookContactTarget() : request.MaxContacts, 1, 5000);
            var contacts = new List<AddressBookContactDto>
            {
                MockContact("mock-contact-001", "Ada Chen", "ada.chen@example.test", "Product", "Product Manager", "global_address_list"),
                MockContact("mock-contact-002", "Ben Lin", "ben.lin@example.test", "Legal", "Counsel", "global_address_list"),
                MockContact("mock-contact-003", "Chris Wang", "chris.wang@example.test", "Sales", "Account Manager", "global_address_list"),
                MockContact("mock-contact-004", "Dana Hsu", "dana.hsu@example.test", "Delivery", "Project Lead", "global_address_list"),
                MockContact("mock-contact-005", "Evan Wu", "evan.wu@example.test", "Engineering", "Staff Engineer", "global_address_list"),
                MockContact("mock-contact-006", "Fiona Tsai", "fiona.tsai@example.test", "Finance", "Finance Manager", "global_address_list"),
                MockContact("mock-contact-007", "Grace Huang", "grace.huang@example.test", "People", "HR Business Partner", "global_address_list"),
                MockContact("mock-contact-008", "Henry Kao", "henry.kao@example.test", "Operations", "Operations Lead", "global_address_list"),
                MockContact("mock-contact-009", "Ivy Lin", "ivy.lin@example.test", "Customer Success", "CS Manager", "global_address_list"),
                MockContact("mock-contact-010", "Jacky Lee", "jacky.lee@example.test", "Security", "Security Analyst", "global_address_list"),
                MockContact("mock-contact-011", "Finance Bot", "finance@example.test", "Finance", "Shared mailbox", "global_address_list"),
                MockContact("mock-contact-012", "Helpdesk", "helpdesk@example.test", "IT", "Shared mailbox", "global_address_list"),
                MockContact("mock-contact-self", "Mock User", "mock.user@example.test", "Operations", "Current user", "global_address_list"),
                MockContact("mock-contact-013", "Vendor Team", "vendor@example.test", "Procurement", "Vendor contact", "official_contacts"),
                MockContact("mock-contact-014", "Mina Park", "mina.park@vendor.example.test", "Partner", "Partner Manager", "official_contacts"),
                MockContact("mock-contact-015", "Noah Sato", "noah.sato@vendor.example.test", "Partner Support", "Support Lead", "official_contacts"),
                MockContact("mock-contact-016", "Olivia Brown", "olivia.brown@customer.example.test", "Customer", "Program Owner", "official_contacts"),
                MockContact("mock-contact-017", "Pierre Martin", "pierre.martin@customer.example.test", "Customer", "Technical Lead", "official_contacts"),
                MockContact("mock-contact-018", "Sam Chen", "sam.chen@example.test", "Product", "Designer", "offline_address_book"),
                MockContact("mock-contact-019", "Sam Chen", "sam.chen.contractor@partner.example.test", "Partner", "Contract Designer", "offline_address_book"),
                MockContact("mock-contact-020", "North Conference Room", "room-north@example.test", "Facilities", "Room mailbox", "room_resources"),
                MockContact("mock-contact-021", "South Conference Room", "room-south@example.test", "Facilities", "Room mailbox", "room_resources"),
                MockContact("mock-contact-022", "War Room", "room-war@example.test", "Facilities", "Room mailbox", "room_resources"),
                MockContact("mock-contact-023", "Release Calendar", "release.calendar@example.test", "Engineering", "Shared calendar", "project_directory"),
                MockContact("mock-contact-024", "Launch Program Mailbox", "launch-program@example.test", "Product", "Shared mailbox", "project_directory"),
                MockContact("mock-contact-025", "Support Queue", "support-queue@example.test", "Customer Success", "Shared mailbox", "project_directory"),
            };

            var departments = new[]
            {
                "Engineering",
                "Product",
                "Finance",
                "Legal",
                "Operations",
                "Sales",
                "Security",
                "People",
                "Customer Success",
                "Facilities",
            };
            var sources = new[] { "global_address_list", "offline_address_book", "official_contacts", "project_directory", "room_resources" };
            var target = Math.Min(max, MockAddressBookContactTarget());
            var groupCount = 20;
            var generatedContactTarget = Math.Max(contacts.Count, target - groupCount);
            for (var index = contacts.Count + 1; contacts.Count < generatedContactTarget; index++)
            {
                var department = departments[index % departments.Length];
                var source = sources[index % sources.Length];
                contacts.Add(MockContact(
                    $"mock-contact-{index:0000}",
                    $"Mock User {index:0000}",
                    $"mock.user.{index:0000}@example.test",
                    department,
                    $"{department} Specialist",
                    source));
            }

            contacts.AddRange(new[]
            {
                MockGroup(
                    "mock-group-001",
                    "Product Launch Working Group",
                    "product-launch@example.test",
                    "Product",
                    "global_address_list",
                    "ada.chen@example.test",
                    "chris.wang@example.test",
                    "sam.chen@example.test",
                    "launch-program@example.test",
                    "mock.user.0036@example.test",
                    "mock.user.0041@example.test",
                    "mock.user.0046@example.test"),
                MockGroup(
                    "mock-group-002",
                    "Finance Approvers",
                    "finance-approvers@example.test",
                    "Finance",
                    "global_address_list",
                    "fiona.tsai@example.test",
                    "finance@example.test",
                    "ben.lin@example.test"),
                MockGroup(
                    "mock-group-003",
                    "Vendor Escalation List",
                    "vendor-escalation@example.test",
                    "Procurement",
                    "offline_address_book",
                    "vendor@example.test",
                    "mina.park@vendor.example.test",
                    "noah.sato@vendor.example.test"),
                MockGroup(
                    "mock-group-004",
                    "All Taipei Office",
                    "all-taipei@example.test",
                    "Operations",
                    "global_address_list",
                    "ada.chen@example.test",
                    "ben.lin@example.test",
                    "dana.hsu@example.test",
                    "evan.wu@example.test",
                    "grace.huang@example.test",
                    "henry.kao@example.test",
                    "ivy.lin@example.test",
                    "jacky.lee@example.test",
                    "mock.user@example.test"),
                MockGroup(
                    "mock-group-005",
                    "Executive Review Circle",
                    "exec-review@example.test",
                    "Leadership",
                    "global_address_list",
                    "product-launch@example.test",
                    "finance-approvers@example.test",
                    "olivia.brown@customer.example.test"),
                MockGroup(
                    "mock-group-006",
                    "Taipei Operations Inner Circle",
                    "taipei-ops-inner@example.test",
                    "Operations",
                    "project_directory",
                    "all-taipei@example.test",
                    "release-command@example.test",
                    "mock.user.0051@example.test",
                    "mock.user.0056@example.test"),
                MockGroup(
                    "mock-group-007",
                    "Release Command",
                    "release-command@example.test",
                    "Engineering",
                    "project_directory",
                    "mock.user@example.test",
                    "mock.user.0061@example.test",
                    "mock.user.0066@example.test",
                    "mock.user.0071@example.test",
                    "room-war@example.test"),
                MockGroup(
                    "mock-group-008",
                    "Customer Escalation Bridge",
                    "customer-escalation@example.test",
                    "Customer Success",
                    "global_address_list",
                    "support-queue@example.test",
                    "vendor-escalation@example.test",
                    "mock.user.0076@example.test",
                    "mock.user.0081@example.test"),
                MockGroup(
                    "mock-group-009",
                    "Nested Loop Alpha",
                    "nested-loop-alpha@example.test",
                    "Engineering",
                    "offline_address_book",
                    "nested-loop-beta@example.test",
                    "mock.user.0086@example.test"),
                MockGroup(
                    "mock-group-010",
                    "Nested Loop Beta",
                    "nested-loop-beta@example.test",
                    "Engineering",
                    "offline_address_book",
                    "nested-loop-alpha@example.test",
                    "mock.user.0091@example.test"),
                MockGroup(
                    "mock-group-011",
                    "Facilities Booking Group",
                    "facilities-booking@example.test",
                    "Facilities",
                    "room_resources",
                    "room-north@example.test",
                    "room-south@example.test",
                    "room-war@example.test",
                    "mock.user.0096@example.test"),
                MockGroup(
                    "mock-group-012",
                    "Release Squad D1",
                    "release-squad-d1@example.test",
                    "Engineering",
                    "project_directory",
                    "mock.user@example.test",
                    "mock.user.0121@example.test",
                    "mock.user.0126@example.test"),
                MockGroup(
                    "mock-group-013",
                    "Operations Team D2",
                    "ops-team-d2@example.test",
                    "Operations",
                    "project_directory",
                    "release-squad-d1@example.test",
                    "mock.user.0131@example.test",
                    "mock.user.0136@example.test"),
                MockGroup(
                    "mock-group-014",
                    "Taipei Department D3",
                    "taipei-department-d3@example.test",
                    "Operations",
                    "global_address_list",
                    "ops-team-d2@example.test",
                    "mock.user.0141@example.test",
                    "mock.user.0146@example.test"),
                MockGroup(
                    "mock-group-015",
                    "APAC Division D4",
                    "apac-division-d4@example.test",
                    "Operations",
                    "global_address_list",
                    "taipei-department-d3@example.test",
                    "mock.user.0151@example.test",
                    "mock.user.0156@example.test"),
                MockGroup(
                    "mock-group-016",
                    "Global Announcements D5",
                    "global-announcements-d5@example.test",
                    "Operations",
                    "global_address_list",
                    "apac-division-d4@example.test",
                    "mock.user.0161@example.test",
                    "mock.user.0166@example.test"),
                MockGroup(
                    "mock-group-017",
                    "Company Broadcast D6",
                    "company-broadcast-d6@example.test",
                    "Operations",
                    "global_address_list",
                    "global-announcements-d5@example.test",
                    "mock.user.0171@example.test",
                    "mock.user.0176@example.test"),
                MockGroup(
                    "mock-group-018",
                    "Executive Broadcast D7",
                    "executive-broadcast-d7@example.test",
                    "Leadership",
                    "global_address_list",
                    "company-broadcast-d6@example.test",
                    "exec-review@example.test",
                    "mock.user.0181@example.test"),
                MockGroup(
                    "mock-group-019",
                    "All Mock Staff",
                    "all-mock-staff@example.test",
                    "Operations",
                    "global_address_list",
                    "all-taipei@example.test",
                    "product-launch@example.test",
                    "finance-approvers@example.test",
                    "customer-escalation@example.test",
                    "facilities-booking@example.test",
                    "executive-broadcast-d7@example.test",
                    "mock.user.0101@example.test",
                    "mock.user.0106@example.test",
                    "mock.user.0111@example.test"),
            });

            return contacts.Take(max).ToList();
        }

        private static AddressBookContactDto MockContact(string id, string name, string email, string department, string jobTitle, string source)
        {
            return new AddressBookContactDto
            {
                Id = id,
                DisplayName = name,
                SmtpAddress = email,
                RawAddress = email,
                AddressType = "SMTP",
                EntryUserType = "olExchangeUserAddressEntry",
                Source = source,
                CompanyName = "Mock Organization",
                Department = department,
                JobTitle = jobTitle,
                Domain = "example.test",
                IsKnown = true,
                Sources = new List<string> { source },
                RelationKinds = new List<string> { "address_book" },
            };
        }

        private static AddressBookContactDto MockGroup(
            string id,
            string name,
            string email,
            string department,
            string source,
            params string[] members)
        {
            var contact = MockContact(id, name, email, department, "Distribution list", source);
            contact.EntryUserType = "olExchangeDistributionListAddressEntry";
            contact.IsGroup = true;
            contact.MemberCount = members.Length;
            contact.MemberSmtpAddresses = members.ToList();
            contact.MemberGroupSmtpAddresses = members
                .Where(member => member.Contains("-group", StringComparison.OrdinalIgnoreCase)
                    || member.EndsWith("-approvers@example.test", StringComparison.OrdinalIgnoreCase)
                    || member.EndsWith("-launch@example.test", StringComparison.OrdinalIgnoreCase)
                    || member.EndsWith("-list@example.test", StringComparison.OrdinalIgnoreCase)
                    || member.EndsWith("-staff@example.test", StringComparison.OrdinalIgnoreCase)
                    || member.EndsWith("-circle@example.test", StringComparison.OrdinalIgnoreCase)
                    || member.EndsWith("-command@example.test", StringComparison.OrdinalIgnoreCase)
                    || member.EndsWith("-bridge@example.test", StringComparison.OrdinalIgnoreCase)
                    || member.Contains("-broadcast-", StringComparison.OrdinalIgnoreCase)
                    || member.Contains("-announcements-", StringComparison.OrdinalIgnoreCase)
                    || member.Contains("-division-", StringComparison.OrdinalIgnoreCase)
                    || member.Contains("-department-", StringComparison.OrdinalIgnoreCase)
                    || member.Contains("-team", StringComparison.OrdinalIgnoreCase)
                    || member.Contains("-squad", StringComparison.OrdinalIgnoreCase)
                    || member.StartsWith("nested-loop-", StringComparison.OrdinalIgnoreCase)
                    || member.StartsWith("all-", StringComparison.OrdinalIgnoreCase)
                    || member.StartsWith("facilities-", StringComparison.OrdinalIgnoreCase)
                    || member.StartsWith("vendor-escalation", StringComparison.OrdinalIgnoreCase)
                    || member.StartsWith("customer-escalation", StringComparison.OrdinalIgnoreCase))
                .ToList();
            return contact;
        }

        private sealed class MockAddressBookRoot
        {
            public string Id { get; set; } = string.Empty;
            public string Name { get; set; } = string.Empty;
            public string AddressListType { get; set; } = string.Empty;
            public string Source { get; set; } = string.Empty;
        }
    }
}

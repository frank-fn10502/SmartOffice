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

            var batch = new AddressBookGroupMembersBatchDto
            {
                GroupId = request.GroupId,
                GroupSmtpAddress = request.GroupSmtpAddress,
                BatchId = Guid.NewGuid().ToString("N"),
                Sequence = 1,
                Reset = true,
                IsFinal = true,
                TotalCount = memberKeys.Count,
                Members = members,
            };
            _mailStore.ApplyAddressBookGroupMembersBatch(batch);
            _addinStatus.RecordPush("mock address book group members", batch.Members.Count);
            await _notifications.Clients.All.SendAsync("AddressBookGroupMembersBatchUpdated", batch, ct);
            if (groupKey.Length > 0)
                _addinStatus.AddLog("info", $"Mock expanded address book group: {groupKey}");
            return true;
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

        private List<AddressBookContactDto> BuildMockAddressBook(AddressBookSyncRequest? request)
        {
            request ??= new AddressBookSyncRequest();
            var max = request.MaxContacts <= 0 ? 0 : Math.Clamp(request.MaxContacts, 1, 5000);
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
                MockContact("mock-contact-013", "Vendor Team", "vendor@example.test", "Procurement", "Vendor contact", "official_contacts"),
                MockContact("mock-contact-014", "Mina Park", "mina.park@vendor.example.test", "Partner", "Partner Manager", "official_contacts"),
                MockContact("mock-contact-015", "Noah Sato", "noah.sato@vendor.example.test", "Partner Support", "Support Lead", "official_contacts"),
                MockContact("mock-contact-016", "Olivia Brown", "olivia.brown@customer.example.test", "Customer", "Program Owner", "official_contacts"),
                MockContact("mock-contact-017", "Pierre Martin", "pierre.martin@customer.example.test", "Customer", "Technical Lead", "official_contacts"),
                MockContact("mock-contact-018", "Sam Chen", "sam.chen@example.test", "Product", "Designer", "offline_address_book"),
                MockContact("mock-contact-019", "Sam Chen", "sam.chen.contractor@partner.example.test", "Partner", "Contract Designer", "offline_address_book"),
                MockGroup(
                    "mock-group-001",
                    "Product Launch Working Group",
                    "product-launch@example.test",
                    "Product",
                    "global_address_list",
                    "ada.chen@example.test",
                    "chris.wang@example.test",
                    "sam.chen@example.test"),
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
                    "jacky.lee@example.test"),
                MockGroup(
                    "mock-group-005",
                    "Executive Review Circle",
                    "exec-review@example.test",
                    "Leadership",
                    "global_address_list",
                    "product-launch@example.test",
                    "finance-approvers@example.test",
                    "olivia.brown@customer.example.test"),
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
            };
            var target = max > 0 ? Math.Min(max, 200) : 200;
            for (var index = contacts.Count + 1; contacts.Count < target; index++)
            {
                var department = departments[index % departments.Length];
                contacts.Add(MockContact(
                    $"mock-contact-{index:000}",
                    $"Mock User {index:000}",
                    $"mock.user.{index:000}@example.test",
                    department,
                    $"{department} Specialist",
                    index % 3 == 0 ? "offline_address_book" : "global_address_list"));
            }

            return max > 0 ? contacts.Take(max).ToList() : contacts;
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
            contact.MemberGroupSmtpAddresses = members.Where(member => member.EndsWith("-approvers@example.test") || member.EndsWith("-launch@example.test")).ToList();
            return contact;
        }
    }
}

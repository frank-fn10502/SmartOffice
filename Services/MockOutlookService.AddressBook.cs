namespace SmartOffice.Hub.Services
{
    public partial class MockOutlookService
    {
        private List<AddressBookContactDto> BuildMockAddressBook(AddressBookSyncRequest? request)
        {
            request ??= new AddressBookSyncRequest();
            var max = Math.Clamp(request.MaxContacts <= 0 ? 1000 : request.MaxContacts, 1, 5000);
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
            };

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
            return contact;
        }
    }
}

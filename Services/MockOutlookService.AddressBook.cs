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
                MockContact("mock-contact-001", "Ada Chen", "ada.chen@example.test", "Product", "Product Manager", "official_contacts"),
                MockContact("mock-contact-002", "Ben Lin", "ben.lin@example.test", "Legal", "Counsel", "official_contacts"),
                MockContact("mock-contact-003", "Chris Wang", "chris.wang@example.test", "Sales", "Account Manager", "global_address_list"),
                MockContact("mock-contact-004", "Dana Hsu", "dana.hsu@example.test", "Delivery", "Project Lead", "global_address_list"),
                MockContact("mock-contact-005", "Finance Bot", "finance@example.test", "Finance", "Shared mailbox", "global_address_list"),
                MockContact("mock-contact-006", "Vendor Team", "vendor@example.test", "Procurement", "Vendor contact", "official_contacts"),
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
    }
}

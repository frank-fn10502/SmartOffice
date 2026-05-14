namespace SmartOffice.Hub.Services
{
    public partial class MailStore
    {
        public void SetAddressBookRoots(AddressBookRootsBatchDto batch)
        {
            lock (_lock)
            {
                _addressBookRoots = (batch?.Roots ?? new List<AddressBookRootDto>())
                    .Where(root => !string.IsNullOrWhiteSpace(root.Id) || !string.IsNullOrWhiteSpace(root.Name))
                    .Select(CloneAddressBookRoot)
                    .ToList();
            }
        }

        public List<AddressBookRootDto> GetAddressBookRoots()
        {
            lock (_lock)
            {
                return _addressBookRoots.Select(CloneAddressBookRoot).ToList();
            }
        }

        public void SetAddressBookListEntriesPage(AddressBookListEntriesPageDto page)
        {
            page ??= new AddressBookListEntriesPageDto();
            lock (_lock)
            {
                if (!string.IsNullOrWhiteSpace(page.RequestId))
                    _addressBookListEntriesByRequestId[page.RequestId] = CloneAddressBookListEntriesPage(page);

                foreach (var contact in page.Contacts)
                    UpsertAddressBookContact(contact);
            }
        }

        public AddressBookListEntriesPageDto GetAddressBookListEntriesPage(string requestId)
        {
            lock (_lock)
            {
                return _addressBookListEntriesByRequestId.TryGetValue(requestId ?? string.Empty, out var page)
                    ? EnrichAddressBookListEntriesPage(page)
                    : new AddressBookListEntriesPageDto { RequestId = requestId ?? string.Empty };
            }
        }

        private AddressBookListEntriesPageDto EnrichAddressBookListEntriesPage(AddressBookListEntriesPageDto page)
        {
            var enrichedContacts = BuildAddressBookContacts()
                .SelectMany(contact => ContactKeys(contact).Select(key => (Key: key, Contact: contact)))
                .Where(item => !string.IsNullOrWhiteSpace(item.Key))
                .GroupBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().Contact, StringComparer.OrdinalIgnoreCase);

            var clone = CloneAddressBookListEntriesPage(page);
            clone.Contacts = clone.Contacts
                .Select(contact => TryGetEnrichedAddressBookContact(contact, enrichedContacts))
                .ToList();
            return clone;
        }

        private static AddressBookContactDto TryGetEnrichedAddressBookContact(
            AddressBookContactDto contact,
            Dictionary<string, AddressBookContactDto> enrichedContacts)
        {
            foreach (var key in ContactKeys(contact))
            {
                if (enrichedContacts.TryGetValue(key, out var enriched))
                    return CloneAddressBookContact(enriched);
            }

            return CloneAddressBookContact(contact);
        }

        private static AddressBookRootDto CloneAddressBookRoot(AddressBookRootDto root)
        {
            return new AddressBookRootDto
            {
                Id = root.Id,
                Name = root.Name,
                AddressListType = root.AddressListType,
                Source = root.Source,
                EntryCount = root.EntryCount,
                CanPageEntries = root.CanPageEntries,
            };
        }

        private static AddressBookListEntriesPageDto CloneAddressBookListEntriesPage(AddressBookListEntriesPageDto page)
        {
            return new AddressBookListEntriesPageDto
            {
                RequestId = page.RequestId,
                AddressListId = page.AddressListId,
                AddressListName = page.AddressListName,
                Offset = page.Offset,
                PageSize = page.PageSize,
                TotalCount = page.TotalCount,
                HasMore = page.HasMore,
                Contacts = page.Contacts.Select(CloneAddressBookContact).ToList(),
            };
        }
    }
}

namespace SmartOffice.Hub.Services
{
    public class OutlookAddressBookBackgroundSyncService : BackgroundService
    {
        private static readonly TimeSpan InitialDelay = TimeSpan.FromSeconds(20);
        private static readonly TimeSpan IdleDelay = TimeSpan.FromSeconds(15);
        private static readonly TimeSpan CommandDelay = TimeSpan.FromSeconds(3);
        private const int PageSize = 50;

        private readonly MailStore _mailStore;
        private readonly OutlookCommandQueue _commandQueue;
        private readonly Dictionary<string, AddressListProgress> _addressListProgress = new(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _requestedGroups = new(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _requestedRelations = new(StringComparer.OrdinalIgnoreCase);

        public OutlookAddressBookBackgroundSyncService(
            MailStore mailStore,
            OutlookCommandQueue commandQueue)
        {
            _mailStore = mailStore;
            _commandQueue = commandQueue;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            await DelayQuietly(InitialDelay, stoppingToken);
            while (!stoppingToken.IsCancellationRequested)
            {
                var didWork = await TryDoOneStepAsync(stoppingToken);
                await DelayQuietly(didWork ? CommandDelay : IdleDelay, stoppingToken);
            }
        }

        private async Task<bool> TryDoOneStepAsync(CancellationToken ct)
        {
            var nextContextRelation = NextRelationToResolve(onlyUserVisibleContext: true);
            if (nextContextRelation is not null)
            {
                await ResolveRelationAsync(nextContextRelation, ct);
                return true;
            }

            var nextContextGroup = NextGroupToExpand(onlyUserVisibleContext: true);
            if (nextContextGroup is not null)
            {
                await FetchGroupMembersAsync(nextContextGroup, ct);
                return true;
            }

            var roots = _mailStore.GetAddressBookRoots();
            if (roots.Count == 0)
            {
                await _commandQueue.ExecuteAsync(new PendingCommand { Type = "fetch_address_book_roots" }, ct: ct);
                return true;
            }

            var nextRoot = roots
                .Where(root => root.CanPageEntries)
                .FirstOrDefault(root => !_addressListProgress.TryGetValue(RootKey(root), out var progress) || !progress.Completed);
            if (nextRoot is not null)
            {
                await FetchOneAddressListPageAsync(nextRoot, ct);
                return true;
            }

            var nextGroup = NextGroupToExpand(onlyUserVisibleContext: false);
            if (nextGroup is not null)
            {
                await FetchGroupMembersAsync(nextGroup, ct);
                return true;
            }

            var nextRelation = NextRelationToResolve(onlyUserVisibleContext: false);
            if (nextRelation is not null)
            {
                await ResolveRelationAsync(nextRelation, ct);
                return true;
            }

            return false;
        }

        private AddressBookContactDto? NextRelationToResolve(bool onlyUserVisibleContext)
        {
            return _mailStore.GetAddressBookContacts(take: 0)
                .Where(contact => ContactNeedsRelationLookup(contact))
                .Where(contact => !onlyUserVisibleContext || HasUserVisibleOutlookContext(contact))
                .FirstOrDefault(contact => !_requestedRelations.Contains(ContactKey(contact)));
        }

        private AddressBookContactDto? NextGroupToExpand(bool onlyUserVisibleContext)
        {
            return _mailStore.GetAddressBookContacts(take: 0)
                .Where(contact => contact.IsGroup && !contact.GroupMembersLoaded && !contact.GroupMembersLoading)
                .Where(contact => !onlyUserVisibleContext || HasUserVisibleOutlookContext(contact))
                .FirstOrDefault(contact => !_requestedGroups.Contains(GroupKey(contact)));
        }

        private async Task ResolveRelationAsync(AddressBookContactDto contact, CancellationToken ct)
        {
            var result = await _commandQueue.ExecuteAsync(new PendingCommand
            {
                Type = "address_book_relation_lookup",
                AddressBookRelationLookupRequest = new AddressBookRelationLookupRequest
                {
                    TargetKind = contact.IsGroup ? "group" : "person",
                    Id = contact.Id,
                    DisplayName = contact.DisplayName,
                    SmtpAddress = contact.SmtpAddress,
                    Email = contact.SmtpAddress,
                    GroupId = contact.IsGroup ? contact.Id : string.Empty,
                    GroupSmtpAddress = contact.IsGroup ? contact.SmtpAddress : string.Empty,
                    Take = 200,
                },
            }, ct: ct);
            if (result.Success) _requestedRelations.Add(ContactKey(contact));
        }

        private async Task FetchGroupMembersAsync(AddressBookContactDto group, CancellationToken ct)
        {
            var result = await _commandQueue.ExecuteAsync(new PendingCommand
            {
                Type = "fetch_address_book_group_members",
                AddressBookGroupMembersRequest = new AddressBookGroupMembersRequest
                {
                    GroupId = group.Id,
                    GroupSmtpAddress = group.SmtpAddress,
                    MaxMembers = 5000,
                },
            }, ct: ct);
            if (result.Success) _requestedGroups.Add(GroupKey(group));
        }

        private async Task FetchOneAddressListPageAsync(AddressBookRootDto root, CancellationToken ct)
        {
            var rootKey = RootKey(root);
            if (!_addressListProgress.TryGetValue(rootKey, out var progress))
            {
                progress = new AddressListProgress();
                _addressListProgress[rootKey] = progress;
            }

            var command = new PendingCommand
            {
                Type = "fetch_address_list_entries",
                AddressBookListEntriesRequest = new AddressBookListEntriesRequest
                {
                    AddressListId = root.Id,
                    AddressListName = root.Name,
                    Offset = progress.Offset,
                    PageSize = PageSize,
                },
            };
            await _commandQueue.ExecuteAsync(command, ct: ct);
            var page = _mailStore.GetAddressBookListEntriesPage(command.Id);
            progress.Offset = Math.Max(progress.Offset + page.Contacts.Count, page.Offset + page.Contacts.Count);
            progress.Completed = !page.HasMore || page.Contacts.Count == 0;
        }

        private static async Task DelayQuietly(TimeSpan delay, CancellationToken ct)
        {
            try
            {
                await Task.Delay(delay, ct);
            }
            catch (OperationCanceledException) { }
        }

        private static string RootKey(AddressBookRootDto root)
        {
            return string.IsNullOrWhiteSpace(root.Id)
                ? (root.Name ?? string.Empty).Trim().ToLowerInvariant()
                : root.Id.Trim().ToLowerInvariant();
        }

        private static string GroupKey(AddressBookContactDto contact)
        {
            return string.IsNullOrWhiteSpace(contact.SmtpAddress)
                ? (contact.Id ?? contact.DisplayName ?? string.Empty).Trim().ToLowerInvariant()
                : contact.SmtpAddress.Trim().ToLowerInvariant();
        }

        private static string ContactKey(AddressBookContactDto contact)
        {
            return string.IsNullOrWhiteSpace(contact.SmtpAddress)
                ? (contact.Id ?? contact.RawAddress ?? contact.DisplayName ?? string.Empty).Trim().ToLowerInvariant()
                : contact.SmtpAddress.Trim().ToLowerInvariant();
        }

        private static bool ContactNeedsRelationLookup(AddressBookContactDto contact)
        {
            if (string.IsNullOrWhiteSpace(ContactKey(contact))) return false;
            if (contact.IsGroup)
                return contact.MemberOfGroupSmtpAddresses.Count == 0;
            return contact.IsLikelySelf || (contact.MemberOfGroupSmtpAddresses.Count == 0 && HasUserVisibleOutlookContext(contact));
        }

        private static bool HasUserVisibleOutlookContext(AddressBookContactDto contact)
        {
            if (contact.IsLikelySelf || contact.IsRelatedToSelf) return true;
            if (contact.MailCount > 0 || contact.CalendarCount > 0) return true;
            if (contact.Sources.Any(source => source.Equals("mail", StringComparison.OrdinalIgnoreCase)
                || source.Equals("calendar", StringComparison.OrdinalIgnoreCase)
                || source.Equals("mail_recipient", StringComparison.OrdinalIgnoreCase))) return true;
            return contact.RelationKinds.Any(kind => !kind.Equals("address_book", StringComparison.OrdinalIgnoreCase));
        }

        private sealed class AddressListProgress
        {
            public int Offset { get; set; }
            public bool Completed { get; set; }
        }
    }
}

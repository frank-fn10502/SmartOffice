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

            return false;
        }

        private AddressBookContactDto? NextGroupToExpand(bool onlyUserVisibleContext)
        {
            return _mailStore.GetAddressBookContacts(take: 0)
                .Where(contact => contact.IsGroup && !contact.GroupMembersLoaded && !contact.GroupMembersLoading)
                .Where(contact => !onlyUserVisibleContext || HasUserVisibleOutlookContext(contact))
                .FirstOrDefault(contact => _requestedGroups.Add(GroupKey(contact)));
        }

        private async Task FetchGroupMembersAsync(AddressBookContactDto group, CancellationToken ct)
        {
            await _commandQueue.ExecuteAsync(new PendingCommand
            {
                Type = "fetch_address_book_group_members",
                AddressBookGroupMembersRequest = new AddressBookGroupMembersRequest
                {
                    GroupId = group.Id,
                    GroupSmtpAddress = group.SmtpAddress,
                    MaxMembers = 5000,
                },
            }, ct: ct);
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

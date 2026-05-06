using Microsoft.AspNetCore.SignalR;
using SmartOffice.Hub.Hubs;
using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public class OutlookFolderCacheService
    {
        private static readonly TimeSpan FolderCacheMaxAge = TimeSpan.FromMinutes(30);

        private readonly MailStore _mailStore;
        private readonly OutlookCommandQueue _commandQueue;
        private readonly IHubContext<NotificationHub> _hub;

        public OutlookFolderCacheService(
            MailStore mailStore,
            OutlookCommandQueue commandQueue,
            IHubContext<NotificationHub> hub)
        {
            _mailStore = mailStore;
            _commandQueue = commandQueue;
            _hub = hub;
        }

        public async Task<bool> EnsureFolderCacheAsync(PendingCommand? ownerCommand, CancellationToken ct)
        {
            if (_mailStore.CountFolders() <= 0 || _mailStore.IsFolderCacheStale(FolderCacheMaxAge))
            {
                await SendFolderCacheProgressAsync(ownerCommand, "Hub is loading Outlook folders before running folder-dependent command.", ct);

                var result = await FetchFolderRootsQueuedAsync(ct);
                if (!result.Success) return false;
            }

            await LoadPendingFolderDiscoveryTargetsAsync(ownerCommand, ct);
            return _mailStore.CountFolders() > 0;
        }

        public async Task<OutlookQueuedCommandResult> FetchFolderRootsQueuedAsync(CancellationToken ct)
        {
            var syncId = Guid.NewGuid().ToString();
            var rootCommand = new PendingCommand
            {
                Type = "fetch_folder_roots",
                FolderDiscoveryRequest = new FolderDiscoveryRequest
                {
                    SyncId = syncId,
                    Reset = true,
                    MaxDepth = 0,
                    MaxChildren = 50,
                },
            };

            var rootResult = await _commandQueue.ExecuteQueuedCommandAsync(
                rootCommand,
                () => _mailStore.CountStoreRoots() > 0,
                ensureReady: true,
                ct: ct);
            if (!rootResult.Success) return rootResult;

            return OutlookQueuedCommandResult.Completed(rootCommand.Id, "completed", "Hub loaded Outlook stores and root folders.");
        }

        private async Task LoadPendingFolderDiscoveryTargetsAsync(PendingCommand? ownerCommand, CancellationToken ct)
        {
            const int maxDiscoveryCommands = 200;
            var dispatched = 0;
            while (dispatched < maxDiscoveryCommands)
            {
                var target = _mailStore.GetPendingFolderDiscoveryTargets().FirstOrDefault();
                if (target is null) return;

                await SendFolderCacheProgressAsync(ownerCommand, $"Hub is loading folder children before running folder-dependent command: {target.FolderPath}", ct, target);

                var command = new PendingCommand
                {
                    Type = "fetch_folder_children",
                    FolderDiscoveryRequest = new FolderDiscoveryRequest
                    {
                        StoreId = target.StoreId,
                        ParentEntryId = target.EntryId,
                        ParentFolderPath = target.FolderPath,
                        MaxDepth = 1,
                        MaxChildren = 200,
                        Reset = false,
                    },
                };
                var result = await _commandQueue.ExecuteQueuedCommandAsync(
                    command,
                    () => _mailStore.IsFolderChildrenLoaded(target.StoreId, target.EntryId, target.FolderPath),
                    ensureReady: true,
                    ct: ct);
                if (!result.Success) return;
                dispatched++;
            }
        }

        private async Task SendFolderCacheProgressAsync(PendingCommand? ownerCommand, string message, CancellationToken ct, FolderDto? target = null)
        {
            if (ownerCommand?.Type != "search_mails" || ownerCommand.SearchMailsRequest is null) return;
            var progress = _mailStore.UpdateMailSearchProgress(new MailSearchProgressDto
            {
                SearchId = ownerCommand.SearchMailsRequest.SearchId,
                CommandId = ownerCommand.Id,
                Status = "running",
                Phase = "load_folders",
                CurrentStoreId = target?.StoreId ?? string.Empty,
                CurrentFolderPath = target?.FolderPath ?? string.Empty,
                Message = message,
                Timestamp = DateTime.Now,
            });
            await _hub.Clients.All.SendAsync("MailSearchProgress", progress, ct);
        }
    }
}

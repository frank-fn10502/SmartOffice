using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public partial class MockOutlookService
    {
        private static string BuildChatReply(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "Mock Outlook 收到空白訊息。";

            return $"Mock Outlook 已收到：「{text.Trim()}」。目前本機資料可直接測試 request、更新與即時廣播。";
        }

        private void RefreshFolderCounts()
        {
            foreach (var folder in _mockFolders)
            {
                folder.ItemCount = _mockMails.Count(mail => mail.FolderPath == folder.FolderPath);
                folder.HasChildren = _mockFolders.Any(child => child.ParentFolderPath == folder.FolderPath);
                folder.ChildrenLoaded = !folder.HasChildren;
                folder.DiscoveryState = folder.HasChildren ? "partial" : "loaded";
            }
        }

        private void AddMockFolder(string name, string folderPath, string parentFolderPath, string storeId, bool isStoreRoot = false)
        {
            _mockFolders.Add(new FolderDto
            {
                Name = name,
                EntryId = MockFolderEntryId(storeId, folderPath),
                FolderPath = folderPath,
                ParentEntryId = string.IsNullOrWhiteSpace(parentFolderPath) ? string.Empty : MockFolderEntryId(storeId, parentFolderPath),
                ParentFolderPath = parentFolderPath,
                StoreId = storeId,
                IsStoreRoot = isStoreRoot,
                FolderType = isStoreRoot ? OutlookFolderType.StoreRoot : OutlookFolderType.Mail,
                DefaultItemType = isStoreRoot ? -1 : 0,
                DiscoveryState = "loaded",
            });
            RefreshFolderCounts();
        }

        private FolderDto? FindFolder(string path)
        {
            return _mockFolders.FirstOrDefault(folder => folder.FolderPath == path);
        }

        private static List<FolderDto> CloneFolders(List<FolderDto> folders)
        {
            return folders.Select(folder => new FolderDto
            {
                Name = folder.Name,
                EntryId = folder.EntryId,
                FolderPath = folder.FolderPath,
                ParentEntryId = folder.ParentEntryId,
                ParentFolderPath = folder.ParentFolderPath,
                ItemCount = folder.ItemCount,
                StoreId = folder.StoreId,
                IsStoreRoot = folder.IsStoreRoot,
                FolderType = folder.FolderType,
                DefaultItemType = folder.DefaultItemType,
                IsHidden = folder.IsHidden,
                IsSystem = folder.IsSystem,
                HasChildren = folder.HasChildren,
                ChildrenLoaded = folder.ChildrenLoaded,
                DiscoveryState = folder.DiscoveryState,
            }).ToList();
        }

        private FolderSyncBatchDto BuildFolderRootsBatch(bool reset)
        {
            RefreshFolderCounts();
            return new FolderSyncBatchDto
            {
                SyncId = Guid.NewGuid().ToString(),
                Sequence = 1,
                Reset = reset,
                IsFinal = true,
                Stores = _mockStores.Select(store => new OutlookStoreDto
                {
                    StoreId = store.StoreId,
                    DisplayName = store.DisplayName,
                    StoreKind = store.StoreKind,
                    StoreFilePath = store.StoreFilePath,
                    RootFolderPath = store.RootFolderPath,
                }).ToList(),
                Folders = CloneFolders(_mockFolders.Where(folder => folder.IsStoreRoot).ToList()),
            };
        }

        private FolderSyncBatchDto BuildFullFolderBatch()
        {
            RefreshFolderCounts();
            return new FolderSyncBatchDto
            {
                SyncId = Guid.NewGuid().ToString(),
                Sequence = 1,
                Reset = true,
                IsFinal = true,
                Stores = _mockStores.Select(store => new OutlookStoreDto
                {
                    StoreId = store.StoreId,
                    DisplayName = store.DisplayName,
                    StoreKind = store.StoreKind,
                    StoreFilePath = store.StoreFilePath,
                    RootFolderPath = store.RootFolderPath,
                }).ToList(),
                Folders = CloneFolders(_mockFolders),
            };
        }

        private FolderSyncBatchDto BuildFolderChildrenBatch(FolderDiscoveryRequest? request)
        {
            RefreshFolderCounts();
            request ??= new FolderDiscoveryRequest();
            var parent = _mockFolders.FirstOrDefault(folder =>
                string.Equals(folder.StoreId, request.StoreId, StringComparison.OrdinalIgnoreCase)
                && (
                    (!string.IsNullOrWhiteSpace(request.ParentEntryId) && string.Equals(folder.EntryId, request.ParentEntryId, StringComparison.OrdinalIgnoreCase))
                    || (!string.IsNullOrWhiteSpace(request.ParentFolderPath) && string.Equals(folder.FolderPath, request.ParentFolderPath, StringComparison.OrdinalIgnoreCase))
                ));

            var folders = new List<FolderDto>();
            if (parent is not null)
            {
                var loadedParent = new FolderDto
                {
                    Name = parent.Name,
                    EntryId = parent.EntryId,
                    FolderPath = parent.FolderPath,
                    ParentEntryId = parent.ParentEntryId,
                    ParentFolderPath = parent.ParentFolderPath,
                    ItemCount = parent.ItemCount,
                    StoreId = parent.StoreId,
                    IsStoreRoot = parent.IsStoreRoot,
                    FolderType = parent.FolderType,
                    DefaultItemType = parent.DefaultItemType,
                    IsHidden = parent.IsHidden,
                    IsSystem = parent.IsSystem,
                    HasChildren = parent.HasChildren,
                    ChildrenLoaded = true,
                    DiscoveryState = "loaded",
                };
                folders.Add(loadedParent);
                folders.AddRange(CloneFolders(_mockFolders
                    .Where(folder => string.Equals(folder.ParentFolderPath, parent.FolderPath, StringComparison.OrdinalIgnoreCase))
                    .Take(Math.Clamp(request.MaxChildren <= 0 ? 50 : request.MaxChildren, 1, 200))
                    .ToList()));
            }

            return new FolderSyncBatchDto
            {
                SyncId = string.IsNullOrWhiteSpace(request.SyncId) ? Guid.NewGuid().ToString() : request.SyncId,
                Sequence = 1,
                Reset = false,
                IsFinal = true,
                Stores = new List<OutlookStoreDto>(),
                Folders = folders,
            };
        }

        private FolderSyncBatchDto BuildFolderCountsBatch(params string?[] folderPaths)
        {
            RefreshFolderCounts();
            var requested = folderPaths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(path => path!)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            return new FolderSyncBatchDto
            {
                SyncId = Guid.NewGuid().ToString(),
                Sequence = 1,
                Reset = false,
                IsFinal = true,
                Stores = new List<OutlookStoreDto>(),
                Folders = CloneFolders(_mockFolders.Where(folder => requested.Contains(folder.FolderPath)).ToList()),
            };
        }

        private void MoveFolderToDeletedItems(string folderPath)
        {
            var target = FindFolder(folderPath);
            if (target is null || target.IsStoreRoot || target.IsHidden || target.IsSystem) return;

            var deletedFolder = DefaultDeletedFolderFor(target.FolderPath);

            if (deletedFolder is null) return;
            if (IsInDefaultDeletedItems(target.FolderPath)) return;

            var oldRootPath = target.FolderPath;
            var newRootPath = UniqueFolderPath(deletedFolder.FolderPath, target.Name, target.StoreId, oldRootPath);
            var movedFolders = _mockFolders
                .Where(folder =>
                    string.Equals(folder.FolderPath, oldRootPath, StringComparison.OrdinalIgnoreCase)
                    || folder.FolderPath.StartsWith($"{oldRootPath}\\", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var folder in movedFolders)
            {
                var oldFolderPath = folder.FolderPath;
                folder.FolderPath = RebasePath(oldFolderPath, oldRootPath, newRootPath);
                folder.ParentFolderPath = string.Equals(oldFolderPath, oldRootPath, StringComparison.OrdinalIgnoreCase)
                    ? deletedFolder.FolderPath
                    : RebasePath(folder.ParentFolderPath, oldRootPath, newRootPath);
                folder.EntryId = MockFolderEntryId(folder.StoreId, folder.FolderPath);
                folder.ParentEntryId = MockFolderEntryId(folder.StoreId, folder.ParentFolderPath);
            }

            foreach (var mail in _mockMails.Where(mail =>
                string.Equals(mail.FolderPath, oldRootPath, StringComparison.OrdinalIgnoreCase)
                || mail.FolderPath.StartsWith($"{oldRootPath}\\", StringComparison.OrdinalIgnoreCase)))
            {
                mail.FolderPath = RebasePath(mail.FolderPath, oldRootPath, newRootPath);
            }
        }

        private string UniqueFolderPath(string parentFolderPath, string folderName, string storeId, string currentPath)
        {
            var basePath = $"{parentFolderPath}\\{folderName}";
            var path = basePath;
            var index = 2;
            while (_mockFolders.Any(folder =>
                !string.Equals(folder.FolderPath, currentPath, StringComparison.OrdinalIgnoreCase)
                && string.Equals(folder.StoreId, storeId, StringComparison.OrdinalIgnoreCase)
                && string.Equals(folder.FolderPath, path, StringComparison.OrdinalIgnoreCase)))
            {
                path = $"{basePath} ({index})";
                index += 1;
            }

            return path;
        }

        private FolderDto? DefaultDeletedFolderFor(string folderPath)
        {
            var target = FindFolder(folderPath);
            var storeId = target?.StoreId ?? _mockFolders.FirstOrDefault(folder =>
                !folder.IsStoreRoot
                && folderPath.StartsWith($"{folder.FolderPath}\\", StringComparison.OrdinalIgnoreCase))?.StoreId;

            return _mockFolders.FirstOrDefault(folder =>
                folder.FolderType == OutlookFolderType.Deleted
                && (
                    string.IsNullOrWhiteSpace(storeId)
                    || string.Equals(folder.StoreId, storeId, StringComparison.OrdinalIgnoreCase)
                ));
        }

        private bool IsInDefaultDeletedItems(string folderPath)
        {
            var deletedFolder = DefaultDeletedFolderFor(folderPath);
            return deletedFolder is not null
                && (
                    string.Equals(folderPath, deletedFolder.FolderPath, StringComparison.OrdinalIgnoreCase)
                    || folderPath.StartsWith($"{deletedFolder.FolderPath}\\", StringComparison.OrdinalIgnoreCase)
                );
        }

        private static string RebasePath(string path, string oldRootPath, string newRootPath)
        {
            if (string.Equals(path, oldRootPath, StringComparison.OrdinalIgnoreCase)) return newRootPath;
            if (!path.StartsWith($"{oldRootPath}\\", StringComparison.OrdinalIgnoreCase)) return path;
            return $"{newRootPath}{path[oldRootPath.Length..]}";
        }

        private static FolderSyncBatchDto ToFullFolderBatch(FolderSyncBatchDto source, FolderSnapshotDto snapshot)
        {
            return new FolderSyncBatchDto
            {
                SyncId = source.SyncId,
                Sequence = source.Sequence,
                Reset = true,
                IsFinal = source.IsFinal,
                Stores = snapshot.Stores,
                Folders = snapshot.Folders,
            };
        }

        private static string MockFolderEntryId(string storeId, string folderPath)
        {
            return $"{storeId}:{folderPath}";
        }

        private static MailItemDto CloneMail(MailItemDto mail)
        {
            return new MailItemDto
            {
                Id = mail.Id,
                Subject = mail.Subject,
                Sender = CloneRecipient(mail.Sender),
                ToRecipients = CloneRecipients(mail.ToRecipients),
                CcRecipients = CloneRecipients(mail.CcRecipients),
                BccRecipients = CloneRecipients(mail.BccRecipients),
                ReceivedTime = mail.ReceivedTime,
                Body = mail.Body,
                BodyHtml = mail.BodyHtml,
                FolderPath = mail.FolderPath,
                MessageClass = mail.MessageClass,
                ConversationId = mail.ConversationId,
                ConversationTopic = mail.ConversationTopic,
                ConversationIndex = mail.ConversationIndex,
                Categories = mail.Categories,
                IsRead = mail.IsRead,
                IsMarkedAsTask = mail.IsMarkedAsTask,
                AttachmentCount = mail.AttachmentCount,
                AttachmentNames = mail.AttachmentNames,
                FlagRequest = mail.FlagRequest,
                FlagInterval = mail.FlagInterval,
                TaskStartDate = mail.TaskStartDate,
                TaskDueDate = mail.TaskDueDate,
                TaskCompletedDate = mail.TaskCompletedDate,
                Importance = mail.Importance,
                Sensitivity = mail.Sensitivity,
            };
        }

        internal static CalendarEventDto CloneCalendarEvent(CalendarEventDto item)
        {
            return new CalendarEventDto
            {
                Id = item.Id,
                Subject = item.Subject,
                Start = item.Start,
                End = item.End,
                Location = item.Location,
                Organizer = CloneRecipient(item.Organizer),
                RequiredAttendees = CloneRecipients(item.RequiredAttendees),
                IsRecurring = item.IsRecurring,
                BusyStatus = item.BusyStatus,
                SmartOfficeOwned = item.SmartOfficeOwned,
                SmartOfficeEventId = item.SmartOfficeEventId,
            };
        }

        private static List<OutlookRecipientDto> CloneRecipients(List<OutlookRecipientDto> recipients)
        {
            return recipients.Select(CloneRecipient).ToList();
        }

        internal static OutlookRecipientDto CloneRecipient(OutlookRecipientDto recipient)
        {
            return new OutlookRecipientDto
            {
                RecipientKind = recipient.RecipientKind,
                DisplayName = recipient.DisplayName,
                SmtpAddress = recipient.SmtpAddress,
                RawAddress = recipient.RawAddress,
                AddressType = recipient.AddressType,
                EntryUserType = recipient.EntryUserType,
                IsGroup = recipient.IsGroup,
                IsResolved = recipient.IsResolved,
                Members = CloneRecipients(recipient.Members),
            };
        }
    }
}

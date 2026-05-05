using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public static class MockOutlookSeedFactory
    {
        public static MockOutlookSeedData Create()
        {
            var stores = new List<OutlookStoreDto>();
            var folders = new List<FolderDto>();
            var now = DateTime.Now;

            AddStore(stores, "mock-store-primary", "主要信箱 - Mock User", "ost", @"C:\Users\mock\AppData\Local\Microsoft\Outlook\mock.user@example.test.ost", MockOutlookPaths.PrimaryRoot);
            AddFolder(folders, "主要信箱 - Mock User", MockOutlookPaths.PrimaryRoot, "", "mock-store-primary", true);
            AddFolder(folders, "Inbox", MockOutlookPaths.Inbox, MockOutlookPaths.PrimaryRoot, "mock-store-primary");
            AddFolder(folders, "客戶專案", MockOutlookPaths.ClientProjects, MockOutlookPaths.Inbox, "mock-store-primary");
            AddFolder(folders, "Sent Items", MockOutlookPaths.Sent, MockOutlookPaths.PrimaryRoot, "mock-store-primary");
            AddFolder(folders, "Drafts", MockOutlookPaths.Drafts, MockOutlookPaths.PrimaryRoot, "mock-store-primary");
            AddFolder(folders, "Deleted Items", MockOutlookPaths.Deleted, MockOutlookPaths.PrimaryRoot, "mock-store-primary");

            AddStore(stores, "mock-store-ops", "營運共享信箱 - Mock Ops", "ost", @"C:\Users\mock\AppData\Local\Microsoft\Outlook\mock.ops@example.test.ost", MockOutlookPaths.OpsRoot);
            AddFolder(folders, "營運共享信箱 - Mock Ops", MockOutlookPaths.OpsRoot, "", "mock-store-ops", true);
            AddFolder(folders, "Ops Queue", MockOutlookPaths.OpsQueue, MockOutlookPaths.OpsRoot, "mock-store-ops");
            AddFolder(folders, "Escalations", MockOutlookPaths.OpsEscalations, MockOutlookPaths.OpsRoot, "mock-store-ops");

            AddStore(stores, "mock-store-legal-cold", "法務冷資料庫.pst", "pst", @"D:\Outlook Archives\法務冷資料庫.pst", MockOutlookPaths.LegalArchiveRoot);
            AddFolder(folders, "法務冷資料庫.pst", MockOutlookPaths.LegalArchiveRoot, "", "mock-store-legal-cold", true);
            AddFolder(folders, "合約保管庫", MockOutlookPaths.Archive, MockOutlookPaths.LegalArchiveRoot, "mock-store-legal-cold");
            AddFolder(folders, "簽核完成 2026", MockOutlookPaths.Archive2026, MockOutlookPaths.Archive, "mock-store-legal-cold");

            AddStore(stores, "mock-store-vendor-ledger", "供應商票據倉.pst", "pst", @"E:\MailBackup\供應商票據倉.pst", MockOutlookPaths.VendorArchiveRoot);
            AddFolder(folders, "供應商票據倉.pst", MockOutlookPaths.VendorArchiveRoot, "", "mock-store-vendor-ledger", true);
            AddFolder(folders, "票據入口", MockOutlookPaths.LegacyInbox, MockOutlookPaths.VendorArchiveRoot, "mock-store-vendor-ledger");
            AddFolder(folders, "供應商對帳", MockOutlookPaths.LegacyVendors, MockOutlookPaths.VendorArchiveRoot, "mock-store-vendor-ledger");

            var mails = new List<MailItemDto>
            {
                Mail("mock-001", "週會議程與客戶需求整理", "Ada Chen", "ada.chen@example.test", now.AddMinutes(-28), MockOutlookPaths.Inbox, false, "客戶,待辦", true, "today", "今天"),
                Mail("mock-002", "Re: 合約附件確認", "Ben Lin", "ben.lin@example.test", now.AddHours(-2), MockOutlookPaths.Inbox, true, "", false, "none", ""),
                Mail("mock-003", "Office 2016 add-in hover 測試", "QA Lab", "qa@example.test", now.AddHours(-4), MockOutlookPaths.Inbox, false, "測試", false, "none", "", bodyHtml: ""),
                Mail("mock-004", "下週 demo 時程", "Chris Wang", "chris.wang@example.test", now.AddDays(-1), MockOutlookPaths.Inbox, true, "追蹤", true, "next_week", "下週"),
                Mail("mock-005", "專案資料夾歸檔樣本", "Dana Hsu", "dana.hsu@example.test", now.AddDays(-2), MockOutlookPaths.ClientProjects, true, "客戶", false, "none", ""),
                Mail("mock-006", "已寄出的測試郵件", "Mock User", "mock.user@example.test", now.AddDays(-3), MockOutlookPaths.Sent, true, "", false, "none", ""),
                Mail("mock-007", "營運交接清單更新", "Ops Robot", "ops@example.test", now.AddHours(-6), MockOutlookPaths.OpsQueue, false, "測試", false, "none", ""),
                Mail("mock-008", "草稿：內部追蹤事項", "Mock User", "mock.user@example.test", now.AddDays(-10), MockOutlookPaths.Drafts, false, "待辦", true, "no_date", "Follow up"),
                Mail("mock-009", "法務封存：NDA 簽核完成", "Legal Desk", "legal@example.test", now.AddDays(-4), MockOutlookPaths.Archive2026, true, "追蹤", false, "none", ""),
                Mail("mock-010", "供應商票據差異表", "Vendor Team", "vendor@example.test", now.AddDays(-5), MockOutlookPaths.LegacyVendors, true, "待辦", true, "this_week", "本週"),
                Mail("mock-011", "票據入口：待補發票掃描", "Finance Bot", "finance@example.test", now.AddDays(-12), MockOutlookPaths.LegacyInbox, false, "", false, "none", ""),
                Mail("mock-012", "營運升級：夜間批次異常", "NOC", "noc@example.test", now.AddMinutes(-75), MockOutlookPaths.OpsEscalations, false, "測試,追蹤", true, "today", "今天"),
            };
            RefreshFolderCounts(folders, mails);

            return new MockOutlookSeedData(
                stores,
                folders,
                mails,
                new List<OutlookCategoryDto>
                {
                    new() { Name = "客戶", Color = "olCategoryColorBlue", ColorValue = 8, ShortcutKey = "" },
                    new() { Name = "待辦", Color = "olCategoryColorRed", ColorValue = 1, ShortcutKey = "" },
                    new() { Name = "測試", Color = "olCategoryColorGreen", ColorValue = 5, ShortcutKey = "" },
                    new() { Name = "追蹤", Color = "olCategoryColorYellow", ColorValue = 4, ShortcutKey = "" },
                },
                BuildRules(),
                BuildCalendar(now));
        }

        private static List<OutlookRuleDto> BuildRules()
        {
            return new List<OutlookRuleDto>
            {
                new()
                {
                    Name = "客戶郵件標記",
                    Enabled = true,
                    ExecutionOrder = 1,
                    Conditions = new List<string> { "sender contains example.test" },
                    Actions = new List<string> { "assign category 客戶" },
                },
                new()
                {
                    Name = "重要追蹤提醒",
                    Enabled = false,
                    ExecutionOrder = 2,
                    Conditions = new List<string> { "subject contains demo" },
                    Actions = new List<string> { "mark importance high", "flag for follow up" },
                    Exceptions = new List<string> { "sender is mock.user@example.test" },
                }
            };
        }

        private static List<CalendarEventDto> BuildCalendar(DateTime now)
        {
            return new List<CalendarEventDto>
            {
                new()
                {
                    Id = "mock-cal-001",
                    Subject = "SmartOffice mock sync review",
                    Start = now.Date.AddHours(15),
                    End = now.Date.AddHours(15).AddMinutes(30),
                    Location = "Teams",
                    Organizer = "mock.user@example.test",
                    RequiredAttendees = "ada.chen@example.test",
                    BusyStatus = "busy",
                },
                new()
                {
                    Id = "mock-cal-002",
                    Subject = "客戶需求釐清",
                    Start = now.Date.AddDays(2).AddHours(10),
                    End = now.Date.AddDays(2).AddHours(11),
                    Location = "會議室 3A",
                    Organizer = "ada.chen@example.test",
                    RequiredAttendees = "mock.user@example.test; dana.hsu@example.test",
                    BusyStatus = "tentative",
                },
                new()
                {
                    Id = "mock-cal-003",
                    Subject = "每週產品站會",
                    Start = now.Date.AddDays(6).AddHours(9),
                    End = now.Date.AddDays(6).AddHours(9).AddMinutes(45),
                    Location = "Teams",
                    Organizer = "mock.user@example.test",
                    RequiredAttendees = "product@example.test",
                    IsRecurring = true,
                    BusyStatus = "busy",
                },
                new()
                {
                    Id = "mock-cal-004",
                    Subject = "月中客戶回顧",
                    Start = now.Date.AddDays(14).AddHours(14),
                    End = now.Date.AddDays(14).AddHours(15),
                    Location = "Teams",
                    Organizer = "chris.wang@example.test",
                    RequiredAttendees = "mock.user@example.test; ada.chen@example.test",
                    BusyStatus = "busy",
                },
                new()
                {
                    Id = "mock-cal-005",
                    Subject = "月底交付檢查",
                    Start = now.Date.AddDays(24).AddHours(16),
                    End = now.Date.AddDays(24).AddHours(17),
                    Location = "會議室 2B",
                    Organizer = "mock.user@example.test",
                    RequiredAttendees = "qa@example.test",
                    BusyStatus = "free",
                }
            };
        }

        private static void AddStore(List<OutlookStoreDto> stores, string storeId, string displayName, string storeKind, string storeFilePath, string rootFolderPath)
        {
            stores.Add(new OutlookStoreDto
            {
                StoreId = storeId,
                DisplayName = displayName,
                StoreKind = storeKind,
                StoreFilePath = storeFilePath,
                RootFolderPath = rootFolderPath,
            });
        }

        private static void AddFolder(List<FolderDto> folders, string name, string folderPath, string parentFolderPath, string storeId, bool isStoreRoot = false)
        {
            folders.Add(new FolderDto
            {
                Name = name,
                EntryId = MockFolderEntryId(storeId, folderPath),
                FolderPath = folderPath,
                ParentEntryId = string.IsNullOrWhiteSpace(parentFolderPath) ? string.Empty : MockFolderEntryId(storeId, parentFolderPath),
                ParentFolderPath = parentFolderPath,
                StoreId = storeId,
                IsStoreRoot = isStoreRoot,
                DiscoveryState = "partial",
            });
        }

        private static string MockFolderEntryId(string storeId, string folderPath)
        {
            return $"{storeId}:{folderPath}";
        }

        private static MailItemDto Mail(
            string id,
            string subject,
            string senderName,
            string senderEmail,
            DateTime receivedTime,
            string folderPath,
            bool isRead,
            string categories,
            bool isMarkedAsTask,
            string flagInterval,
            string flagRequest,
            string? bodyHtml = null)
        {
            var body = $"Mock 郵件內容：{subject}\n\n這封郵件用於本機測試 Web UI、drag/drop 與 contract 行為。";
            var mail = new MailItemDto
            {
                Id = id,
                Subject = subject,
                SenderName = senderName,
                SenderEmail = senderEmail,
                ReceivedTime = receivedTime,
                Body = body,
                BodyHtml = bodyHtml ?? $"<article><h2>{subject}</h2><p>Mock 郵件內容，用於本機測試 Web UI 與 Outlook contract。</p></article>",
                FolderPath = folderPath,
                Categories = categories,
                IsRead = isRead,
                IsMarkedAsTask = isMarkedAsTask,
                FlagInterval = flagInterval,
                FlagRequest = flagRequest,
                TaskDueDate = isMarkedAsTask ? DateTime.Now.Date.AddDays(1) : null,
                Importance = isMarkedAsTask ? "high" : "normal",
                Sensitivity = "normal",
            };
            ApplyMockAttachmentSummary(mail);
            return mail;
        }

        private static void ApplyMockAttachmentSummary(MailItemDto mail)
        {
            var attachments = MockOutlookAttachmentFactory.Build(mail);
            mail.AttachmentCount = attachments.Count;
            mail.AttachmentNames = string.Join("、", attachments.Select(attachment => attachment.Name));
        }

        private static void RefreshFolderCounts(List<FolderDto> folders, List<MailItemDto> mails)
        {
            foreach (var folder in folders)
            {
                folder.ItemCount = mails.Count(mail => mail.FolderPath == folder.FolderPath);
                folder.HasChildren = folders.Any(child => child.ParentFolderPath == folder.FolderPath);
                folder.ChildrenLoaded = !folder.HasChildren;
                folder.DiscoveryState = folder.HasChildren ? "partial" : "loaded";
            }
        }
    }

    public record MockOutlookSeedData(
        List<OutlookStoreDto> Stores,
        List<FolderDto> Folders,
        List<MailItemDto> Mails,
        List<OutlookCategoryDto> Categories,
        List<OutlookRuleDto> Rules,
        List<CalendarEventDto> Calendar);
}

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
            AddFolder(folders, "Inbox", MockOutlookPaths.Inbox, MockOutlookPaths.PrimaryRoot, "mock-store-primary", folderType: OutlookFolderType.Inbox);
            AddFolder(folders, "客戶專案", MockOutlookPaths.ClientProjects, MockOutlookPaths.Inbox, "mock-store-primary");
            AddFolder(folders, "Sent Items", MockOutlookPaths.Sent, MockOutlookPaths.PrimaryRoot, "mock-store-primary", folderType: OutlookFolderType.Sent);
            AddFolder(folders, "Drafts", MockOutlookPaths.Drafts, MockOutlookPaths.PrimaryRoot, "mock-store-primary", folderType: OutlookFolderType.Drafts);
            AddFolder(folders, "Deleted Items", MockOutlookPaths.Deleted, MockOutlookPaths.PrimaryRoot, "mock-store-primary", folderType: OutlookFolderType.Deleted);

            AddStore(stores, "mock-store-ops", "營運共享信箱 - Mock Ops", "ost", @"C:\Users\mock\AppData\Local\Microsoft\Outlook\mock.ops@example.test.ost", MockOutlookPaths.OpsRoot);
            AddFolder(folders, "營運共享信箱 - Mock Ops", MockOutlookPaths.OpsRoot, "", "mock-store-ops", true);
            AddFolder(folders, "Ops Queue", MockOutlookPaths.OpsQueue, MockOutlookPaths.OpsRoot, "mock-store-ops");
            AddFolder(folders, "Escalations", MockOutlookPaths.OpsEscalations, MockOutlookPaths.OpsRoot, "mock-store-ops");
            AddFolder(folders, "Deleted Items", MockOutlookPaths.OpsDeleted, MockOutlookPaths.OpsRoot, "mock-store-ops", folderType: OutlookFolderType.Deleted);

            AddStore(stores, "mock-store-legal-cold", "法務冷資料庫.pst", "pst", @"D:\Outlook Archives\法務冷資料庫.pst", MockOutlookPaths.LegalArchiveRoot);
            AddFolder(folders, "法務冷資料庫.pst", MockOutlookPaths.LegalArchiveRoot, "", "mock-store-legal-cold", true);
            AddFolder(folders, "合約保管庫", MockOutlookPaths.Archive, MockOutlookPaths.LegalArchiveRoot, "mock-store-legal-cold");
            AddFolder(folders, "簽核完成 2026", MockOutlookPaths.Archive2026, MockOutlookPaths.Archive, "mock-store-legal-cold");
            AddFolder(folders, "Deleted Items", MockOutlookPaths.LegalDeleted, MockOutlookPaths.LegalArchiveRoot, "mock-store-legal-cold", folderType: OutlookFolderType.Deleted);

            AddStore(stores, "mock-store-vendor-ledger", "供應商票據倉.pst", "pst", @"E:\MailBackup\供應商票據倉.pst", MockOutlookPaths.VendorArchiveRoot);
            AddFolder(folders, "供應商票據倉.pst", MockOutlookPaths.VendorArchiveRoot, "", "mock-store-vendor-ledger", true);
            AddFolder(folders, "票據入口", MockOutlookPaths.LegacyInbox, MockOutlookPaths.VendorArchiveRoot, "mock-store-vendor-ledger");
            AddFolder(folders, "供應商對帳", MockOutlookPaths.LegacyVendors, MockOutlookPaths.VendorArchiveRoot, "mock-store-vendor-ledger");
            AddFolder(folders, "Deleted Items", MockOutlookPaths.VendorDeleted, MockOutlookPaths.VendorArchiveRoot, "mock-store-vendor-ledger", folderType: OutlookFolderType.Deleted);

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
            mails.AddRange(BuildInboxStressMails(now));
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

        private static List<MailItemDto> BuildInboxStressMails(DateTime now)
        {
            var subjects = new[]
            {
                "請確認：跨部門週報資料與待回覆事項",
                "Re: 客戶環境 Outlook 2016 測試結果",
                "FW: 採購單號與發票抬頭修正",
                "長標題測試：這封郵件刻意放入很長很長的主旨，用來確認列表截斷、hover、tag wrapping 與按鈕區塊不會互相覆蓋",
                "每日系統通知 - 無需回覆",
                "附件命名很長的報表請協助下載",
                "未讀郵件樣式與粗體 subject 檢查",
                "分類很多的測試郵件",
                "旗標：今天下班前回覆",
                "供應商來信：交期調整",
                "內部討論串第 4 封：保留 Re 前綴",
                "客戶專案風險清單更新",
                "測試空分類與一般重要性",
                "高優先度：現場問題回報",
                "會議記錄與行動項目整理",
                "請補充附件：報價單、合約、驗收截圖",
                "短主旨",
                "多收件人與群組顯示測試",
                "read/unread toggle 後列表位置檢查",
                "搜尋結果與 folder list 共用 row 樣式檢查",
                "邊界案例：沒有附件但是有很長的摘要文字",
                "客戶回覆：第二階段 PoC 時程",
                "營運窗口異動通知",
                "法務意見回覆",
                "付款條件確認",
                "正式版驗收前檢查清單",
                "Outlook cached mode 行為紀錄",
                "拖曳多封郵件測試資料",
                "刪除按鈕與開啟按鈕排列檢查",
                "列表最後一筆可視範圍檢查",
                "凌晨批次通知",
                "下午會議前提醒",
                "設計稿截圖與標註",
                "臨時插單：請評估影響",
                "客戶滿意度回饋",
                "封存前最後確認",
            };
            var senders = new[]
            {
                ("Ada Chen", "ada.chen@example.test"),
                ("Ben Lin", "ben.lin@example.test"),
                ("Chris Wang", "chris.wang@example.test"),
                ("Customer Success Team With A Long Display Name", "customer.success@example.test"),
                ("QA Lab", "qa@example.test"),
                ("Finance Bot", "finance@example.test"),
            };
            var categoryOptions = new[] { "", "客戶", "待辦", "測試", "追蹤", "客戶,待辦", "測試,追蹤", "客戶,追蹤,待辦" };
            var result = new List<MailItemDto>();

            for (var index = 0; index < subjects.Length; index++)
            {
                var sender = senders[index % senders.Length];
                var isMarkedAsTask = index % 5 == 0 || index % 11 == 0;
                var flagInterval = isMarkedAsTask ? (index % 2 == 0 ? "today" : "this_week") : "none";
                var flagRequest = isMarkedAsTask ? (flagInterval == "today" ? "今天" : "本週") : string.Empty;
                result.Add(Mail(
                    $"mock-inbox-{index + 100:000}",
                    subjects[index],
                    sender.Item1,
                    sender.Item2,
                    now.AddMinutes(-(95 + index * 37)),
                    MockOutlookPaths.Inbox,
                    isRead: index % 3 != 0,
                    categories: categoryOptions[index % categoryOptions.Length],
                    isMarkedAsTask: isMarkedAsTask,
                    flagInterval: flagInterval,
                    flagRequest: flagRequest));
            }

            return result;
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
            var today = now.Date;
            var events = new List<CalendarEventDto>
            {
                Calendar("mock-cal-001", "SmartOffice mock sync review", today.AddHours(15), 30, "Teams", "busy", Recipient("required", "Ada Chen", "ada.chen@example.test")),
                Calendar("mock-cal-002", "客戶需求釐清", today.AddDays(2).AddHours(10), 60, "會議室 3A", "tentative", Recipient("required", "Mock User", "mock.user@example.test"), Recipient("required", "Dana Hsu", "dana.hsu@example.test")),
                Calendar("mock-cal-003", "每週產品站會", today.AddDays(6).AddHours(9), 45, "Teams", "busy", new[] { Group("required", "Product Team", "product@example.test", "Ada Chen", "Ben Lin") }, true),
                Calendar("mock-cal-004", "月中客戶回顧", today.AddDays(14).AddHours(14), 60, "Teams", "busy", Recipient("required", "Mock User", "mock.user@example.test"), Recipient("required", "Ada Chen", "ada.chen@example.test")),
                Calendar("mock-cal-005", "月底交付檢查", today.AddDays(24).AddHours(16), 60, "會議室 2B", "free", Group("required", "QA Lab", "qa@example.test", "QA Lab Member")),
                Calendar("mock-cal-006", "跨日上線值班", today.AddDays(3).AddHours(22), 600, "War room", "busy", Recipient("required", "Mock User", "mock.user@example.test")),
                Calendar("mock-cal-007", "跨週專案封版", today.AddDays(10).AddHours(13), 2940, "Teams", "busy", Group("required", "Delivery Team", "delivery@example.test", "Mock User", "Ada Chen", "Chris Wang")),
            };

            var packedDay = today.AddDays(4);
            var packedSubjects = new[]
            {
                "晨間 triage",
                "需求 grooming",
                "供應商電話",
                "UI review",
                "合約條款確認",
                "客服升級案件",
                "資料彙整",
                "主管同步",
                "收尾檢查",
            };

            for (var i = 0; i < packedSubjects.Length; i++)
            {
                var start = packedDay.AddHours(8).AddMinutes(i * 50);
                events.Add(Calendar(
                    $"mock-cal-packed-{i + 1:00}",
                    packedSubjects[i],
                    start,
                    i % 3 == 0 ? 45 : 30,
                    i % 2 == 0 ? "Teams" : "會議室 5C",
                    i % 4 == 0 ? "tentative" : "busy",
                    Recipient("required", "Mock User", "mock.user@example.test"),
                    Recipient("required", "Ada Chen", "ada.chen@example.test")));
            }

            for (var dayOffset = -6; dayOffset <= 28; dayOffset += 3)
            {
                var start = today.AddDays(dayOffset).AddHours(9 + Math.Abs(dayOffset % 5));
                events.Add(Calendar(
                    $"mock-cal-rhythm-{dayOffset + 20:00}",
                    $"例行追蹤 {start:MM/dd}",
                    start,
                    25 + Math.Abs(dayOffset % 4) * 10,
                    dayOffset % 2 == 0 ? "Teams" : "會議室 2A",
                    dayOffset % 4 == 0 ? "free" : "busy",
                    Recipient("required", "Mock User", "mock.user@example.test")));
            }

            for (var i = 0; i < 8; i++)
            {
                var start = today.AddDays(12 + (i / 2)).AddHours(10 + (i % 2) * 4);
                events.Add(Calendar(
                    $"mock-cal-cluster-{i + 1:00}",
                    $"客戶專案 block {i + 1}",
                    start,
                    90,
                    i % 2 == 0 ? "客戶現場" : "Teams",
                    "busy",
                    Recipient("required", "Mock User", "mock.user@example.test"),
                    Recipient("required", "Dana Hsu", "dana.hsu@example.test")));
            }

            return events.OrderBy(item => item.Start).ToList();
        }

        private static CalendarEventDto Calendar(
            string id,
            string subject,
            DateTime start,
            int durationMinutes,
            string location,
            string busyStatus,
            OutlookRecipientDto requiredAttendee,
            params OutlookRecipientDto[] additionalRequiredAttendees)
        {
            return Calendar(id, subject, start, durationMinutes, location, busyStatus, new[] { requiredAttendee }.Concat(additionalRequiredAttendees), false);
        }

        private static CalendarEventDto Calendar(
            string id,
            string subject,
            DateTime start,
            int durationMinutes,
            string location,
            string busyStatus,
            IEnumerable<OutlookRecipientDto> requiredAttendees,
            bool recurring)
        {
            return new CalendarEventDto
            {
                Id = id,
                Subject = subject,
                Start = start,
                End = start.AddMinutes(durationMinutes),
                Location = location,
                Organizer = Recipient("organizer", "Mock User", "mock.user@example.test"),
                RequiredAttendees = requiredAttendees.ToList(),
                BusyStatus = busyStatus,
                IsRecurring = recurring,
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

        private static void AddFolder(
            List<FolderDto> folders,
            string name,
            string folderPath,
            string parentFolderPath,
            string storeId,
            bool isStoreRoot = false,
            OutlookFolderType folderType = OutlookFolderType.Mail)
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
                FolderType = isStoreRoot ? OutlookFolderType.StoreRoot : folderType,
                DefaultItemType = isStoreRoot ? -1 : 0,
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
            string senderDisplayName,
            string senderSmtpAddress,
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
                Sender = Recipient("sender", senderDisplayName, senderSmtpAddress),
                ToRecipients = new List<OutlookRecipientDto>
                {
                    id == "mock-001"
                        ? Group("to", "Product Team", "product@example.test", "Ada Chen", "Ben Lin", "Chris Wang")
                        : Recipient("to", "Mock User", "mock.user@example.test"),
                },
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

        private static OutlookRecipientDto Recipient(string kind, string displayName, string smtpAddress)
        {
            return new OutlookRecipientDto
            {
                RecipientKind = kind,
                DisplayName = displayName,
                SmtpAddress = smtpAddress,
                RawAddress = smtpAddress,
                AddressType = "SMTP",
                EntryUserType = "olExchangeUserAddressEntry",
                IsResolved = true,
            };
        }

        private static OutlookRecipientDto Group(string kind, string displayName, string smtpAddress, params string[] memberNames)
        {
            return new OutlookRecipientDto
            {
                RecipientKind = kind,
                DisplayName = displayName,
                SmtpAddress = smtpAddress,
                RawAddress = smtpAddress,
                AddressType = "SMTP",
                EntryUserType = "olExchangeDistributionListAddressEntry",
                IsGroup = true,
                IsResolved = true,
                Members = memberNames
                    .Select(name => Recipient("member", name, $"{name.ToLowerInvariant().Replace(" ", ".")}@example.test"))
                    .ToList(),
            };
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

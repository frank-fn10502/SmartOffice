using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public static class MockOutlookAttachmentFactory
    {
        public static List<MailAttachmentDto> Build(MailItemDto mail)
        {
            var attachments = mail.Id switch
            {
                var id when MockInboxSequence(id) is { } sequence && sequence % 7 == 0 => new[]
                {
                    ("客戶現場截圖與問題重現步驟-檔名刻意很長用來測試列表截斷.pdf", "application/pdf", 348_000L),
                    ("附件二-跨部門回覆彙整表.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 126_000L),
                    ("mail-thread-export.eml", "message/rfc822", 58_000L),
                },
                var id when MockInboxSequence(id) is { } sequence && sequence % 4 == 0 => new[]
                {
                    ("檢查清單.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 84_000L),
                    ("會議截圖.png", "image/png", 420_000L),
                },
                "mock-001" => new[]
                {
                    ("客戶需求摘要.pdf", "application/pdf", 128_000L),
                    ("demo-agenda.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", 256_000L),
                },
                "mock-002" => new[]
                {
                    ("合約附件.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", 96_000L),
                    ("報價明細.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 74_000L),
                },
                "mock-003" => new[]
                {
                    ("hover-test-cases.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 42_000L),
                },
                "mock-004" => new[]
                {
                    ("下週-demo-runbook.pdf", "application/pdf", 188_000L),
                    ("demo-assets.zip", "application/zip", 512_000L),
                },
                "mock-005" => new[]
                {
                    ("專案資料夾歸檔清單.csv", "text/csv", 18_000L),
                },
                "mock-006" => new[]
                {
                    ("已寄出附件範例.txt", "text/plain", 3_200L),
                },
                "mock-007" => new[]
                {
                    ("封存通知.eml", "message/rfc822", 64_000L),
                },
                "mock-008" => new[]
                {
                    ("草稿附件-placeholder.pdf", "application/pdf", 22_000L),
                },
                "mock-009" => new[]
                {
                    ("上月客戶回覆附件.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", 88_000L),
                },
                _ => Array.Empty<(string Name, string ContentType, long Size)>(),
            };

            return attachments
                .Select((attachment, index) => new MailAttachmentDto
                {
                    MailId = mail.Id,
                    AttachmentId = (index + 1).ToString(),
                    Name = attachment.Item1,
                    ContentType = attachment.Item2,
                    Size = attachment.Item3,
                })
                .ToList();
        }

        private static int? MockInboxSequence(string id)
        {
            const string prefix = "mock-inbox-";
            if (!id.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return null;
            return int.TryParse(id[prefix.Length..], out var sequence) ? sequence : null;
        }
    }
}

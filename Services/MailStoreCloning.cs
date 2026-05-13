using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public partial class MailStore
    {
        private static List<OutlookStoreDto> CloneStores(List<OutlookStoreDto> stores)
        {
            return stores.Select(CloneStore).ToList();
        }

        private static OutlookStoreDto CloneStore(OutlookStoreDto store)
        {
            return new OutlookStoreDto
            {
                StoreId = store.StoreId,
                DisplayName = store.DisplayName,
                StoreKind = store.StoreKind,
                StoreFilePath = store.StoreFilePath,
                RootFolderPath = store.RootFolderPath,
            };
        }

        private static List<FolderDto> CloneFolders(List<FolderDto> folders)
        {
            return folders.Select(CloneFolder).ToList();
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

        private static List<OutlookRecipientDto> CloneRecipients(List<OutlookRecipientDto> recipients)
        {
            return recipients.Select(CloneRecipient).ToList();
        }

        private static OutlookRecipientDto CloneRecipient(OutlookRecipientDto recipient)
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

        private static MailItemDto CloneMailMetadata(MailItemDto mail)
        {
            var clone = CloneMail(mail);
            clone.Body = string.Empty;
            clone.BodyHtml = string.Empty;
            return clone;
        }

        private static FolderDto CloneFolder(FolderDto folder)
        {
            return new FolderDto
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
            };
        }

        private static MailSearchProgressDto CloneMailSearchProgress(MailSearchProgressDto progress)
        {
            return new MailSearchProgressDto
            {
                SearchId = progress.SearchId,
                CommandId = progress.CommandId,
                Status = progress.Status,
                Phase = progress.Phase,
                ProcessedStores = progress.ProcessedStores,
                TotalStores = progress.TotalStores,
                ProcessedFolders = progress.ProcessedFolders,
                TotalFolders = progress.TotalFolders,
                ResultCount = progress.ResultCount,
                CurrentStoreId = progress.CurrentStoreId,
                CurrentFolderPath = progress.CurrentFolderPath,
                Message = progress.Message,
                Timestamp = progress.Timestamp,
            };
        }

        private static SearchMailsRequest CloneSearchMailsRequest(SearchMailsRequest request)
        {
            return new SearchMailsRequest
            {
                SearchId = request.SearchId,
                StoreId = request.StoreId,
                ScopeFolderPaths = new List<string>(request.ScopeFolderPaths),
                AllowGlobalScope = request.AllowGlobalScope,
                IncludeSubFolders = request.IncludeSubFolders,
                Keyword = request.Keyword,
                TextFields = new List<string>(request.TextFields),
                CategoryNames = new List<string>(request.CategoryNames),
                HasAttachments = request.HasAttachments,
                FlagState = request.FlagState,
                ReadState = request.ReadState,
                ReceivedFrom = request.ReceivedFrom,
                ReceivedTo = request.ReceivedTo,
            };
        }

        private static MailAttachmentsDto CloneMailAttachments(MailAttachmentsDto attachments)
        {
            return new MailAttachmentsDto
            {
                MailId = attachments.MailId,
                FolderPath = attachments.FolderPath,
                Attachments = attachments.Attachments.Select(CloneMailAttachment).ToList(),
            };
        }

        private static MailConversationDto CloneMailConversation(MailConversationDto conversation)
        {
            return new MailConversationDto
            {
                MailId = conversation.MailId,
                FolderPath = conversation.FolderPath,
                ConversationId = conversation.ConversationId,
                ConversationTopic = conversation.ConversationTopic,
                Mails = conversation.Mails.Select(CloneMail).ToList(),
            };
        }

        private static MailAttachmentDto CloneMailAttachment(MailAttachmentDto attachment)
        {
            return new MailAttachmentDto
            {
                MailId = attachment.MailId,
                Id = attachment.Id,
                AttachmentId = attachment.AttachmentId,
                Index = attachment.Index,
                FileName = attachment.FileName,
                DisplayName = attachment.DisplayName,
                Name = attachment.Name,
                ContentType = attachment.ContentType,
                Size = attachment.Size,
                IsExported = attachment.IsExported,
                ExportedAttachmentId = attachment.ExportedAttachmentId,
                Path = attachment.Path,
                LocalPath = attachment.LocalPath,
                FullPath = attachment.FullPath,
                ExportedPath = attachment.ExportedPath,
            };
        }

        private static ExportedMailAttachmentDto CloneExportedAttachment(ExportedMailAttachmentDto attachment)
        {
            return new ExportedMailAttachmentDto
            {
                MailId = attachment.MailId,
                FolderPath = attachment.FolderPath,
                Id = attachment.Id,
                AttachmentId = attachment.AttachmentId,
                Index = attachment.Index,
                ExportedAttachmentId = attachment.ExportedAttachmentId,
                FileName = attachment.FileName,
                DisplayName = attachment.DisplayName,
                Name = attachment.Name,
                ContentType = attachment.ContentType,
                Size = attachment.Size,
                Path = attachment.Path,
                LocalPath = attachment.LocalPath,
                FullPath = attachment.FullPath,
                ExportedPath = attachment.ExportedPath,
                ExportedAt = attachment.ExportedAt,
            };
        }
    }
}

using SmartOffice.Hub.Contracts;

namespace SmartOffice.Hub.Controllers
{
    public partial class OutlookController
    {
        private static object OperationAccepted(PendingCommand command, object? data = null, string? requestNameOverride = null)
        {
            return ResultEnvelope(
                command.Id,
                command.Type,
                "accepted",
                "Request accepted. Poll the paired fetch-result-* endpoint for state and data.",
                data,
                requestNameOverride);
        }

        private static object ResultEnvelope(string requestId, string commandType, string state, string message, object? data = null, string? requestNameOverride = null)
        {
            var requestName = string.IsNullOrWhiteSpace(requestNameOverride) ? RequestName(commandType) : requestNameOverride;
            return new
            {
                requestId,
                request = requestName,
                state,
                message,
                data = ResultDataWithFetchEndpoint(commandType, data, requestName),
            };
        }

        private static object ErrorEnvelope(string commandType, string status, string message, object? data = null)
        {
            return new
            {
                request = RequestName(commandType),
                status,
                state = "failed",
                message,
                data = data ?? new { },
            };
        }

        private static Dictionary<string, object?> ResultDataWithFetchEndpoint(string commandType, object? data, string? requestNameOverride = null)
        {
            var requestName = string.IsNullOrWhiteSpace(requestNameOverride) ? RequestName(commandType) : requestNameOverride;
            var result = new Dictionary<string, object?>
            {
                ["fetchResultEndpoint"] = $"/api/outlook/{requestName.Replace("request-", "fetch-result-")}",
            };
            if (data is null) return result;
            foreach (var property in data.GetType().GetProperties())
                result[property.Name] = property.GetValue(data);
            return result;
        }

        private static string ResultState(string status)
        {
            return status switch
            {
                "pending" => "running",
                "completed" or "mocked" => "completed",
                "outlook_unavailable" => "unavailable",
                "timeout" => "timeout",
                _ => "failed",
            };
        }

        private static string RequestName(string commandType)
        {
            return commandType switch
            {
                "fetch_folder_roots" => "request-folders",
                "fetch_folder_children" => "request-folder-children",
                "find_folder" => "request-find-folder",
                "fetch_mails" => "request-mails",
                "fetch_mail_body" => "request-mail-body",
                "fetch_mail_attachments" => "request-mail-attachments",
                "fetch_mail_conversation" => "request-mail-conversation",
                "export_mail_attachment" => "request-export-mail-attachment",
                "fetch_rules" => "request-rules",
                "manage_rule" => "request-manage-rule",
                "fetch_categories" => "request-categories",
                "ping" => "request-signalr-ping",
                "fetch_calendar" => "request-calendar",
                "fetch_calendar_rooms" => "request-calendar-rooms",
                "create_calendar_event" => "request-create-calendar-event",
                "update_calendar_event" => "request-update-calendar-event",
                "delete_calendar_event" => "request-delete-calendar-event",
                "fetch_address_book_roots" => "request-address-book-roots",
                "fetch_address_list_entries" => "request-address-list-entries",
                "fetch_address_book" => "request-address-book",
                "fetch_address_book_group_members" => "request-address-book-group-members",
                "address_book_relation_lookup" => "request-address-book-relation",
                "update_mail_properties" => "request-update-mail-properties",
                "upsert_category" => "request-upsert-category",
                "create_folder" => "request-create-folder",
                "delete_folder" => "request-delete-folder",
                "move_mail" => "request-move-mail",
                "move_mails" => "request-move-mails",
                "delete_mail" => "request-delete-mail",
                _ => commandType,
            };
        }
    }
}

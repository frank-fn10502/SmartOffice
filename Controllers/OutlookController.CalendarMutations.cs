using Microsoft.AspNetCore.Mvc;
using SmartOffice.Hub.Models;
using SmartOffice.Hub.Services;

namespace SmartOffice.Hub.Controllers
{
    public partial class OutlookController
    {
        [HttpPost("request-create-calendar-event")]
        public async Task<IActionResult> RequestCreateCalendarEvent([FromBody] CalendarEventCommandRequest req, CancellationToken ct)
        {
            var error = ValidateCalendarEventMutation(req, requireEventId: false);
            if (error is not null) return error;

            var cmd = new PendingCommand
            {
                Type = "create_calendar_event",
                CalendarEventRequest = NormalizeCalendarEventRequest(req, generateSmartOfficeEventId: true),
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        [HttpPost("request-calendar-rooms")]
        public async Task<IActionResult> RequestCalendarRooms(CancellationToken ct)
        {
            return await DispatchCommandAsync(new PendingCommand { Type = "fetch_calendar_rooms" }, ct);
        }

        [HttpPost("request-update-calendar-event")]
        public async Task<IActionResult> RequestUpdateCalendarEvent([FromBody] CalendarEventCommandRequest req, CancellationToken ct)
        {
            var error = ValidateCalendarEventMutation(req, requireEventId: true);
            if (error is not null) return error;

            if (IsKnownCalendarEventOwnershipMismatch(req.EventId, req.SmartOfficeEventId))
                return CalendarOwnershipError("update_calendar_event", req.EventId);

            var cmd = new PendingCommand
            {
                Type = "update_calendar_event",
                CalendarEventRequest = NormalizeCalendarEventRequest(req, generateSmartOfficeEventId: false),
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        [HttpPost("request-delete-calendar-event")]
        public async Task<IActionResult> RequestDeleteCalendarEvent([FromBody] CalendarEventCommandRequest req, CancellationToken ct)
        {
            var error = ApiRequestValidation.RequireFields(("eventId", req?.EventId), ("smartOfficeEventId", req?.SmartOfficeEventId));
            if (error is not null) return error;

            if (IsKnownCalendarEventOwnershipMismatch(req?.EventId, req?.SmartOfficeEventId))
                return CalendarOwnershipError("delete_calendar_event", req?.EventId ?? string.Empty);

            var cmd = new PendingCommand
            {
                Type = "delete_calendar_event",
                CalendarEventRequest = NormalizeCalendarEventRequest(req, generateSmartOfficeEventId: false),
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        [HttpPost("fetch-result-create-calendar-event")]
        public IActionResult FetchResultCreateCalendarEvent([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "create_calendar_event");
        }

        [HttpPost("fetch-result-calendar-rooms")]
        public IActionResult FetchResultCalendarRooms([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "fetch_calendar_rooms");
        }

        [HttpPost("fetch-result-update-calendar-event")]
        public IActionResult FetchResultUpdateCalendarEvent([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "update_calendar_event");
        }

        [HttpPost("fetch-result-delete-calendar-event")]
        public IActionResult FetchResultDeleteCalendarEvent([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "delete_calendar_event");
        }

        private IActionResult? ValidateCalendarEventMutation(CalendarEventCommandRequest? req, bool requireEventId)
        {
            var error = requireEventId
                ? ApiRequestValidation.RequireFields(("eventId", req?.EventId), ("smartOfficeEventId", req?.SmartOfficeEventId), ("subject", req?.Subject))
                : ApiRequestValidation.RequireFields(("subject", req?.Subject));
            if (error is not null) return error;

            if (req?.Start is null || req.End is null)
                return ApiRequestValidation.MissingRequiredFields("start", "end");

            if (req.Start >= req.End)
            {
                return BadRequest(new
                {
                    request = RequestName(requireEventId ? "update_calendar_event" : "create_calendar_event"),
                    status = "invalid_calendar_range",
                    state = "failed",
                    message = "start must be earlier than end.",
                    data = new { },
                });
            }

            return null;
        }

        private CalendarEventCommandRequest NormalizeCalendarEventRequest(CalendarEventCommandRequest? req, bool generateSmartOfficeEventId)
        {
            req ??= new CalendarEventCommandRequest();
            req.Start = UtcDateTime.Normalize(req.Start);
            req.End = UtcDateTime.Normalize(req.End);
            req.RequiredAttendees ??= new List<OutlookRecipientDto>();
            req.Resources ??= new List<OutlookRecipientDto>();
            req.SmartOfficeEventId = string.IsNullOrWhiteSpace(req.SmartOfficeEventId)
                ? (generateSmartOfficeEventId ? Guid.NewGuid().ToString() : string.Empty)
                : req.SmartOfficeEventId.Trim();
            return req;
        }

        private bool IsKnownCalendarEventOwnershipMismatch(string? eventId, string? smartOfficeEventId)
        {
            if (string.IsNullOrWhiteSpace(eventId)) return false;
            var known = _mailStore.GetCalendarEvents().FirstOrDefault(item =>
                string.Equals(item.Id, eventId, StringComparison.OrdinalIgnoreCase));
            if (known is null) return false;
            if (!known.SmartOfficeOwned) return true;
            return string.IsNullOrWhiteSpace(smartOfficeEventId)
                || !string.Equals(known.SmartOfficeEventId, smartOfficeEventId.Trim(), StringComparison.Ordinal);
        }

        private IActionResult CalendarOwnershipError(string commandType, string eventId)
        {
            return Conflict(new
            {
                request = RequestName(commandType),
                status = "not_smartoffice_owned",
                state = "failed",
                message = "SmartOffice 只能更新或刪除由 SmartOffice 建立且 smartOfficeEventId 相符的 calendar event。",
                eventId,
                data = new { },
            });
        }
    }
}

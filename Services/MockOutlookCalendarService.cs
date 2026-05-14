namespace SmartOffice.Hub.Services
{
    internal sealed class MockOutlookCalendarService
    {
        private readonly List<CalendarEventDto> _events;
        private readonly List<CalendarRoomDto> _rooms;

        public MockOutlookCalendarService(List<CalendarEventDto> events)
        {
            _events = events;
            _rooms = BuildRooms();
        }

        public List<CalendarEventDto> AllEvents() => _events.Select(MockOutlookService.CloneCalendarEvent).ToList();

        public List<CalendarRoomDto> Rooms() => _rooms.Select(CloneRoom).ToList();

        public List<CalendarEventDto> Filter(FetchCalendarRequest? request)
        {
            var start = request?.StartDate?.Date ?? DateTime.Now.Date;
            var end = request?.EndDate?.Date ?? start.AddDays(Math.Max(1, request?.DaysForward ?? 31));
            return _events
                .Where(item => item.Start >= start && item.Start < end)
                .OrderBy(item => item.Start)
                .Select(MockOutlookService.CloneCalendarEvent)
                .ToList();
        }

        public List<CalendarEventDto> Create(CalendarEventCommandRequest? request)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.Subject) || request.Start is null || request.End is null)
                return Filter(new FetchCalendarRequest());

            _events.Add(new CalendarEventDto
            {
                Id = $"mock-smartoffice-{Guid.NewGuid():N}",
                SmartOfficeEventId = string.IsNullOrWhiteSpace(request.SmartOfficeEventId) ? Guid.NewGuid().ToString() : request.SmartOfficeEventId,
                SmartOfficeOwned = true,
                Subject = request.Subject.Trim(),
                Start = UtcDateTime.Normalize(request.Start) ?? DateTime.UtcNow,
                End = UtcDateTime.Normalize(request.End) ?? DateTime.UtcNow.AddHours(1),
                Location = request.Location.Trim(),
                BusyStatus = string.IsNullOrWhiteSpace(request.BusyStatus) ? "busy" : request.BusyStatus.Trim(),
                Organizer = MockOrganizer(),
                RequiredAttendees = request.RequiredAttendees
                    .Concat(request.Resources.Select(ResourceRecipient))
                    .Select(MockOutlookService.CloneRecipient)
                    .ToList(),
            });
            return CurrentWindow();
        }

        public List<CalendarEventDto>? Update(CalendarEventCommandRequest? request)
        {
            var item = FindSmartOfficeEvent(request);
            if (item is null) return null;

            item.Subject = request!.Subject.Trim();
            item.Start = UtcDateTime.Normalize(request.Start) ?? item.Start;
            item.End = UtcDateTime.Normalize(request.End) ?? item.End;
            item.Location = request.Location.Trim();
            item.BusyStatus = string.IsNullOrWhiteSpace(request.BusyStatus) ? item.BusyStatus : request.BusyStatus.Trim();
            item.RequiredAttendees = request.RequiredAttendees
                .Concat(request.Resources.Select(ResourceRecipient))
                .Select(MockOutlookService.CloneRecipient)
                .ToList();
            return CurrentWindow();
        }

        public List<CalendarEventDto>? Delete(CalendarEventCommandRequest? request)
        {
            var item = FindSmartOfficeEvent(request);
            if (item is null) return null;
            _events.Remove(item);
            return CurrentWindow();
        }

        private CalendarEventDto? FindSmartOfficeEvent(CalendarEventCommandRequest? request)
        {
            var eventId = request?.EventId;
            var smartOfficeEventId = request?.SmartOfficeEventId;
            if (string.IsNullOrWhiteSpace(eventId)) return null;
            if (string.IsNullOrWhiteSpace(smartOfficeEventId)) return null;
            return _events.FirstOrDefault(item =>
                string.Equals(item.Id, eventId, StringComparison.OrdinalIgnoreCase)
                && item.SmartOfficeOwned
                && string.Equals(item.SmartOfficeEventId, smartOfficeEventId.Trim(), StringComparison.Ordinal));
        }

        private List<CalendarEventDto> CurrentWindow()
        {
            var now = DateTime.Now.Date;
            return _events
                .Where(item => item.Start >= now.AddDays(-31) && item.Start < now.AddDays(62))
                .OrderBy(item => item.Start)
                .Select(MockOutlookService.CloneCalendarEvent)
                .ToList();
        }

        private static List<CalendarRoomDto> BuildRooms()
        {
            return new List<CalendarRoomDto>
            {
                Room("mock-room-3a", "會議室 3A", "room-3a@example.test"),
                Room("mock-room-5c", "會議室 5C", "room-5c@example.test"),
                Room("mock-room-2b", "會議室 2B", "room-2b@example.test"),
                Room("mock-room-war", "War room", "war-room@example.test"),
            };
        }

        private static CalendarRoomDto Room(string id, string name, string smtp) => new()
        {
            Id = id,
            DisplayName = name,
            SmtpAddress = smtp,
            RawAddress = smtp,
            Source = "Mock room list",
        };

        private static CalendarRoomDto CloneRoom(CalendarRoomDto room) => new()
        {
            Id = room.Id,
            DisplayName = room.DisplayName,
            SmtpAddress = room.SmtpAddress,
            RawAddress = room.RawAddress,
            Source = room.Source,
        };

        private static OutlookRecipientDto MockOrganizer() => new()
        {
            RecipientKind = "organizer",
            DisplayName = "Mock User",
            SmtpAddress = "mock.user@example.test",
            RawAddress = "mock.user@example.test",
            AddressType = "SMTP",
            EntryUserType = "olExchangeUserAddressEntry",
            IsResolved = true,
        };

        private static OutlookRecipientDto ResourceRecipient(OutlookRecipientDto recipient)
        {
            recipient.RecipientKind = "resource";
            return recipient;
        }
    }
}

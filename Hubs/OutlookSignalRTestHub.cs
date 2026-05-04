using Microsoft.AspNetCore.SignalR;

namespace SmartOffice.Hub.Hubs
{
    /// <summary>
    /// 測試用 Outlook AddIn SignalR channel。
    /// 這個 hub 不參與目前 production Outlook polling protocol，只用來驗證
    /// 工作機 AddIn 是否能連線、接收 command、回報 message/result。
    /// </summary>
    public class OutlookSignalRTestHub : Microsoft.AspNetCore.SignalR.Hub
    {
        private const string AddinGroup = "outlook-addin-test";

        public async Task RegisterOutlookAddinTest(OutlookSignalRTestClientInfo info)
        {
            await Groups.AddToGroupAsync(Context.ConnectionId, AddinGroup);
            await Clients.All.SendAsync("OutlookSignalRTestAddinConnected", new OutlookSignalRTestConnectionEvent
            {
                ConnectionId = Context.ConnectionId,
                ClientName = string.IsNullOrWhiteSpace(info.ClientName) ? "Outlook AddIn" : info.ClientName,
                Workstation = info.Workstation,
                Version = info.Version,
                Timestamp = DateTime.Now
            });
        }

        public async Task SendOutlookSignalRTestCommand(OutlookSignalRTestCommand command)
        {
            if (string.IsNullOrWhiteSpace(command.Id))
                command.Id = Guid.NewGuid().ToString();
            if (command.CreatedAt == default)
                command.CreatedAt = DateTime.Now;

            await Clients.Group(AddinGroup).SendAsync("OutlookSignalRTestCommand", command);
            await Clients.All.SendAsync("OutlookSignalRTestCommandDispatched", command);
        }

        public async Task ReportOutlookSignalRTestMessage(OutlookSignalRTestMessage message)
        {
            message.ConnectionId = Context.ConnectionId;
            if (message.Timestamp == default)
                message.Timestamp = DateTime.Now;

            await Clients.All.SendAsync("OutlookSignalRTestMessage", message);
        }

        public async Task ReportOutlookSignalRTestResult(OutlookSignalRTestResult result)
        {
            result.ConnectionId = Context.ConnectionId;
            if (result.Timestamp == default)
                result.Timestamp = DateTime.Now;

            await Clients.All.SendAsync("OutlookSignalRTestResult", result);
        }

        public override async Task OnDisconnectedAsync(Exception? exception)
        {
            await Clients.All.SendAsync("OutlookSignalRTestAddinDisconnected", new OutlookSignalRTestConnectionEvent
            {
                ConnectionId = Context.ConnectionId,
                ClientName = "Outlook AddIn",
                Timestamp = DateTime.Now
            });
            await base.OnDisconnectedAsync(exception);
        }
    }

    public class OutlookSignalRTestClientInfo
    {
        public string ClientName { get; set; } = string.Empty;
        public string Workstation { get; set; } = string.Empty;
        public string Version { get; set; } = string.Empty;
    }

    public class OutlookSignalRTestCommand
    {
        public string Id { get; set; } = string.Empty;
        public string Type { get; set; } = "ping";
        public string Payload { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; } = DateTime.Now;
    }

    public class OutlookSignalRTestMessage
    {
        public string ConnectionId { get; set; } = string.Empty;
        public string Source { get; set; } = "addin";
        public string Level { get; set; } = "info";
        public string Text { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }

    public class OutlookSignalRTestResult
    {
        public string ConnectionId { get; set; } = string.Empty;
        public string CommandId { get; set; } = string.Empty;
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string Payload { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }

    public class OutlookSignalRTestConnectionEvent
    {
        public string ConnectionId { get; set; } = string.Empty;
        public string ClientName { get; set; } = string.Empty;
        public string Workstation { get; set; } = string.Empty;
        public string Version { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }
}

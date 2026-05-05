using Microsoft.AspNetCore.SignalR;
using SmartOffice.Hub.Hubs;
using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public class OutlookCommandQueue
    {
        private static readonly TimeSpan ConnectionTimeout = TimeSpan.FromSeconds(45);
        private static readonly TimeSpan PingTimeout = TimeSpan.FromSeconds(20);
        private static readonly TimeSpan CommandTimeout = TimeSpan.FromSeconds(90);
        private static readonly TimeSpan ReadyFreshness = TimeSpan.FromSeconds(20);

        private readonly SemaphoreSlim _queue = new(1, 1);
        private readonly MailStore _mailStore;
        private readonly CommandResultStore _commandResults;
        private readonly AddinStatusStore _addinStatus;
        private readonly OutlookSignalRCommandDispatcher _commandDispatcher;
        private readonly MockOutlookService _mockOutlook;
        private readonly IHubContext<NotificationHub> _notifications;
        private DateTime? _lastReadyAt;

        public OutlookCommandQueue(
            MailStore mailStore,
            CommandResultStore commandResults,
            AddinStatusStore addinStatus,
            OutlookSignalRCommandDispatcher commandDispatcher,
            MockOutlookService mockOutlook,
            IHubContext<NotificationHub> notifications)
        {
            _mailStore = mailStore;
            _commandResults = commandResults;
            _addinStatus = addinStatus;
            _commandDispatcher = commandDispatcher;
            _mockOutlook = mockOutlook;
            _notifications = notifications;
        }

        public async Task<OutlookQueuedCommandResult> ExecuteAsync(
            PendingCommand command,
            Func<bool>? dataReady = null,
            bool ensureReady = true,
            CancellationToken ct = default)
        {
            await _queue.WaitAsync(ct);
            try
            {
                return await ExecuteQueuedCommandAsync(command, dataReady, ensureReady, ct);
            }
            finally
            {
                _queue.Release();
            }
        }

        public async Task<T> ExecuteExclusiveAsync<T>(Func<CancellationToken, Task<T>> operation, CancellationToken ct = default)
        {
            await _queue.WaitAsync(ct);
            try
            {
                return await operation(ct);
            }
            finally
            {
                _queue.Release();
            }
        }

        public async Task<OutlookQueuedCommandResult> ExecuteQueuedCommandAsync(
            PendingCommand command,
            Func<bool>? dataReady = null,
            bool ensureReady = true,
            CancellationToken ct = default)
        {
            if (ensureReady && command.Type != "ping")
            {
                var ready = await EnsureOutlookReadyAsync(command.Type, ct);
                if (!ready.Success) return ready;
            }

            return await DispatchAndWaitAsync(command, dataReady, CommandTimeout, ct);
        }

        private async Task<OutlookQueuedCommandResult> EnsureOutlookReadyAsync(string commandType, CancellationToken ct)
        {
            if (_mockOutlook.IsEnabled) return OutlookQueuedCommandResult.Completed(string.Empty, "mock_ready", "Mock Outlook backend is ready.");
            if (_lastReadyAt.HasValue && DateTime.Now - _lastReadyAt.Value < ReadyFreshness)
                return OutlookQueuedCommandResult.Completed(string.Empty, "ready", "Outlook AddIn was recently verified.");

            if (!await WaitForSignalRConnectionAsync(ConnectionTimeout, ct))
                return OutlookQueuedCommandResult.AddinUnavailable("No Outlook AddIn SignalR connection is available.");

            var ping = await DispatchAndWaitAsync(new PendingCommand { Type = "ping" }, null, PingTimeout, ct);
            if (!ping.Success) return ping;

            _lastReadyAt = DateTime.Now;

            return OutlookQueuedCommandResult.Completed(string.Empty, "ready", "Outlook AddIn is ready.");
        }

        private async Task<bool> WaitForSignalRConnectionAsync(TimeSpan timeout, CancellationToken ct)
        {
            var deadline = DateTime.Now.Add(timeout);
            while (DateTime.Now < deadline && !ct.IsCancellationRequested)
            {
                if (_addinStatus.HasSignalRConnection()) return true;
                await Task.Delay(250, ct);
            }

            return _addinStatus.HasSignalRConnection();
        }

        private async Task<OutlookQueuedCommandResult> DispatchAndWaitAsync(
            PendingCommand command,
            Func<bool>? dataReady,
            TimeSpan timeout,
            CancellationToken ct)
        {
            _commandResults.RecordDispatched(command);

            if (await _mockOutlook.TryDispatchAsync(command, ct))
            {
                var result = new OutlookCommandResult
                {
                    CommandId = command.Id,
                    Success = true,
                    Message = $"{command.Type} completed by mock backend",
                    Timestamp = DateTime.Now,
                };
                _commandResults.RecordResult(result);
                await _notifications.Clients.All.SendAsync("CommandResult", result, ct);
                return OutlookQueuedCommandResult.Completed(command.Id, "mocked", result.Message);
            }

            if (!await _commandDispatcher.DispatchAsync(command, ct))
            {
                _commandResults.RecordUnavailable(command);
                return OutlookQueuedCommandResult.AddinUnavailable(command.Id, "No Outlook AddIn SignalR connection is available.");
            }

            var deadline = DateTime.Now.Add(timeout);
            while (DateTime.Now < deadline && !ct.IsCancellationRequested)
            {
                var status = _commandResults.Get(command.Id);
                if (status is not null && status.Status != "pending")
                {
                    if (status.Status == "completed")
                    {
                        if (dataReady is not null && !dataReady.Invoke())
                            return OutlookQueuedCommandResult.Failed(command.Id, "cache_not_ready", $"{command.Type} completed but the expected cache is not ready.");

                        return OutlookQueuedCommandResult.Completed(command.Id, "completed", status.Message);
                    }

                    return status.Status == "addin_unavailable"
                        ? OutlookQueuedCommandResult.AddinUnavailable(command.Id, status.Message)
                        : OutlookQueuedCommandResult.Failed(command.Id, status.Status, status.Message);
                }

                await Task.Delay(250, ct);
            }

            var timeoutResult = new OutlookCommandResult
            {
                CommandId = command.Id,
                Success = false,
                Message = $"{command.Type} timed out waiting for Outlook AddIn result.",
                Timestamp = DateTime.Now,
            };
            _commandResults.RecordResult(timeoutResult);
            await _notifications.Clients.All.SendAsync("CommandResult", timeoutResult, CancellationToken.None);
            return OutlookQueuedCommandResult.TimedOut(command.Id, timeoutResult.Message);
        }
    }

    public class OutlookQueuedCommandResult
    {
        public string CommandId { get; init; } = string.Empty;
        public string Status { get; init; } = string.Empty;
        public string Message { get; init; } = string.Empty;
        public bool Success { get; init; }
        public int HttpStatusCode { get; init; } = StatusCodes.Status200OK;

        public static OutlookQueuedCommandResult Completed(string commandId, string status, string message)
        {
            return new OutlookQueuedCommandResult
            {
                CommandId = commandId,
                Status = status,
                Message = message,
                Success = true,
            };
        }

        public static OutlookQueuedCommandResult AddinUnavailable(string message)
        {
            return AddinUnavailable(string.Empty, message);
        }

        public static OutlookQueuedCommandResult AddinUnavailable(string commandId, string message)
        {
            return new OutlookQueuedCommandResult
            {
                CommandId = commandId,
                Status = "addin_unavailable",
                Message = message,
                Success = false,
                HttpStatusCode = StatusCodes.Status409Conflict,
            };
        }

        public static OutlookQueuedCommandResult Failed(string commandId, string status, string message)
        {
            return new OutlookQueuedCommandResult
            {
                CommandId = commandId,
                Status = string.IsNullOrWhiteSpace(status) ? "failed" : status,
                Message = message,
                Success = false,
                HttpStatusCode = StatusCodes.Status502BadGateway,
            };
        }

        public static OutlookQueuedCommandResult TimedOut(string commandId, string message)
        {
            return new OutlookQueuedCommandResult
            {
                CommandId = commandId,
                Status = "timeout",
                Message = message,
                Success = false,
                HttpStatusCode = StatusCodes.Status504GatewayTimeout,
            };
        }
    }
}

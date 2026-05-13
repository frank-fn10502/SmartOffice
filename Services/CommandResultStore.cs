using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public class CommandResultStore
    {
        private const int MaxCommands = 500;
        private readonly object _lock = new();
        private readonly Dictionary<string, OutlookCommandStatusDto> _commands = new();
        private readonly Dictionary<string, PendingCommand> _requestCommands = new();
        private readonly Queue<string> _order = new();

        public void RecordDispatched(PendingCommand command)
        {
            lock (_lock)
            {
                if (!_commands.ContainsKey(command.Id))
                    _order.Enqueue(command.Id);

                _commands[command.Id] = new OutlookCommandStatusDto
                {
                    CommandId = command.Id,
                    Type = command.Type,
                    Status = "pending",
                    DispatchTimestamp = DateTime.Now,
                };
                _requestCommands[command.Id] = command;

                TrimIfNeeded();
            }
        }

        public void RecordUnavailable(PendingCommand command)
        {
            lock (_lock)
            {
                if (!_commands.TryGetValue(command.Id, out var status))
                {
                    status = new OutlookCommandStatusDto
                    {
                        CommandId = command.Id,
                        Type = command.Type,
                        DispatchTimestamp = DateTime.Now,
                    };
                    _commands[command.Id] = status;
                    _order.Enqueue(command.Id);
                }
                _requestCommands[command.Id] = command;

                status.Status = "outlook_unavailable";
                status.Success = false;
                status.Message = "Outlook request executor is not available.";
                status.ResultTimestamp = DateTime.Now;
                TrimIfNeeded();
            }
        }

        public void RecordResult(OutlookCommandResult result)
        {
            lock (_lock)
            {
                if (!_commands.TryGetValue(result.CommandId, out var status))
                {
                    status = new OutlookCommandStatusDto
                    {
                        CommandId = result.CommandId,
                        DispatchTimestamp = result.Timestamp,
                    };
                    _commands[result.CommandId] = status;
                    _order.Enqueue(result.CommandId);
                }

                status.Status = result.Success ? "completed" : "failed";
                status.Success = result.Success;
                status.Message = result.Message;
                status.Payload = result.Payload;
                status.ResultTimestamp = result.Timestamp;
                TrimIfNeeded();
            }
        }

        public OutlookCommandStatusDto? Get(string commandId)
        {
            lock (_lock)
            {
                if (!_commands.TryGetValue(commandId, out var status))
                    return null;

                return Clone(status);
            }
        }

        public PendingCommand? GetRequestCommand(string requestId)
        {
            lock (_lock)
            {
                return _requestCommands.TryGetValue(requestId, out var command) ? command : null;
            }
        }

        public List<OutlookCommandStatusDto> GetRecent()
        {
            lock (_lock)
            {
                return _order
                    .Where(_commands.ContainsKey)
                    .Select(id => Clone(_commands[id]))
                    .Reverse()
                    .ToList();
            }
        }

        private void TrimIfNeeded()
        {
            while (_order.Count > MaxCommands)
            {
                var oldestId = _order.Dequeue();
                _commands.Remove(oldestId);
                _requestCommands.Remove(oldestId);
            }
        }

        private static OutlookCommandStatusDto Clone(OutlookCommandStatusDto status)
        {
            return new OutlookCommandStatusDto
            {
                CommandId = status.CommandId,
                Type = status.Type,
                Status = status.Status,
                Success = status.Success,
                Message = status.Message,
                Payload = status.Payload,
                DispatchTimestamp = status.DispatchTimestamp,
                ResultTimestamp = status.ResultTimestamp,
            };
        }
    }
}

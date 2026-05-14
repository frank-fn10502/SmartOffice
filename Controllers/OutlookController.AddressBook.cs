using Microsoft.AspNetCore.Mvc;
using SmartOffice.Hub.Contracts;

namespace SmartOffice.Hub.Controllers
{
    public partial class OutlookController
    {
        /// <summary>
        /// Web UI、AI 或 MCP client 要求 Outlook address book。
        /// </summary>
        [HttpPost("request-address-book")]
        public async Task<IActionResult> RequestAddressBook([FromBody] AddressBookSyncRequest? req, CancellationToken ct)
        {
            req ??= new AddressBookSyncRequest();
            req.AddressListId = req.AddressListId?.Trim() ?? string.Empty;
            req.AddressListName = req.AddressListName?.Trim() ?? string.Empty;
            req.GroupId = req.GroupId?.Trim() ?? string.Empty;
            req.GroupSmtpAddress = req.GroupSmtpAddress?.Trim() ?? string.Empty;
            req.Offset = Math.Max(0, req.Offset);
            req.PageSize = Math.Clamp(req.PageSize <= 0 ? 100 : req.PageSize, 1, 500);

            if (!string.IsNullOrWhiteSpace(req.GroupId) || !string.IsNullOrWhiteSpace(req.GroupSmtpAddress))
                return await DispatchAddressBookGroupMembers(req, ct);

            if (!string.IsNullOrWhiteSpace(req.AddressListId) || !string.IsNullOrWhiteSpace(req.AddressListName))
                return await DispatchAddressBookListEntries(req, ct);

            return await DispatchCommandAsync(new PendingCommand { Type = "fetch_address_book_roots" }, ct, "request-address-book");
        }

        [HttpPost("request-address-book-relation")]
        public IActionResult RequestAddressBookRelation([FromBody] AddressBookRelationLookupRequest? req)
        {
            req ??= new AddressBookRelationLookupRequest();
            req.Query = req.Query?.Trim() ?? string.Empty;
            req.TargetKind = req.TargetKind?.Trim() ?? string.Empty;
            req.Id = req.Id?.Trim() ?? string.Empty;
            req.DisplayName = req.DisplayName?.Trim() ?? string.Empty;
            req.SmtpAddress = req.SmtpAddress?.Trim() ?? string.Empty;
            req.Email = req.Email?.Trim() ?? string.Empty;
            req.GroupId = req.GroupId?.Trim() ?? string.Empty;
            req.GroupSmtpAddress = req.GroupSmtpAddress?.Trim() ?? string.Empty;
            req.Take = Math.Clamp(req.Take <= 0 ? 50 : req.Take, 1, 500);

            if (string.IsNullOrWhiteSpace(req.Query)
                && string.IsNullOrWhiteSpace(req.Id)
                && string.IsNullOrWhiteSpace(req.DisplayName)
                && string.IsNullOrWhiteSpace(req.SmtpAddress)
                && string.IsNullOrWhiteSpace(req.Email)
                && string.IsNullOrWhiteSpace(req.GroupId)
                && string.IsNullOrWhiteSpace(req.GroupSmtpAddress))
            {
                return BadRequest(ErrorEnvelope(
                    "address_book_relation_lookup",
                    "missing_required_fields",
                    "Missing required request field(s): query, email, smtpAddress, id, groupSmtpAddress, or groupId."));
            }

            var command = new PendingCommand
            {
                Type = "address_book_relation_lookup",
                AddressBookRelationLookupRequest = req,
            };
            _commandResults.RecordDispatched(command);

            var current = _mailStore.GetAddressBookRelationLookup(req);
            if (!ShouldFetchAddressBookRelation(current))
            {
                _commandResults.RecordResult(new OutlookCommandResult
                {
                    CommandId = command.Id,
                    Success = true,
                    Message = "Address book relation result is ready.",
                    Timestamp = DateTime.Now,
                });
                return Ok(OperationAccepted(command));
            }

            _ = Task.Run(() => _commandQueue.ExecuteExclusiveAsync(
                operationCt => RequestAddressBookRelationQueuedAsync(command, operationCt),
                CancellationToken.None));
            return Ok(OperationAccepted(command));
        }

        private async Task<bool> RequestAddressBookRelationQueuedAsync(PendingCommand command, CancellationToken ct)
        {
            var result = await _commandQueue.ExecuteQueuedCommandAsync(command, null, ensureReady: true, ct: ct);
            if (!result.Success)
            {
                _commandResults.RecordResult(new OutlookCommandResult
                {
                    CommandId = command.Id,
                    Success = false,
                    Message = result.Message,
                    Timestamp = DateTime.Now,
                });
            }

            return result.Success;
        }

        private static bool ShouldFetchAddressBookRelation(AddressBookRelationLookupResponse current)
        {
            if (current.State == "not_found") return true;
            if (current.State == "ambiguous") return false;
            if (current.Target is null) return true;
            if (current.IsGroup)
                return !current.GroupMembersLoaded && !current.GroupMembersLoading;
            return current.MemberOfGroups.Count == 0 && current.ContainingGroups.Count == 0;
        }

        [HttpPost("request-address-book-roots")]
        public async Task<IActionResult> RequestAddressBookRoots(CancellationToken ct)
        {
            return await DispatchCommandAsync(new PendingCommand { Type = "fetch_address_book_roots" }, ct);
        }

        [HttpPost("request-address-list-entries")]
        public async Task<IActionResult> RequestAddressListEntries([FromBody] AddressBookListEntriesRequest? req, CancellationToken ct)
        {
            req ??= new AddressBookListEntriesRequest();
            req.AddressListId = req.AddressListId?.Trim() ?? string.Empty;
            req.AddressListName = req.AddressListName?.Trim() ?? string.Empty;
            req.Offset = Math.Max(0, req.Offset);
            req.PageSize = Math.Clamp(req.PageSize <= 0 ? 100 : req.PageSize, 1, 500);
            if (string.IsNullOrWhiteSpace(req.AddressListId) && string.IsNullOrWhiteSpace(req.AddressListName))
                return BadRequest(ErrorEnvelope("fetch_address_list_entries", "missing_required_fields", "Missing required request field(s): addressListId or addressListName."));

            return await DispatchCommandAsync(new PendingCommand
            {
                Type = "fetch_address_list_entries",
                AddressBookListEntriesRequest = req,
            }, ct);
        }

        [HttpPost("request-address-book-group-members")]
        public async Task<IActionResult> RequestAddressBookGroupMembers([FromBody] AddressBookGroupMembersRequest? req, CancellationToken ct)
        {
            req ??= new AddressBookGroupMembersRequest();
            req.GroupId = req.GroupId?.Trim() ?? string.Empty;
            req.GroupSmtpAddress = req.GroupSmtpAddress?.Trim() ?? string.Empty;
            req.MaxMembers = Math.Clamp(req.MaxMembers <= 0 ? 5000 : req.MaxMembers, 1, 5000);
            if (string.IsNullOrWhiteSpace(req.GroupId) && string.IsNullOrWhiteSpace(req.GroupSmtpAddress))
                return BadRequest(ErrorEnvelope("fetch_address_book_group_members", "missing_required_fields", "Missing required request field(s): groupId or groupSmtpAddress."));

            var cached = _mailStore.GetAddressBookGroupMembers(req);
            if (!req.ForceRefresh && cached.State == "completed")
                return Ok(ResultEnvelope(cached.RequestId, "fetch_address_book_group_members", "completed", "Address book data is ready.", cached));
            if (cached.State == "loading")
                return Ok(ResultEnvelope(cached.RequestId, "fetch_address_book_group_members", "running", "Address book data is loading.", cached));

            var cmd = new PendingCommand
            {
                Type = "fetch_address_book_group_members",
                AddressBookGroupMembersRequest = req,
            };
            _mailStore.BeginAddressBookGroupExpansion(req, cmd.Id);
            return await DispatchCommandAsync(cmd, ct);
        }

        private Task<IActionResult> DispatchAddressBookListEntries(AddressBookSyncRequest req, CancellationToken ct)
        {
            return DispatchCommandAsync(new PendingCommand
            {
                Type = "fetch_address_list_entries",
                AddressBookListEntriesRequest = new AddressBookListEntriesRequest
                {
                    AddressListId = req.AddressListId,
                    AddressListName = req.AddressListName,
                    Offset = req.Offset,
                    PageSize = req.PageSize,
                },
            }, ct, "request-address-book");
        }

        private async Task<IActionResult> DispatchAddressBookGroupMembers(AddressBookSyncRequest req, CancellationToken ct)
        {
            var groupReq = new AddressBookGroupMembersRequest
            {
                GroupId = req.GroupId,
                GroupSmtpAddress = req.GroupSmtpAddress,
                MaxMembers = Math.Clamp(req.MaxGroupMembers <= 0 ? 5000 : req.MaxGroupMembers, 1, 5000),
                ForceRefresh = req.ForceRefresh,
            };
            var cached = _mailStore.GetAddressBookGroupMembers(groupReq);
            if (!groupReq.ForceRefresh && cached.State == "completed")
                return Ok(ResultEnvelope(cached.RequestId, "fetch_address_book_group_members", "completed", "Address book data is ready.", cached, "request-address-book"));
            if (cached.State == "loading")
                return Ok(ResultEnvelope(cached.RequestId, "fetch_address_book_group_members", "running", "Address book data is loading.", cached, "request-address-book"));

            var groupCmd = new PendingCommand
            {
                Type = "fetch_address_book_group_members",
                AddressBookGroupMembersRequest = groupReq,
            };
            _mailStore.BeginAddressBookGroupExpansion(groupReq, groupCmd.Id);
            return await DispatchCommandAsync(groupCmd, ct, "request-address-book");
        }
    }
}

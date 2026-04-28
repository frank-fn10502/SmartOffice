# AGENTS.md

Guidance for AI coding agents working in this repository.

## Language Policy

AI agents MUST communicate with the user in Traditional Chinese, while keeping technical proper nouns in English where that is clearer or conventional.

Examples:

- Use `SignalR`, `Swagger`, `Docker`, `devcontainer`, `MCP`, `DTO`, `Controller`, `Add-in`, `Office COM/VSTO` as English technical terms.
- Explain implementation notes in Traditional Chinese.
- If an English paragraph is necessary, provide a Traditional Chinese translation or summary immediately after it.
- Code, API names, file paths, commands, JSON fields, class names, and commit-style identifiers should remain in their original language.

## Project Mission

SmartOffice.Hub is a local mediation layer between Office 2016 add-ins, a Web UI, and AI/MCP clients. Preserve that boundary:

- Add-ins own Office automation and Office-specific integration.
- The Hub owns HTTP APIs, SignalR notifications, command routing, temporary state, and later AI/MCP integration.
- The Web UI owns user inspection, manual requests, chat, and diagnostics.

When making changes, prefer small, explicit contracts over hidden coupling. Office 2016 and restricted corporate environments are part of the design constraint, not an implementation detail.

## Repository Layout

- `Program.cs`: application startup and dependency registration.
- `Controllers/`: HTTP API boundaries. Keep Office-specific routes clearly named.
- `Hubs/`: SignalR hubs for live browser updates.
- `Models/`: DTOs shared across Hub, add-ins, Web UI, and possible MCP clients.
- `Services/`: in-memory stores, queues, and application services.
- `wwwroot/`: static Web UI. It intentionally avoids a frontend build chain.

## Current Technology Choices

- ASP.NET Core on .NET 8.
- SignalR for real-time dashboard updates.
- Swagger through Swashbuckle.
- Static HTML/CSS/JavaScript in `wwwroot`.
- In-memory stores for the current prototype stage.

Do not introduce a database, frontend build system, background job framework, or AI SDK unless the task clearly requires it.

## Coding Rules

- Keep DTO changes backward-compatible when possible. Add-ins may lag behind the Hub.
- Avoid renaming JSON fields casually; the Office add-in and Web UI depend on them.
- Prefer route additions over breaking existing routes.
- Use `DateTime` consistently with the existing code unless a task explicitly migrates time handling.
- Add comments for architectural intent, protocol boundaries, or security-sensitive decisions. Do not add comments that merely restate the code.
- Keep APIs narrow and predictable for future MCP exposure.
- The Web UI is intentionally static and dependency-light. Avoid npm, bundlers, and large client frameworks unless requested.

## Security Notes

Assume mail bodies, folder names, and chat messages may contain sensitive business data.

Current prototype behavior is permissive:

- CORS accepts any origin.
- Swagger is always enabled.
- There is no authentication.
- Cached data is process-local memory.

If touching networking, AI, MCP, or file export behavior, call out privacy and security implications in the change summary.

## Add-in Protocol

The Outlook add-in uses a polling protocol:

1. Web UI, AI, or MCP client enqueues a command through the Hub.
2. Outlook add-in calls `GET /api/outlook/poll`.
3. Hub returns one command or `{ "type": "none" }`.
4. Add-in performs Office automation locally.
5. Add-in pushes results back to the Hub.
6. Hub updates cache and broadcasts SignalR events.

Keep this pattern intact unless replacing it deliberately across all callers.

## Validation

Preferred validation mode is Quick Mode: keep editing on the host machine and compile in a temporary Docker container.

```bash
./scripts/build-in-container.sh
```

If the host machine has the .NET SDK installed, this is also acceptable:

```bash
dotnet build
```

For API behavior changes, also exercise Swagger or add/update `.http` examples.

For Web UI changes, start the app and inspect:

```text
http://localhost:2805/
```

## Documentation Expectations

When adding a new Office add-in surface:

- Document the route prefix and command types.
- Add DTO descriptions for request/response contracts.
- Note whether data is cached, streamed, or persisted.
- Include the SignalR events the UI or tools should listen for.

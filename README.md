# SmartOffice.Hub

SmartOffice.Hub is the local bridge between Office 2016 add-ins, a Web UI, and AI/MCP tooling. The project exists for locked-down Windows office environments where the desktop Office applications cannot safely or directly talk to cloud AI services.

The current implementation focuses on Outlook. The same Hub pattern is intended to be reused by future Word, Excel, PowerPoint, or other Office add-ins.

## Purpose

The Hub sits in the middle of three actors:

- Office add-ins: open a chat window, read Office context, push results, and poll for commands.
- Web UI: lets a user inspect Outlook data, request actions, chat, and monitor add-in status.
- AI/MCP clients: can call the Hub APIs to inspect Office context or ask an add-in to perform an operation.

This design keeps Office 2016 automation local and explicit. The add-in remains responsible for Office COM/VSTO interaction, while this service handles API boundaries, command routing, real-time UI notifications, and temporary state.

## Current Architecture

```text
Web UI / AI / MCP
       |
       | REST + SignalR
       v
SmartOffice.Hub
       |
       | long-poll commands + push results
       v
Office 2016 Add-in
```

Important pieces:

- `Program.cs`: ASP.NET Core startup, CORS, Swagger, static files, SignalR, and in-memory services.
- `Controllers/OutlookController.cs`: REST API used by Web UI, AI/MCP clients, and the Outlook add-in.
- `Hubs/NotificationHub.cs`: SignalR endpoint for real-time Web UI updates.
- `Services/Stores.cs`: in-memory mail/folder/chat/status stores and the command queue.
- `Models/Dtos.cs`: shared DTO contract between Hub, Web UI, and add-ins.
- `wwwroot/`: static dashboard for mail browsing, chat, and add-in diagnostics.

## Runtime Model

The Web UI and AI/MCP side request work by enqueueing commands:

- `POST /api/outlook/request-folders`
- `POST /api/outlook/request-mails`

The Outlook add-in long-polls for pending commands:

- `GET /api/outlook/poll`

The add-in pushes results back to the Hub:

- `POST /api/outlook/push-folders`
- `POST /api/outlook/push-mails`

The Web UI receives updates through SignalR:

- `/hub/notifications`
- Events include `FoldersUpdated`, `MailsUpdated`, `NewChatMessage`, `AddinStatus`, and `AddinLog`.

## Run Locally

Requirements:

- .NET 8 SDK

Start the Hub:

```bash
dotnet run
```

The development profile currently listens on:

```text
http://localhost:2805
```

Useful URLs:

- Dashboard: `http://localhost:2805/`
- Swagger: `http://localhost:2805/swagger`

## Development Modes

There are three supported ways to work on this project.

### Host Mode

Use the .NET SDK installed on your host machine:

```bash
dotnet run
dotnet build
```

This is simple when the host already has a compatible .NET 8 SDK.

### Quick Mode

Quick Mode keeps the editor and normal development environment on the host machine, but runs compilation inside a temporary Docker container.

```bash
./scripts/build-in-container.sh
```

This is the preferred build workflow when you do not want to install or maintain the .NET SDK directly on the host. The script builds a reusable local image from `.devcontainer/Dockerfile` when needed, then runs compilation in a temporary container. The build container is removed after the build finishes.

You can change the local image tag or build configuration:

```bash
SMARTOFFICE_BUILD_IMAGE=smartoffice-hub-dev:local CONFIGURATION=Release ./scripts/build-in-container.sh
```

### Full Container Mode

The optional `.devcontainer` folder lets VS Code reopen the entire workspace inside a .NET 8 development container.

Use this when you want the editor terminal, SDK, and C# tooling to all run in Docker. The devcontainer uses `.devcontainer/Dockerfile` so future native packages and tooling can be added in one place.

The devcontainer intentionally does not run `dotnet restore` automatically. Restore and run commands are manual so opening the container does not unexpectedly download packages.

See `.devcontainer/README.md`.

## API Notes

The Outlook route prefix is:

```text
/api/outlook
```

Main Web UI / AI request endpoints:

- `POST /request-folders`: enqueue a folder fetch command.
- `POST /request-mails`: enqueue a mail fetch command.
- `GET /folders`: read cached folders.
- `GET /mails`: read cached mails.
- `POST /chat`: append and broadcast a chat message.
- `GET /chat`: read cached chat messages.

Main add-in endpoints:

- `GET /poll`: long-poll for one pending command, with a 30 second timeout.
- `POST /push-folders`: replace cached folders and broadcast updates.
- `POST /push-mails`: replace cached mails and broadcast updates.

Admin endpoints:

- `GET /admin/status`
- `GET /admin/logs`
- `POST /admin/log`

## Security Assumptions

This project is currently shaped for a trusted local or intranet environment:

- CORS allows any origin with credentials.
- Swagger is enabled unconditionally.
- Data is stored in memory only.
- There is no authentication or authorization yet.

Before using this outside a controlled workstation or lab network, add authentication, restrict CORS, decide whether Swagger should be development-only, and review what mail content may be exposed to AI/MCP clients.

## Development Direction

Near-term work likely belongs in these areas:

- Add a provider-agnostic AI service layer.
- Add MCP-facing endpoints/tools around Office context and command dispatch.
- Split Office-specific APIs as more add-ins are added, for example `/api/word`, `/api/excel`, and `/api/powerpoint`.
- Add durable storage or a bounded cache if state must survive process restarts.
- Add command correlation and completion/error reporting so Web UI and AI clients can track a request end-to-end.
- Add tests around command queue behavior, DTO contracts, and controller responses.

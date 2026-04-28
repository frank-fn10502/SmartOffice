# SmartOffice Web UI

This folder contains the static dashboard served by SmartOffice.Hub.

The UI is intentionally lightweight because the target environment may be locked down and may not allow npm, bundlers, or frequent dependency changes.

## Files

```text
wwwroot/
├── index.html       # Dashboard markup and page-level JavaScript
├── styles.css       # Dashboard styling
└── folder-tree.js   # Outlook folder tree rendering logic
```

## Responsibilities

- Request Outlook folder and mail fetches from the Hub.
- Display cached Outlook folders and mail.
- Send and receive chat messages.
- Show Outlook add-in connection status and logs.
- Subscribe to SignalR events for real-time updates.

## External Runtime Dependency

The page loads SignalR from a CDN:

```html
https://cdnjs.cloudflare.com/ajax/libs/microsoft-signalr/8.0.0/signalr.min.js
```

For stricter offline or intranet-only deployments, vendor this file locally and update `index.html`.

## API Usage

The UI talks to the Hub through same-origin endpoints:

- `POST /api/outlook/request-folders`
- `POST /api/outlook/request-mails`
- `GET /api/outlook/folders`
- `GET /api/outlook/mails`
- `GET /api/outlook/chat`
- `POST /api/outlook/chat`
- `GET /api/outlook/admin/status`
- `GET /api/outlook/admin/logs`

SignalR endpoint:

- `/hub/notifications`

SignalR events consumed by the UI:

- `FoldersUpdated`
- `MailsUpdated`
- `NewChatMessage`
- `AddinStatus`
- `AddinLog`

## Development Notes

- Keep this UI dependency-light unless the project explicitly moves to a frontend build pipeline.
- Do not put secrets or AI provider keys in client-side files.
- Treat rendered mail content as sensitive. The current HTML mail view uses an iframe to isolate display, but it is not a full sanitization boundary.


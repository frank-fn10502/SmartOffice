# Backend Extension Guidelines

SmartOffice.Hub 的後端要能支援多個 Office VSTO AddIn。新增 Office domain 時，請維持「共用 Hub 基礎」與「單一 Office feature」的邊界。

## Feature Boundary

每個 Office AddIn backend feature 應有自己的 controller、service registration、store、mock 與 Swagger document：

- Controller route 使用 `api/<domain>`，例如 `api/outlook`、`api/excel`。
- SignalR route 使用 domain 命名，例如 `/hub/outlook-addin`。不要讓不同 AddIn 共用同一個 command event payload，除非 payload 已抽成明確的共用 contract。
- Swagger document 使用 `<domain>-v1`，每個 domain 的 operation filter 只描述該 domain API。
- Mock backend 與 store 放在 domain service 內，避免 Outlook mail/folder/calendar state 被其他 AddIn 依賴。

`Program.cs` 不應直接堆疊大量 domain service 註冊。每個 domain 應提供一個 feature registration extension，集中處理：

- service registration
- Swagger document registration
- SignalR hub mapping
- mock seed 或 mock backend startup

Outlook 目前的入口是 `Services/OutlookFeatureRegistration.cs`。

## Contract Shape

Web UI、AI 與 raw HTTP caller 面對的流程必須保持一致：

1. `POST /api/<domain>/request-*` 建立工作。
2. response 回傳 `requestId`、`request`、`state`、`message` 與 `data.fetchResultEndpoint`。
3. caller 用 paired `POST /api/<domain>/fetch-result-*` 輪詢狀態與分頁資料。

不要新增「只有 Web UI 知道怎麼讀」的 side-channel。SignalR 可以用於 worker status、logs、progress notification 或 diagnostics，但主要資料仍必須透過 request/fetch-result 讀取。

## Queue And Load Control

VSTO/COM automation 是高風險資源。新增 domain command queue 時必須明確處理：

- 單一 AddIn 的 command serialization 或節流。
- 大量資料的分頁、slice、batch 與 max count。
- worker unavailable、timeout、request failed 的可讀 response。
- caller 斷線後，Hub 已接受的 request 是否仍應完成。

不要在 startup 載入所有 Office domain 資料。startup 只做目前預設 workspace 必要資料；其他 domain/view 必須 lazy load。

## Naming

Outlook-specific 型別、class 與 route 可以維持 Outlook 命名。不要為了抽象而把已經穩定的 Outlook contract 改成泛型 Office contract。只有真的會被多個 AddIn 共用的基礎元件才放到共用層。

新增 Word、Excel、PowerPoint 或其他 AddIn 時，優先複用流程形狀，而不是複用 Outlook DTO。不同 Office object model 的資料語意應由各自 domain DTO 表達。

# Validation Notes for AI Agents

## Quick Mode

偏好的驗證模式是在本機編輯，使用暫存 Docker container 編譯：

```bash
./scripts/build-in-container.sh
```

如果 `webui/package.json` 存在，Quick Mode 會先執行前端 build，再執行 .NET build。
不要在 host 直接執行 npm 指令；前端 npm install/build 必須在 Docker/devcontainer 內執行。

## Host Mode

如果本機已安裝 .NET 8 SDK，也可以執行：

```bash
dotnet build
```

## API 檢查

如果修改 API behavior，也請透過 Swagger 或 `.http` example 檢查。

常用網址：

```text
http://localhost:2805/swagger
```

## Web UI 檢查

如果修改 Web UI，請用 Docker 啟動 Hub 後檢查：

```bash
./scripts/start-dev-container.sh
```

這個 script 預設載入 `appsettings.Mock.json`，讓 Web UI 檢查不依賴真實 Office Add-in。

```text
http://localhost:2805/
```

檢查完請停止 container：

```bash
./scripts/stop-dev-container.sh
```

只需要 build 驗證時，請執行：

```bash
./scripts/build-in-container.sh
```

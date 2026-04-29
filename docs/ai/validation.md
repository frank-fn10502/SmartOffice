# Validation Notes for AI Agents

## Quick Mode

偏好的驗證模式是在本機編輯，使用暫存 Docker container 編譯：

```bash
./scripts/build-in-container.sh
```

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

如果修改 Web UI，請啟動 app 後檢查：

```text
http://localhost:2805/
```

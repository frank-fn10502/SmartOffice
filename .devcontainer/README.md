# SmartOffice.Hub devcontainer

這個資料夾提供「可選」的 devcontainer 開發環境。它不取代本機快速開發流程，只是在你需要完整 container 開發環境時提供一致的 .NET 8 SDK、Node.js 22 與 VS Code 設定。

## 目前設計

- `devcontainer.json` 使用本資料夾內的 `Dockerfile`。
- `Dockerfile` 建立 .NET 8 devcontainer baseline，並安裝 Node.js 22，供 Vue/Vite build 使用。
- 沒有設定 `postCreateCommand`，所以開啟 container 後不會自動執行 `dotnet restore`、安裝 NuGet package 或安裝 npm package。
- `2805` 會自動 forward，對應 `Properties/launchSettings.json` 的本機開發 port。

## 完整容器模式 Full Container Mode

當你想讓 editor terminal、.NET SDK、C# tooling 都在 Docker 裡執行時，才使用這個模式。

VS Code 操作：

1. 安裝 Docker。
2. 安裝 Dev Containers extension。
3. 執行 `Dev Containers: Reopen in Container`。
4. 容器啟動後手動執行：

```bash
dotnet restore SmartOffice.Hub.sln
cd webui && npm install && cd ..
dotnet run --urls http://0.0.0.0:2805
```

開啟：

```text
http://localhost:2805/
```

## 快速模式 Quick Mode

日常偏好的工作方式仍然是 Quick Mode：在本機編輯，只用暫存 container 編譯。

```bash
./scripts/build-in-container.sh
```

這個腳本會在需要時先用 `.devcontainer/Dockerfile` 建立 reusable local image，然後用 `docker run --rm` 開一個暫存 container 執行 build。若 `webui/package.json` 存在，會先在 container 內執行 npm install 與 `npm run build`，再執行 `dotnet build`。

Quick Mode 會把 `webui/node_modules` 掛到 Docker volume，而不是寫入 repository 工作目錄。也就是說，image 與 npm packages 會留在 Docker 管理的空間以加速後續編譯，但每次實際編譯用的 container 不會常駐。

## 未來額外套件

之後如果需要安裝 native library、CLI tool、MCP 相關工具或企業內部憑證設定，請優先集中在 `.devcontainer/Dockerfile` 裡，避免散落在 README 或個人機器設定。

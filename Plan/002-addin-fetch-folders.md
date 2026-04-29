# Task 002：AddIn 實作 Fetch Folders

## 執行位置

本任務請在工作機的完整 SmartOffice / Outlook AddIn solution 中執行。

不要修改 `SmartOffice.Hub` 程式碼。

## 目標

AddIn 收到 `fetch_folders` 後，讀取 Outlook folder tree，並 push 到 Hub。

## Hub Contract

Poll 會收到：

```json
{
  "type": "fetch_folders"
}
```

完成後呼叫：

```http
POST /api/outlook/push-folders
```

## Payload Shape

```json
[
  {
    "name": "Inbox",
    "folderPath": "\\\\Mailbox - User\\Inbox",
    "itemCount": 10,
    "subFolders": []
  }
]
```

## 實作步驟

1. 在 command handler 加入 `fetch_folders`。
2. 使用 Outlook object model 列出 store / root folder / child folders。
3. 轉成 Hub 的 `FolderDto` JSON。
4. `folderPath` 必須可供後續 fetch mails 或 move mail 使用。
5. POST `/api/outlook/push-folders`。
6. 失敗時呼叫 `/api/outlook/admin/log`，log 不可包含敏感 folder 全名，除非已匿名化。

## 驗證

1. Web UI 按 Fetch Folders。
2. 確認資料夾樹顯示。
3. 確認 Hub log 沒有 error。

## 完成後更新

更新工作機 `Plan/WORKSTATION-STATUS.md`：

- `002-addin-fetch-folders` 改為 done。
- 下一個任務改為 `003-addin-fetch-mails.md`。

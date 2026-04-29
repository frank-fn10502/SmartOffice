# Task 008：AddIn 實作建立 Folder

## 執行位置

本任務請在工作機的完整 SmartOffice / Outlook AddIn solution 中執行。

不要修改 `SmartOffice.Hub` 程式碼。

## 目標

AddIn 支援 Hub 發出的 `create_folder` command。

## Request Shape

```json
{
  "parentFolderPath": "\\\\Mailbox - User\\Projects",
  "name": "Sample Folder"
}
```

## 實作步驟

1. 在 command handler 加入 `create_folder`。
2. 使用 `parentFolderPath` 定位 parent folder。
3. 檢查同名 child folder 是否已存在。
4. 不存在才建立。
5. 成功後 push folders，讓 Web UI 更新。

## 注意事項

- 不要刪除 folder。
- 不要移動 folder。
- folder name 可能含敏感資料，log 要匿名化。
- 名稱為空或包含 Outlook 不接受字元時要拒絕。

## 驗證

1. enqueue `create_folder`。
2. 確認 Outlook 出現資料夾。
3. Web UI Fetch Folders 後可看到新資料夾。

## 完成後更新

更新工作機 `Plan/WORKSTATION-STATUS.md`：

- `008-addin-create-folder` 改為 done。
- 下一個任務改為 `009-addin-move-mail.md`。

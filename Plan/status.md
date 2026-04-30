# Plan Status

本檔供 VS Code Copilot `Plan Runner` 使用，用來依序挑選 `Plan/` 任務並記錄狀態。

## 使用規則

- `Status` 可為 `todo`、`doing`、`done`、`blocked`。
- `Plan Runner` 每次只處理第一個 `todo` 或 `doing` 任務。
- `Plan Worker` 只處理被指派的單一任務，不自行挑選下一個任務。
- 這些任務多數應在工作機完整 SmartOffice / Outlook AddIn solution 中執行；除非任務明確要求 Hub contract 變更，否則不要修改 `SmartOffice.Hub` 程式碼。
- 測試與回報不得包含真實 mail body、folder name、rule name、calendar subject、attendee、客戶名稱或公司內部資訊。

## Tasks

### T000 - 工作機讀取 Hub Contract
Status: todo
File: `Plan/000-workstation-read-contract.md`
Run location: Workstation SmartOffice solution
Validation: 文件建立完成，且不含真實 business data
Notes:

### T001 - AddIn 確認 Hub Poll Command
Status: todo
File: `Plan/001-addin-poll-command.md`
Run location: Workstation SmartOffice solution
Validation: 啟動 Hub 與 Outlook AddIn，確認無 command 時可正常 poll
Notes:

### T002 - AddIn Fetch Folders
Status: todo
File: `Plan/002-addin-fetch-folders.md`
Run location: Workstation SmartOffice solution
Validation: 依任務檔案執行
Notes:

### T003 - AddIn Fetch Mails
Status: todo
File: `Plan/003-addin-fetch-mails.md`
Run location: Workstation SmartOffice solution
Validation: 依任務檔案執行
Notes:

### T004 - AddIn Mail Metadata
Status: todo
File: `Plan/004-addin-mail-metadata.md`
Run location: Workstation SmartOffice solution
Validation: 依任務檔案執行
Notes:

### T005 - AddIn Fetch Rules
Status: todo
File: `Plan/005-addin-fetch-rules.md`
Run location: Workstation SmartOffice solution
Validation: 依任務檔案執行
Notes:

### T006 - AddIn Fetch Calendar
Status: todo
File: `Plan/006-addin-fetch-calendar.md`
Run location: Workstation SmartOffice solution
Validation: 依任務檔案執行
Notes:

### T007 - AddIn Mark Mail
Status: todo
File: `Plan/007-addin-mark-mail.md`
Run location: Workstation SmartOffice solution
Validation: 依任務檔案執行
Notes:

### T008 - AddIn Create Folder
Status: todo
File: `Plan/008-addin-create-folder.md`
Run location: Workstation SmartOffice solution
Validation: 依任務檔案執行
Notes:

### T009 - AddIn Move Mail
Status: todo
File: `Plan/009-addin-move-mail.md`
Run location: Workstation SmartOffice solution
Validation: 依任務檔案執行
Notes:

### T010 - AddIn Command Result Log
Status: todo
File: `Plan/010-addin-command-result-log.md`
Run location: Workstation SmartOffice solution
Validation: 依任務檔案執行
Notes:

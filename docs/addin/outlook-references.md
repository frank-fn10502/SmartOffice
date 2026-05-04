# Office 2016 Add-in 線上文件

本文件只紀錄 Office 2016 Add-in 開發與 mock 校準時可查的線上文件入口。工作機與 Hub 的 JSON 格式請看 `docs/addin/signalr-contract.md`；工作機實測資料、差異與錯誤回報格式請看 `docs/addin/test-report.md`。

最後確認日期：2026-04-29。

## 使用原則

- 優先使用 Microsoft Learn 官方文件。
- 第三方文章只能作為輔助，不應作為 SmartOffice.Hub contract 依據。
- Office 2016 desktop 是主要目標環境；不要只看最新 API 文件就假設 Office 2016 可用。
- 如果工作機實測結果與文件描述不一致，或開發機需要真實資料校準 Web UI、mock、Add-in mapping、檔案寫入或 protocol，請用 `docs/addin/test-report.md` 的格式回報。

## VSTO / COM Add-in

Office 2016 desktop 深度整合通常會碰到 VSTO、COM automation 或 Outlook object model。這些文件最適合查詢 `Application`、`NameSpace`、`Folder`、`MailItem`、`Items` 等行為。

- [Office solutions development overview (VSTO)](https://learn.microsoft.com/en-us/visualstudio/vsto/office-solutions-development-overview-vsto?view=visualstudio)：VSTO Office solution 的總覽。
- [Outlook object model overview](https://learn.microsoft.com/en-us/visualstudio/vsto/outlook-object-model-overview?view=vs-2022)：Outlook VSTO 專案如何使用 Outlook object model。
- [Outlook VBA object model reference](https://learn.microsoft.com/en-us/office/vba/api/overview/outlook)：Outlook object model 的 VBA 參考；VSTO C# 常需要把 VBA sample 翻成 C# interop。
- [NameSpace object (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace)：MAPI root、default folders、store access。
- [Folder object (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.folder)：Outlook folder 與 nested folders。
- [Folders object (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.folders)：同一層 folder collection。
- [Folder.Folders property (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.folder.folders)：讀取子資料夾。
- [MailItem object (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem)：郵件 item、subject、sender、body、received time 等欄位。
- [Application.AdvancedSearch method (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.application.advancedsearch)：非同步搜尋；scope 可含同一個 store 內的多個 folder，不能跨 store。
- [Application.AdvancedSearchComplete event (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.application.advancedsearchcomplete)：`AdvancedSearch` 完成事件，避免以 blocking loop 等待。
- [Items.Restrict method (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.items.restrict)：在單一 folder items 內做條件篩選，適合搭配日期與 max count 縮小讀取範圍。

## Office JavaScript Add-in / Office.js

如果工作機 Add-in 是 Office.js 或混合架構，必須先查 Office 2016 支援的 requirement set。

- [Requirements for running Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/add-in-requirements)：Office Add-in 的 client / server / Outlook account 需求。
- [Office versions and requirement sets](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/office-versions-and-requirement-sets)：不同 Office 版本可用 API 的判斷方式。
- [Office Common API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/office-add-in-requirement-sets)：Common API requirement set 清單。
- [Outlook JavaScript API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)：Outlook `Mailbox` requirement set 與 manifest 宣告方式。
- [Outlook add-ins overview](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/read-scenario)：Outlook add-in activation、read / compose mode 與支援帳號。
- [Specify Office applications and API requirements with the add-in only manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)：使用 manifest 限制 host 與 API requirement。

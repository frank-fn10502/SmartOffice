# Plan Status

## 目前狀態

舊的工作機任務已清空，現在回到 Outlook AddIn 功能盤點階段。

## 目前文件

- `Plan/000-outlook-addin-required-features.md`

## 下一步

1. 使用者確認、增減或調整目前所需功能。
2. 工作機確認 AddIn 架構：VSTO / COM AddIn 或 Office.js。
3. 工作機依官方文件與 Outlook 2016 實測回報第一輪結果。
4. 依 `docs/ai/plan-splitting.md` 拆成可交給工作機 AI 執行的小任務。

## 注意事項

- 除非明確要求修改 Hub contract，後續 Plan 任務預設不要修改 SmartOffice.Hub 程式碼。
- 真實測試結果必須匿名化，不可提交敏感 business data。

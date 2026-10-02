# 會員手機更新：安裝與驗證

GitHub 的 Members.html 透過 config.js 的 CONFIG.API_URL 呼叫 Apps Script doPost，並非直接寫入試算表。本次需要先部署後端，再合併前端。

## 1. 更新 Apps Script

1. 打開提供 CONFIG.API_URL 的 Apps Script 專案。在「部署 → 管理部署作業」核對網頁應用程式 URL 與 config.js 完全相同。不要使用 Code.gs 註解中的其他通知網址。
2. 新增指令碼檔案 `MemberPhoneUpdate.gs`，貼入本資料夾同名檔案的完整內容。保留既有 Code.gs，不要用這個附加檔覆蓋整份 Code.gs。
3. 在原本 `doPost(e)` 的 `switch (action)` 中，Members 區段加入這一行（放在 default 之前）：

   ```javascript
   case 'updateMemberPhone': result = updateMemberPhone(data); break;
   ```

4. 儲存。到「部署 → 管理部署作業」，選擇目前 config.js 使用的部署，按編輯，版本選「新版本」，再部署。更新既有部署可以保留原本 /exec 網址；不需要修改 config.js。只按儲存不會更新既有版本部署。

附加檔沿用你提供的 COL、IDX、normalizePhone、formatPhoneNumber、invalidatePhoneCache_、updateMembershipSheet 等既有定義，必須放在相同 Apps Script 專案。

## 2. 更新 GitHub 網頁

後端部署完成後，合併此修改分支／Pull Request 到 main。等待既有 GitHub 網頁發布流程完成，再重新整理 Members.html。

## 3. 使用

輸入舊手機 →「查詢與配號」→ 在客戶資料下方輸入新手機 →「儲存新手機」→ 確認會員編號、姓名與新舊號碼。

- 手機須為 09 開頭 10 碼，可含空格或連字號。
- 保留同一筆 CRM 的會員編號、姓名、消費設定、來店次數與日期；僅改 B 欄電話。
- 若新號碼已存在於任何 CRM 客戶，拒絕更新，不自動合併資料。
- 舊號碼重複或會員編號已變動時，拒絕更新並要求重新確認。
- 成功後清除電話快取、同步 Membership，網頁重新查詢新號碼。
- RawData、iPad 訂單與歸檔中的歷史電話保持原始值；這次不進行歷史紀錄搬移。以電話關聯的歷史查詢不會自動將舊號碼紀錄合併到新號碼。已開啟的其他頁面如有預載資料，需要重新整理。
- 若提示 CRM 已更新但列表同步失敗，按列表的「重新整理（同步）」。若連線錯誤無法確認結果，先查新號碼，再查舊號碼，確認後才重試。

## 驗證

本地執行：`node tests/member-phone-update.test.cjs`。

測試使用模擬試算表與頁面元素，涵蓋資料保留、格式與重複檢查、舊資料防護、鎖定忙碌、同步部分失敗、查詢回應順序、重複送出與網路失敗；不會連線或修改真實會員。尚需在部署後以測試會員確認 Apps Script 權限與實際試算表同步。

新函式與既有 processCustomerVisit 共用 ScriptLock。既有未使用此鎖的管理函式（例如手動排序、會員刪除）與試算表手動編輯不受此鎖保護，請避免同時操作同一會員。網頁內的新增、刪除與更新操作已互相鎖定。

官方部署說明：https://developers.google.com/apps-script/concepts/deployments#edit_a_versioned_deployment

/**
 * 新增到原 Apps Script 專案；保留既有 Code.gs。
 * 在 doPost 的 switch(action) 新增：
 * case 'updateMemberPhone': result = updateMemberPhone(data); break;
 * 使用既有 COL、IDX、SPREADSHEET_ID、CRM_SHEET_NAME 與同步／快取函式。
 */
function updateMemberPhone(data) {
  data = data || {};
  const oldPhone = String(data.oldPhone || '').trim().replace(/[\s-]/g, '');
  const newPhone = String(data.newPhone || '').trim().replace(/[\s-]/g, '');
  const expectedMemberId = String(data.memberId || '').trim();
  if (!/^09\d{8}$/.test(oldPhone) || !/^09\d{8}$/.test(newPhone)) {
    return { error: '手機號碼須為 09 開頭的 10 碼數字（可含空格或連字號）。' };
  }
  if (!expectedMemberId) return { error: '請先查詢具有會員編號的客戶。' };
  if (oldPhone === newPhone) return { error: '新手機號碼與原號碼相同。' };

  // 與 processCustomerVisit 共用同一把鎖，避免來店寫入／排序同時執行。
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { error: '系統忙碌中，請稍後再試。' };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CRM_SHEET_NAME);
    if (!sheet) return { error: '找不到 CRM 工作表。' };
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { error: 'CRM 無資料。' };
    const rows = sheet.getRange(2, 1, lastRow - 1, COL.MEMBER_ID).getValues();
    const matches = [];
    let newPhoneExists = false;
    rows.forEach((row, index) => {
      const phone = normalizePhone(row[IDX('PHONE')]);
      if (phone === oldPhone) matches.push(index);
      if (phone === newPhone) newPhoneExists = true;
    });
    if (matches.length !== 1) {
      return { error: matches.length ? 'CRM 有多筆相同舊號碼，請先整理重複資料。' : '找不到舊手機號碼，請重新查詢會員。' };
    }
    if (newPhoneExists) return { error: '新手機號碼已登記於其他客戶，請先確認資料；不會自動合併。' };
    const index = matches[0];
    if (String(rows[index][IDX('MEMBER_ID')] || '').trim() !== expectedMemberId) {
      return { error: '會員編號已變動，請重新查詢後再更新。' };
    }

    // 僅更新原 CRM 列的電話；保留姓名、會員編號、來店次數等所有其他欄位。
    const phone = formatPhoneNumber(newPhone);
    invalidatePhoneCache_();
    sheet.getRange(index + 2, COL.PHONE).setValue(phone);
    SpreadsheetApp.flush();

    // CRM 已寫入後，衍生資料同步失敗不可誤報成「沒有更新」。
    const warnings = [];
    try {
      invalidatePhoneCache_();
    } catch (err) {
      warnings.push('查詢快取清除失敗，部分查詢可能暫時顯示舊資料。');
    }
    try {
      updateMembershipSheet();
    } catch (err) {
      warnings.push('會員列表同步失敗，請按「重新整理（同步）」重試。');
    }
    return {
      success: true,
      phone: phone,
      memberId: expectedMemberId,
      message: '手機號碼已更新，會員編號 ' + expectedMemberId + ' 與其他資料均保留。',
      warning: warnings.join('\n')
    };
  } finally {
    lock.releaseLock();
  }
}

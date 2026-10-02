/**
 * V2：手機更新與既有客戶合併。完整取代原 MemberPhoneUpdate.gs。
 * Code.gs 保留原路由，不需再新增：
 * case 'updateMemberPhone': result = updateMemberPhone(data); break;
 */
function updateMemberPhone(data) {
  data = data || {};
  const oldPhone = String(data.oldPhone || '').trim().replace(/[\s-]/g, '');
  const newPhone = String(data.newPhone || '').trim().replace(/[\s-]/g, '');
  const memberId = String(data.memberId || '').trim();
  if (!/^09\d{8}$/.test(oldPhone) || !/^09\d{8}$/.test(newPhone)) {
    return { error: '手機號碼須為 09 開頭的 10 碼數字（可含空格或連字號）。' };
  }
  if (!memberId) return { error: '請先查詢具有會員編號的客戶。' };
  if (oldPhone === newPhone) return { error: '新手機號碼與原號碼相同。' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { error: '系統忙碌中，請稍後再試。' };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CRM_SHEET_NAME);
    if (!sheet) return { error: '找不到 CRM 工作表。' };
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { error: 'CRM 無資料。' };
    const width = sheet.getLastColumn();
    const rows = sheet.getRange(2, 1, lastRow - 1, width).getValues();
    const oldMatches = [], newMatches = [];
    rows.forEach((row, index) => {
      const phone = normalizePhone(row[IDX('PHONE')]);
      if (phone === oldPhone) oldMatches.push(index);
      if (phone === newPhone) newMatches.push(index);
    });
    if (oldMatches.length !== 1) {
      return { error: oldMatches.length ? '舊號碼有多筆 CRM 資料，請先整理。' : '找不到舊號碼；若剛才曾送出更新，請查詢新號碼確認結果，不要重複合併。' };
    }
    if (newMatches.length > 1) return { error: '新號碼有多筆 CRM 資料，請先整理。' };
    const sourceIndex = oldMatches[0];
    const source = rows[sourceIndex];
    if (String(source[IDX('MEMBER_ID')] || '').trim() !== memberId) {
      return { error: '會員編號已變動，請重新查詢後再更新。' };
    }
    const phone = formatPhoneNumber(newPhone);
    let backupName = '';
    if (newMatches.length) {
      if (width < COL.LAST_VISIT_TIMESTAMP) return { error: 'CRM 欄位不完整，無法安全合併。' };
      const targetIndex = newMatches[0];
      const target = rows[targetIndex];
      const targetId = String(target[IDX('MEMBER_ID')] || '').trim();
      if (targetId && targetId !== memberId) return { error: '新號碼已有不同會員編號 ' + targetId + '，不可合併。' };
      const sourceRange = sheet.getRange(sourceIndex + 2, 1, 1, width);
      const targetRange = sheet.getRange(targetIndex + 2, 1, 1, width);
      if (sourceRange.getFormulas()[0].some(Boolean) || targetRange.getFormulas()[0].some(Boolean)) {
        return { error: '待合併的 CRM 列含有公式，請先確認，避免合併覆蓋公式。' };
      }
      // 原會員整列不變，只改電話；新號碼資料僅備份，不匯入原會員。
      const merged = source.slice();
      merged[IDX('PHONE')] = phone;
      // 確認期間若任一列變動（包括來店次數），拒絕使用舊預覽。
      const token = Utilities.base64EncodeWebSafe(Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        JSON.stringify([sourceIndex, targetIndex, source, target]), Utilities.Charset.UTF_8
      ));
      if (data.confirmMerge !== true) {
        return {
          requiresMerge: true, mergeToken: token,
          source: { name: String(source[IDX('NAME')] || ''), phone: formatPhoneNumber(oldPhone), visits: source[IDX('VISIT_COUNT')] },
          target: { name: String(target[IDX('NAME')] || ''), phone: phone, visits: target[IDX('VISIT_COUNT')] },
          memberId: memberId
        };
      }
      if (data.mergeToken !== token) return { error: '確認期間客戶資料已變動，請重新查詢並確認合併。' };
      // 備份成功後才開始合併；保留整張 CRM 的值、公式與格式。
      backupName = 'CRM_Backup_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss') + '_' + Utilities.getUuid().slice(0, 8);
      sheet.copyTo(ss).setName(backupName);
      SpreadsheetApp.flush();
      invalidatePhoneCache_();
      try {
        sourceRange.setValues([merged]);
        // 清空重複列而不 deleteRow，避免其他列位移。電話快取會跳過空白列。
        targetRange.clearContent();
        SpreadsheetApp.flush();
      } catch (err) {
        const recoveryErrors = [];
        try { sourceRange.setValues([source]); } catch (e) { recoveryErrors.push(e.message); }
        try { targetRange.setValues([target]); } catch (e) { recoveryErrors.push(e.message); }
        try { SpreadsheetApp.flush(); invalidatePhoneCache_(); } catch (e) { recoveryErrors.push(e.message); }
        return { error: recoveryErrors.length
          ? '合併失敗且自動還原未完整完成。請停止重試，依備份工作表 ' + backupName + ' 核對還原。'
          : '合併失敗，兩筆原資料已還原。備份工作表：' + backupName };
      }
    } else {
      // 已確認的合併請求不可在目標消失時悄悄改為普通換號。
      if (data.confirmMerge === true) return { error: '新號碼資料已變動，請重新查詢。' };
      invalidatePhoneCache_();
      sheet.getRange(sourceIndex + 2, COL.PHONE).setValue(phone);
      SpreadsheetApp.flush();
    }
    const warnings = [];
    try { invalidatePhoneCache_(); } catch (err) { warnings.push('查詢快取清除失敗，部分查詢可能暫時顯示舊資料。'); }
    try { updateMembershipSheet(); } catch (err) { warnings.push('會員列表同步失敗，請按「重新整理（同步）」重試。'); }
    return {
      success: true, phone: phone, memberId: memberId, backupSheet: backupName,
      message: backupName ? '手機已更新，會員編號 ' + memberId + ' 與所有原會員資料保持不變；新號碼的重複 CRM 資料已備份並移出有效資料。'
        : '手機號碼已更新，會員編號 ' + memberId + ' 與其他資料均保留。',
      warning: warnings.join('\n')
    };
  } finally { lock.releaseLock(); }
}


const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const root = path.join(__dirname, '..');
const backend = fs.readFileSync(path.join(root, 'apps-script/MemberPhoneUpdate.gs'), 'utf8');
const member = ['date', '0912-345-678', '測試會員', 'Card', 'Low', 'One', 'Two', '', '是Y', 'F33(4)', '', '備註', 12, 'last visit'];
function server(rows = [member.slice()], options = {}) {
  const calls = { writes: 0, cache: 0, sync: 0, release: 0 };
  const sheet = {
    getLastRow: () => rows.length + 1,
    getLastColumn: () => 14,
    copyTo: () => { if(options.backupFailure) throw Error('backup'); return { setName: () => {} }; },
    getRange: (r, c) => ({
      getValues: () => rows.map(row => row.slice()),
      getFormulas: () => [Array(14).fill(options.formulas ? '=1' : '')],
      setValues: values => { calls.writes++; rows[r - 2] = Array.from(values[0]); },
      clearContent: () => { if(options.clearFailure) throw Error('clear'); calls.writes++; rows[r - 2] = Array(14).fill(''); },
      setValue: value => { calls.writes++; rows[r - 2][c - 1] = value; }
    })
  };
  const ctx = vm.createContext({
    COL: { PHONE: 2, MEMBER_ID: 10, LAST_VISIT_TIMESTAMP: 14 }, IDX: name => ({ PHONE: 1, MEMBER_ID: 9, NAME: 2, VISIT_COUNT: 12 })[name],
    Utilities: { base64EncodeWebSafe: v => v, computeDigest: (algo, text) => text, DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' }, formatDate: () => 'test', getUuid: () => '12345678' },
    SPREADSHEET_ID: 'test', CRM_SHEET_NAME: 'CRM',
    normalizePhone: p => String(p || '').replace(/\D/g, ''),
    formatPhoneNumber: p => p.replace(/(\d{4})(\d{3})(\d{3})/, '$1-$2-$3'),
    SpreadsheetApp: { openById: () => ({ getSheetByName: () => sheet }), flush() {} },
    LockService: { getScriptLock: () => ({ tryLock: () => !options.busy, releaseLock: () => calls.release++ }) },
    invalidatePhoneCache_: () => calls.cache++,
    updateMembershipSheet: () => { calls.sync++; if (options.syncFailure) throw Error('sync failed'); }
  });
  vm.runInContext(backend, ctx);
  return { run: data => ctx.updateMemberPhone(data), rows, calls };
}
const payload = { oldPhone: '0912-345-678', newPhone: '0987654321', memberId: 'F33(4)' };
let s = server();
assert.equal(s.run(payload).success, true);
assert.deepEqual(s.rows[0], member.map((v, i) => i === 1 ? '0987-654-321' : v));
assert.deepEqual(s.calls, { writes: 1, cache: 2, sync: 1, release: 1 });
for (const data of [
  { ...payload, newPhone: '0912345678' }, { ...payload, newPhone: '123' },
  { ...payload, newPhone: 'abc0987654321' }, { ...payload, memberId: 'M1' },
  { ...payload, oldPhone: '0900000000' }, { ...payload, memberId: '' }
]) {
  s = server(); assert.ok(s.run(data).error); assert.equal(s.calls.writes, 0);
}
s = server([member.slice(), member.map((v, i) => i === 1 ? '0987-654-321' : v)]);
assert.equal(s.run(payload).requiresMerge, true); assert.equal(s.calls.writes, 0);
s = server([member.slice(), member.slice()]);
assert.ok(s.run(payload).error); assert.equal(s.calls.writes, 0);
s = server([], {}); assert.ok(s.run(payload).error);
s = server(undefined, { busy: true }); assert.ok(s.run(payload).error); assert.equal(s.calls.writes, 0);
s = server(undefined, { syncFailure: true });
const partial = s.run(payload); assert.equal(partial.success, true); assert.ok(partial.warning); assert.equal(s.calls.release, 1);

// Collision confirmation preserves every original field, including visits and dates.
const duplicate = member.map((v, i) => i === 1 ? '0987-654-321' : i === 9 ? '' : i === 12 ? 99 : 'different');
s = server([member.slice(), duplicate.slice()]);
let preview = s.run(payload);
assert.equal(preview.requiresMerge, true); assert.equal(s.calls.writes, 0);
assert.ok(s.run({ ...payload, confirmMerge: true, mergeToken: 'stale' }).error);
assert.equal(s.run({ ...payload, confirmMerge: true, mergeToken: preview.mergeToken }).success, true);
assert.deepEqual(s.rows[0], member.map((v,i) => i === 1 ? '0987-654-321' : v));
assert.ok(s.rows[1].every(v => v === ''));
assert.ok(s.run({ ...payload, confirmMerge: true, mergeToken: preview.mergeToken }).error);
s = server([member.slice(), duplicate.map((v,i) => i === 9 ? 'M8' : v)]);
assert.ok(s.run(payload).error); assert.equal(s.calls.writes, 0);
s = server([member.slice(), duplicate.slice()]); preview = s.run(payload); s.rows[1][12]++;
assert.ok(s.run({ ...payload, confirmMerge: true, mergeToken: preview.mergeToken }).error);
for (const option of ['backupFailure', 'clearFailure', 'formulas']) {
  s = server([member.slice(), duplicate.slice()], { [option]: true });
  preview = s.run(payload);
  if (option === 'formulas') assert.ok(preview.error);
  else if (option === 'backupFailure') assert.throws(() => s.run({ ...payload, confirmMerge: true, mergeToken: preview.mergeToken }));
  else assert.ok(s.run({ ...payload, confirmMerge: true, mergeToken: preview.mergeToken }).error);
  assert.deepEqual(s.rows, [member, duplicate]);
}

const html = fs.readFileSync(path.join(root, 'Members.html'), 'utf8');
const script = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]).join('\n');
const ids = [...html.matchAll(/id="([^"]+)"/g)].map(m => m[1]);
const elements = Object.fromEntries(ids.map(id => [id, {
  value: '', textContent: '', style: {}, disabled: false,
  classList: { add() {}, remove() {} }, focus() {}
}]));
const alerts = [];
const ctx = vm.createContext({
  CONFIG: { API_URL: 'test' }, window: {}, console, setTimeout: () => 1, clearTimeout() {},
  alert: msg => alerts.push(msg), confirm: () => true,
  document: { getElementById: id => { assert.ok(elements[id], `missing element: ${id}`); return elements[id]; } }
});
vm.runInContext(script, ctx);
vm.runInContext('loadTable = () => {};', ctx);
const result = phone => ({ hasId: true, existingId: 'F33(4)', customer: { phone, name: '測試會員', memberId: 'F33(4)' } });
(async () => {
  let pending = [];
  ctx.callApi = (action, data) => new Promise(resolve => pending.push({ action, data, resolve }));
  elements.phoneInput.value = '0912345678';
  const first = ctx.checkPhone();
  elements.phoneInput.value = '0900000000';
  const second = ctx.checkPhone();
  pending[1].resolve(result('0900000000')); await second;
  pending[0].resolve(result('0912345678')); await first;
  assert.equal(elements['show-phone'].textContent, '0900000000');
  ctx.invalidateMemberLookup();
  assert.equal(elements.phoneUpdateBox.style.display, 'none');
  assert.equal(elements.btnAdd.disabled, true);

  ctx.callApi = async () => result('0912-345-678');
  elements.phoneInput.value = '0912345678'; await ctx.checkPhone();
  assert.equal(elements.phoneUpdateBox.style.display, 'block');
  let updates = 0;
  let finish;
  ctx.callApi = (action, data) => {
    if (action === 'updateMemberPhone') {
      updates++; assert.deepEqual(JSON.parse(JSON.stringify(data)), payload);
      return new Promise(resolve => finish = resolve);
    }
    return Promise.resolve(result('0987-654-321'));
  };
  elements.newPhoneInput.value = 'bad'; await ctx.confirmUpdatePhone(); assert.equal(updates, 0);
  elements.newPhoneInput.value = '0912345678'; await ctx.confirmUpdatePhone(); assert.equal(updates, 0);
  elements.newPhoneInput.value = '0987-654-321';
  const update = ctx.confirmUpdatePhone();
  await ctx.confirmUpdatePhone(); await ctx.confirmDelete(); await ctx.checkPhone();
  assert.equal(updates, 1); assert.equal(elements.btnCheckPhone.disabled, true);
  finish({ success: true }); await update;
  assert.equal(elements.phoneInput.value, '0987654321');
  assert.equal(elements['show-phone'].textContent, '0987-654-321');
  assert.equal(elements.btnCheckPhone.disabled, false);

  elements.newPhoneInput.value = '0900000000';
  ctx.callApi = async () => { throw Error('network'); };
  await ctx.confirmUpdatePhone();
  assert.equal(elements.phoneUpdateBox.style.display, 'none');
  assert.equal(elements.btnUpdatePhone.disabled, false);
  assert.ok(alerts.some(msg => msg.includes('未能確認更新結果')));
  // Cancel merge preview: no second request and original selection remains usable.
  ctx.callApi = async () => result('0912-345-678');
  elements.phoneInput.value = '0912345678'; await ctx.checkPhone();
  elements.newPhoneInput.value = '0987654321';
  let confirms = 0, mergeCalls = 0;
  ctx.confirm = () => ++confirms === 1;
  const mergePreview = { requiresMerge: true, mergeToken: 'token', memberId: 'F33(4)', source: { name: 'Old', phone: '0912345678' }, target: { name: 'New', phone: '0987654321' } };
  ctx.callApi = async () => { mergeCalls++; return mergePreview; };
  await ctx.confirmUpdatePhone(); assert.equal(mergeCalls, 1);
  assert.equal(elements.phoneUpdateBox.style.display, 'block');
  assert.equal(elements.statusMsg.textContent, '已取消，資料未變更。');
  ctx.confirm = () => true;
  ctx.callApi = async (action, data) => {
    if (action !== 'updateMemberPhone') return result('0987-654-321');
    if (!data.confirmMerge) return mergePreview;
    assert.equal(data.mergeToken, 'token');
    return { success: true, backupSheet: 'backup' };
  };
  await ctx.confirmUpdatePhone();
  assert.equal(elements.phoneInput.value, '0987654321');
  console.log('PASS: backend validation/preservation/locking/partial sync and frontend stale lookup/double submit/failure recovery');
})().catch(err => { console.error(err); process.exitCode = 1; });

// 旅遊記帳後端：貼到 Google 試算表的「擴充功能 → Apps Script」，部署成網頁應用程式。
const EXP = 'expenses';
const MEM = 'members';

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(header); }
  return s;
}
function expSheet_() { return sheet_(EXP, ['id', 'time', 'item', 'amount', 'payer', 'shares']); }
function memSheet_() { return sheet_(MEM, ['name']); }

// 避免使用者輸入被試算表當成公式
function safe_(v) {
  v = String(v == null ? '' : v).slice(0, 100);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function readAll_() {
  const expenses = expSheet_().getDataRange().getValues().slice(1)
    .filter(r => r[0] !== '')
    .map(r => ({
      id: String(r[0]),
      item: String(r[2]),
      amount: Number(r[3]) || 0,
      payer: String(r[4]),
      shares: String(r[5]).split('|').filter(Boolean)
    }));
  const members = memSheet_().getDataRange().getValues().slice(1)
    .map(r => String(r[0])).filter(Boolean);
  return { members: members, expenses: expenses };
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function deleteWhere_(s, col, value) {
  const v = s.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][col]) === value) s.deleteRow(i + 1);
}

function doGet() { return out_(readAll_()); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'add') {
      const x = b.expense;
      const amount = Number(x.amount);
      if (!(amount > 0) || !x.payer || !x.shares || !x.shares.length) throw new Error('資料不完整');
      expSheet_().appendRow([safe_(x.id), new Date(), safe_(x.item), amount, safe_(x.payer),
        x.shares.map(safe_).join('|')]);
    } else if (b.action === 'delete') {
      deleteWhere_(expSheet_(), 0, String(b.id));
    } else if (b.action === 'addMember') {
      const name = safe_(String(b.name).trim().slice(0, 20));
      if (name && readAll_().members.indexOf(name) < 0) memSheet_().appendRow([name]);
    } else if (b.action === 'removeMember') {
      deleteWhere_(memSheet_(), 0, String(b.name));
    }
    return out_(readAll_());
  } catch (err) {
    return out_({ error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

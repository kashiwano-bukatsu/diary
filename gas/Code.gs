/**
 * DAY 用 GAS(スプレッドシートに紐づけて使う)
 * 使い方は SETUP.md を見てください。
 */

// ★ここを、自分だけの合言葉に書き換える(アプリの設定画面にも同じものを入れる)
const SECRET = 'ここを自分だけの合言葉に書き換える';

const TZ = 'Asia/Tokyo';
const START = '2026-11-17'; // DAY 0 の日

const NAMES = { rec: '記録', sales: '営業', num: '数字', sched: '時間割' };
const HEAD = {
  rec: ['日付', 'DAY', '今日の1つ', '達成', '最低ライン', '夜の記録', '明日の1つ', '記録した'],
  sales: ['日付', '声かけ', '話を聞いてもらえた', '提案', '成約', '面談数', 'つないだ人数', 'グループの人数'],
  num: ['日付', '事業', '売上', 'メモ'],
  sched: ['項目', '分']
};
const SALES_COL = { koe: 2, listen: 3, prop: 4, deal: 5, meet: 6, conn: 7, group: 8 };

/** 最初に1回だけ実行する(シートと見出しを作る) */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(NAMES).forEach(function (k) {
    let s = ss.getSheetByName(NAMES[k]);
    if (!s) s = ss.insertSheet(NAMES[k]);
    s.getRange(1, 1, 1, HEAD[k].length).setValues([HEAD[k]]).setFontWeight('bold');
    s.setFrozenRows(1);
    if (k !== 'sched') s.getRange(1, 1, 1000, 1).setNumberFormat('@'); // 日付は文字として持つ
  });
  const t = ss.getSheetByName(NAMES.sched);
  if (t.getLastRow() < 2) {
    t.getRange(2, 1, 4, 2).setValues([['営業・声かけ', 180], ['制作・撮影', 150], ['運営・事務', 150], ['休み・移動', 120]]);
  }
  const first = ss.getSheets().filter(function (s) { return Object.keys(NAMES).every(function (k) { return s.getName() !== NAMES[k]; }); });
  first.forEach(function (s) { if (s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s); });
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() { return out({ ok: true, name: 'DAY' }); }

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.key !== SECRET || SECRET === 'ここを自分だけの合言葉に書き換える') return out({ ok: false, error: 'auth' });
    if (req.action === 'init') return out({ ok: true, data: init_() });
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      switch (req.action) {
        case 'saveRecord': saveRecord_(req); break;
        case 'tally': tally_(req); break;
        case 'addNumber': addNumber_(req); break;
        default: return out({ ok: false, error: 'unknown action' });
      }
    } finally { lock.releaseLock(); }
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

/* ---------- ヘルパー ---------- */
function sh_(k) {
  const s = SpreadsheetApp.getActive().getSheetByName(NAMES[k]);
  if (!s) throw new Error('先に setup を実行してください');
  return s;
}
function ds_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v); }
function noon_(date) { return Date.parse(date + 'T12:00:00+09:00'); }
function addDays_(date, n) { return Utilities.formatDate(new Date(noon_(date) + n * 86400000), TZ, 'yyyy-MM-dd'); }
function dayNo_(date) { return Math.round((noon_(date) - noon_(START)) / 86400000); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function ensureRow_(k, date) {
  const s = sh_(k);
  const last = s.getLastRow();
  if (last >= 2) {
    const col = s.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < col.length; i++) if (ds_(col[i][0]) === date) return { s: s, r: i + 2 };
  }
  const r = last + 1;
  const w = HEAD[k].length;
  s.getRange(r, 1).setNumberFormat('@');
  const row = new Array(w).fill('');
  row[0] = date;
  s.getRange(r, 1, 1, w).setValues([row]);
  if (k === 'rec') s.getRange(r, 2).setValue(dayNo_(date));
  return { s: s, r: r };
}

/* ---------- 書き込み ---------- */
function saveRecord_(p) {
  const x = ensureRow_('rec', p.date);
  if (p.one !== undefined) x.s.getRange(x.r, 3).setValue(p.one);
  if (p.done !== undefined) x.s.getRange(x.r, 4).setValue(p.done ? true : '');
  if (p.min !== undefined) x.s.getRange(x.r, 5).setValue(p.min ? true : '');
  if (p.note !== undefined) {
    x.s.getRange(x.r, 6).setValue(p.note);
    x.s.getRange(x.r, 8).setValue(String(p.note).trim() ? true : '');
  }
  if (p.next) {
    const y = ensureRow_('rec', addDays_(p.date, 1));
    y.s.getRange(y.r, 3).setValue(p.next);
    y.s.getRange(y.r, 7).setValue('');
    x.s.getRange(x.r, 7).setValue(p.next);
  }
}

function tally_(p) {
  const c = SALES_COL[p.field];
  if (!c) throw new Error('bad field');
  const x = ensureRow_('sales', p.date);
  const cell = x.s.getRange(x.r, c);
  let cur = cell.getValue();
  if (p.field === 'group' && cur === '') cur = latestGroup_(p.date);
  cell.setValue(Math.max((Number(cur) || 0) + Number(p.delta || 0), 0));
}

function latestGroup_(beforeOrOn) {
  const s = sh_('sales');
  const last = s.getLastRow();
  if (last < 2) return 0;
  const v = s.getRange(2, 1, last - 1, 8).getValues();
  let best = '', val = 0;
  v.forEach(function (r) {
    const d = ds_(r[0]);
    if (r[7] !== '' && d <= beforeOrOn && d >= best) { best = d; val = Number(r[7]) || 0; }
  });
  return val;
}

function addNumber_(p) {
  const s = sh_('num');
  const r = s.getLastRow() + 1;
  s.getRange(r, 1).setNumberFormat('@');
  s.getRange(r, 1, 1, 4).setValues([[p.date, p.biz, Number(p.amount) || 0, p.memo || '']]);
}

/* ---------- 読み込み ---------- */
function init_() {
  const today = today_();
  const from60 = addDays_(today, -59);
  const from14 = addDays_(today, -13);
  const from62 = addDays_(today, -61);

  const records = {};
  let total = 0;
  const rs = sh_('rec');
  if (rs.getLastRow() >= 2) {
    rs.getRange(2, 1, rs.getLastRow() - 1, 8).getValues().forEach(function (r) {
      const d = ds_(r[0]);
      if (r[7] === true) total++;
      if (d >= from60) records[d] = { one: r[2], done: r[3] === true, min: r[4] === true, note: r[5], rec: r[7] === true };
    });
  }

  const sales = {};
  const ss = sh_('sales');
  if (ss.getLastRow() >= 2) {
    ss.getRange(2, 1, ss.getLastRow() - 1, 8).getValues().forEach(function (r) {
      const d = ds_(r[0]);
      if (d >= from14) sales[d] = { koe: Number(r[1]) || 0, listen: Number(r[2]) || 0, prop: Number(r[3]) || 0, deal: Number(r[4]) || 0, meet: Number(r[5]) || 0, conn: Number(r[6]) || 0 };
    });
  }

  const numbers = [];
  const ns = sh_('num');
  if (ns.getLastRow() >= 2) {
    ns.getRange(2, 1, ns.getLastRow() - 1, 4).getValues().forEach(function (r) {
      const d = ds_(r[0]);
      if (d >= from62) numbers.push({ date: d, biz: r[1], amount: Number(r[2]) || 0, memo: r[3] });
    });
  }

  const schedule = [];
  const ts = sh_('sched');
  if (ts.getLastRow() >= 2) {
    ts.getRange(2, 1, ts.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (r[0] !== '' && Number(r[1]) > 0) schedule.push({ name: String(r[0]), min: Number(r[1]) });
    });
  }

  return { today: today, total: total, records: records, sales: sales, group: latestGroup_(today), numbers: numbers, schedule: schedule };
}

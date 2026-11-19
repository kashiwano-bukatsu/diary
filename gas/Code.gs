/**
 * DAY 用 GAS(スプレッドシートに紐づけて使う)
 * 使い方は SETUP.md を見てください。
 */

// ★ここを、「名前|生年月日8桁」の形に書き換える(アプリのログイン画面で入れる名前と生年月日と同じにする)
// 例: 'なまえ|19900101'   ←この例のままにしないこと
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

  const cal = cal_(today);
  return { today: today, total: total, records: records, sales: sales, group: latestGroup_(today), numbers: numbers, schedule: schedule, calOk: cal.ok, calErr: cal.error || '', events: cal.events || [] };
}

/* ---------- Googleカレンダー(読み取りだけ) ---------- */
/** 最初に1回、GASの画面でこの関数を実行して「許可」を出す */
function authorizeCalendar() {
  const n = CalendarApp.getAllCalendars().length;
  Logger.log('カレンダー ' + n + ' 個にアクセスできます');
}

/** 今日の予定を、全部のカレンダーから集める(許可がまだなら calOk:false を返すだけで、他は動く) */
function cal_(today) {
  try {
    const start = new Date(today + 'T00:00:00+09:00');
    const end = new Date(start.getTime() + 86400000);
    const seen = {};
    const evs = [];
    CalendarApp.getAllCalendars().forEach(function (c) {
      try { if (c.isHidden()) return; } catch (e) {}
      c.getEvents(start, end).forEach(function (ev) {
        try { if (ev.getMyStatus() === CalendarApp.GuestStatus.NO) return; } catch (e) {}
        const s = ev.getStartTime().getTime(), e2 = ev.getEndTime().getTime();
        const key = ev.getTitle() + '|' + s + '|' + e2;
        if (seen[key]) return;
        seen[key] = true;
        evs.push({ title: ev.getTitle(), s: s, e: e2, allDay: ev.isAllDayEvent(), cal: c.getName(), col: c.getColor() });
      });
    });
    return { ok: true, events: evs };
  } catch (err) {
    return { ok: false, error: String(err), events: [] };
  }
}

/* ---------- カシワのBUKATSU「スケジュール」→ Googleカレンダー ---------- */
const KASHIWA_SHEET_ID = '1leak2KE05r1wWz3ozGhelhBqareKHqEuhElbvDarrSg'; // 管理シート(読み取りだけ)
const KASHIWA_CAL = 'カシワのBUKATSU';
const KASHIWA_DAYS = 90;     // 何日先まで写すか
const KASHIWA_HOURS = 2;     // 終了時間がシートに無いときの長さ(時間)

/** 最初に1回これを実行する(今すぐ1回写して、3時間ごとの自動実行も登録する) */
function startKashiwaSync() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncKashiwa') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncKashiwa').timeBased().everyHours(3).create();
  syncKashiwa();
}

function kashiwaCal_() {
  const found = CalendarApp.getCalendarsByName(KASHIWA_CAL);
  if (found.length) return found[0];
  const c = CalendarApp.createCalendar(KASHIWA_CAL, { timeZone: TZ });
  c.setColor(CalendarApp.Color.RED);
  return c;
}

function pad2_(n) { return ('0' + n).slice(-2); }

/** セルの日付 → 'yyyy-MM-dd'(Date でも「2026/11/20」のような文字でも読む) */
function kDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const m = String(v).normalize('NFKC').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]) : '';
}

/** セルの時間 → [開始 'HH:mm', 終了 'HH:mm' or ''](時刻でも「19:00」「19:00-21:00」でも読む) */
function kTimes_(v) {
  if (v instanceof Date) return [Utilities.formatDate(v, TZ, 'HH:mm'), ''];
  const m = String(v).normalize('NFKC').match(/\d{1,2}:\d{2}/g);
  if (!m) return ['', ''];
  const f = function (x) { const p = x.split(':'); return pad2_(p[0]) + ':' + p[1]; };
  return [f(m[0]), m[1] ? f(m[1]) : ''];
}

/** シートを読んで、カレンダーを合わせる(足りない予定は作る・変わった予定は直す・消えた予定は消す) */
function syncKashiwa() {
  const src = SpreadsheetApp.openById(KASHIWA_SHEET_ID).getSheetByName('スケジュール');
  if (!src) throw new Error('「スケジュール」シートが見つかりません');
  const vals = src.getDataRange().getValues();
  const head = vals[0].map(function (h) { return String(h).trim(); });
  const col = function (n) { return head.indexOf(n); };
  const cD = col('開催日'), cT = col('開始時間'), cP = col('場所'), cWho = col('予約者名'),
        cK = col('開催種目'), cC = col('コート'), cM = col('メモ'), cN = col('備考');
  if (cD < 0) throw new Error('「開催日」の列が見つかりません');

  const today = today_();
  const last = addDays_(today, KASHIWA_DAYS);
  const want = {};
  let skipped = 0;
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    const d = kDate_(r[cD]);
    if (!d) { if (String(r[cD]).trim()) skipped++; continue; }
    if (d < today || d > last) continue;
    const t = cT >= 0 ? kTimes_(r[cT]) : ['', ''];
    const kind = cK >= 0 ? String(r[cK]).trim() : '';
    const place = cP >= 0 ? String(r[cP]).trim() : '';
    const title = (kind || 'BUKATSU') + (place ? ' @' + place : '');
    const desc = [
      cC >= 0 && r[cC] !== '' ? 'コート: ' + r[cC] : '',
      cWho >= 0 && r[cWho] !== '' ? '予約者: ' + r[cWho] : '',
      cM >= 0 && r[cM] !== '' ? 'メモ: ' + r[cM] : '',
      cN >= 0 && r[cN] !== '' ? '備考: ' + r[cN] : '',
      '(管理シート「スケジュール」から自動で反映。直すときはシート側を直してください)'
    ].filter(String).join('\n');
    let s, e, allDay = false;
    if (t[0]) {
      s = new Date(d + 'T' + t[0] + ':00+09:00');
      e = t[1] ? new Date(d + 'T' + t[1] + ':00+09:00') : new Date(s.getTime() + KASHIWA_HOURS * 3600000);
      if (e <= s) e = new Date(s.getTime() + KASHIWA_HOURS * 3600000);
    } else {
      s = new Date(d + 'T00:00:00+09:00'); allDay = true;
    }
    want[title + '|' + s.getTime()] = { title: title, s: s, e: e, allDay: allDay, desc: desc, d: d };
  }

  const cal = kashiwaCal_();
  const from = new Date(today + 'T00:00:00+09:00');
  const to = new Date(noon_(last) + 86400000);
  let made = 0, fixed = 0, removed = 0;
  const have = {};
  cal.getEvents(from, to).forEach(function (ev) {
    const key = ev.getTitle() + '|' + ev.getStartTime().getTime();
    const w = want[key];
    if (!w || have[key]) { ev.deleteEvent(); removed++; return; }
    have[key] = true;
    if (ev.getDescription() !== w.desc) { ev.setDescription(w.desc); fixed++; }
    if (!w.allDay && ev.getEndTime().getTime() !== w.e.getTime()) { ev.setTime(w.s, w.e); fixed++; }
  });
  Object.keys(want).forEach(function (k) {
    if (have[k]) return;
    const w = want[k];
    if (w.allDay) cal.createAllDayEvent(w.title, w.s, { description: w.desc });
    else cal.createEvent(w.title, w.s, w.e, { description: w.desc });
    made++;
    Utilities.sleep(150);
  });
  Logger.log('カシワの予定: 追加 ' + made + ' / 直した ' + fixed + ' / 消した ' + removed + ' / 日付が読めない行 ' + skipped);
}

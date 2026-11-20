/**
 * Команда курʼєрів — сервер у Google Таблиці (таблиця створюється сама).
 * 1) Токен бота — у властивостях скрипта (BOT_TOKEN), ID власників — нижче.
 * 2) Розгорнути → Нове розгортання → Вебзастосунок → Виконувати від: Мене, Доступ: Усі.
 */
// Токен бота зберігається в «Налаштування проєкту → Властивості скрипта» під назвою BOT_TOKEN,
// тому код можна оновлювати копіюванням без жодних правок.
const BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN') || '';
const OWNER_IDS = ['7190012846', '5186267799']; // Женя, Гоша — бачать усю команду

const RATE = 30;          // % курʼєр віддає власнику
const RATE_HIGH = 25;     // % якщо за місяць заробив від порогу
const THRESHOLD = 20000;  // поріг, грн за місяць
const TZ = 'Europe/Kiev';
const HISTORY_DAYS = 125;

const SHEETS = {
  Couriers: ['tg_id', 'name', 'username', 'days', 'hours', 'joined', 'active', 'phone', 'source', 'ref', 'ref_name', 'ref_paid', 'note'],
  Entries: ['tg_id', 'name', 'date', 'hours', 'earned', 'orders', 'updated'],
  Payments: ['pid', 'tg_id', 'name', 'date', 'amount', 'note'],
};
const TEXT_COLS = {Couriers: ['A', 'D', 'F', 'H', 'J', 'L'], Entries: ['A', 'C'], Payments: ['A', 'B', 'D']};
const SOURCES = ['friend', 'olx', 'telegram', 'tiktok', 'other'];

function doGet() { return out({ok: true, msg: 'Сервер працює'}); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out({ok: false, error: 'bad_json'}); }
  const user = verify(req.initData);
  if (!user) return out({ok: false, error: 'auth'});
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    ensure();
    return out(handle(user, req || {}));
  } catch (err) {
    return out({ok: false, error: String((err && err.message) || err)});
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- перевірка, що запит справді з Telegram ---------- */
function verify(initData) {
  if (!initData || typeof initData !== 'string') return null;
  const p = {};
  initData.split('&').forEach(function (part) {
    const i = part.indexOf('=');
    if (i < 0) return;
    const dec = function (s) { return decodeURIComponent(s.replace(/\+/g, '%20')); };
    p[dec(part.slice(0, i))] = dec(part.slice(i + 1));
  });
  if (!p.hash) return null;
  const secret = Utilities.computeHmacSha256Signature(Utilities.newBlob(BOT_TOKEN).getBytes(), Utilities.newBlob('WebAppData').getBytes());
  const check = function (skip) {
    const s = Object.keys(p).filter(function (k) { return skip.indexOf(k) < 0; }).sort().map(function (k) { return k + '=' + p[k]; }).join('\n');
    return hex(Utilities.computeHmacSha256Signature(Utilities.newBlob(s).getBytes(), secret)) === p.hash;
  };
  if (!check(['hash']) && !check(['hash', 'signature'])) return null;
  if (Date.now() / 1000 - Number(p.auth_date || 0) > 86400 * 30) return null;
  let u;
  try { u = JSON.parse(p.user); } catch (e) { return null; }
  return u && u.id ? u : null;
}
function hex(bytes) { return bytes.map(function (b) { return ((b + 256) % 256).toString(16).padStart(2, '0'); }).join(''); }

/* ---------- таблиця ---------- */
let _ss = null;
// Працює і з таблиці, і як окремий проєкт: тоді таблиця створюється сама
function ss() {
  if (_ss) return _ss;
  _ss = SpreadsheetApp.getActive();
  if (_ss) return _ss;
  const props = PropertiesService.getScriptProperties(), id = props.getProperty('SHEET_ID');
  if (id) return (_ss = SpreadsheetApp.openById(id));
  _ss = SpreadsheetApp.create('Команда курʼєрів');
  props.setProperty('SHEET_ID', _ss.getId());
  return _ss;
}
function sheet(n) { return ss().getSheetByName(n); }
function ensure() {
  Object.keys(SHEETS).forEach(function (n) {
    let sh = sheet(n);
    if (!sh) {
      sh = ss().insertSheet(n);
      TEXT_COLS[n].forEach(function (c) { sh.getRange(c + ':' + c).setNumberFormat('@'); });
      sh.appendRow(SHEETS[n]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, SHEETS[n].length).setFontWeight('bold');
      return;
    }
    // нові колонки в старій таблиці
    const head = sh.getDataRange().getValues()[0];
    SHEETS[n].forEach(function (h, i) {
      if (head.indexOf(h) >= 0) return;
      const col = head.length + 1;
      const letter = String.fromCharCode(64 + col);
      if (TEXT_COLS[n].indexOf(String.fromCharCode(65 + i)) >= 0) sh.getRange(letter + ':' + letter).setNumberFormat('@');
      sh.getRange(1, col).setValue(h).setFontWeight('bold');
      head.push(h);
    });
  });
}
function setCol(n, row, name, value) {
  const head = sheet(n).getDataRange().getValues()[0], i = head.indexOf(name);
  if (i >= 0) sheet(n).getRange(row, i + 1).setValue(value);
}
function rows(n) {
  const v = sheet(n).getDataRange().getValues(), head = v.shift();
  return v.map(function (r, i) { const o = {_row: i + 2}; head.forEach(function (h, j) { o[h] = r[j]; }); return o; });
}
function dstr(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || ''); }
function today() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function shift(d, n) { const x = new Date(d + 'T12:00:00Z'); x.setUTCDate(x.getUTCDate() + n); return x.toISOString().slice(0, 10); }
function clean(s, max) { return String(s == null ? '' : s).replace(/[\u0000-\u001f<>]/g, '').trim().slice(0, max); }
function num(v, min, max) { const n = Number(String(v == null ? '' : v).replace(',', '.')); if (!isFinite(n)) return min; return Math.max(min, Math.min(max, Math.round(n * 10) / 10)); }
function days(v) {
  const a = Array.isArray(v) ? v : String(v || '').split(',');
  return a.map(Number).filter(function (d, i, s) { return d >= 1 && d <= 7 && s.indexOf(d) === i; }).sort();
}
function isTrue(v) { return v === true || String(v).toUpperCase() === 'TRUE' || v === 1; }
function courierOut(r) { return {id: String(r.tg_id), name: String(r.name), username: String(r.username || ''), days: days(r.days), hours: Number(r.hours) || 0, joined: dstr(r.joined), active: isTrue(r.active), phone: String(r.phone || ''), source: String(r.source || ''), ref: String(r.ref || ''), refName: String(r.ref_name || ''), refPaid: dstr(r.ref_paid), note: String(r.note || '')}; }
function cleanPhone(v) { return String(v || '').replace(/[^\d+]/g, '').slice(0, 16); }
function entryOut(r) { return {id: String(r.tg_id), date: dstr(r.date), hours: Number(r.hours) || 0, earned: Number(r.earned) || 0, orders: Number(r.orders) || 0}; }
function payOut(r) { return {pid: String(r.pid), id: String(r.tg_id), date: dstr(r.date), amount: Number(r.amount) || 0, note: String(r.note || '')}; }
function findCourier(id) { return rows('Couriers').filter(function (r) { return String(r.tg_id) === id; })[0]; }

/* ---------- дії ---------- */
function handle(u, req) {
  const me = String(u.id), owner = OWNER_IDS.indexOf(me) >= 0, a = req.action;
  const tgName = clean([u.first_name, u.last_name].filter(Boolean).join(' '), 40) || 'Курʼєр';

  if (a === 'register') {
    if (!findCourier(me)) {
      const src = SOURCES.indexOf(req.source) >= 0 ? req.source : '';
      let ref = /^\d{3,15}$/.test(String(req.ref || '')) && String(req.ref) !== me ? String(req.ref) : '';
      let refName = clean(req.refName, 40);
      if (ref) { const rc = findCourier(ref); if (rc) refName = String(rc.name); else ref = ''; }
      sheet('Couriers').appendRow([me, clean(req.name, 40) || tgName, clean(u.username, 40), days(req.days).join(','), num(req.hours, 1, 16), today(), true,
        cleanPhone(req.phone), ref || refName ? 'friend' : src, ref, refName, '', '']);
    }
  } else if (a === 'schedule') {
    const c = findCourier(me);
    if (!c) throw new Error('not_registered');
    sheet('Couriers').getRange(c._row, 2, 1, 4).setValues([[clean(req.name, 40) || c.name, clean(u.username, 40), days(req.days).join(','), num(req.hours, 1, 16)]]);
    if (req.phone !== undefined) setCol('Couriers', c._row, 'phone', cleanPhone(req.phone));
  } else if (a === 'entry' || a === 'delEntry') {
    const target = owner && req.for ? String(req.for) : me;
    const c = findCourier(target);
    if (!c) throw new Error('not_registered');
    if (!owner && !isTrue(c.active)) throw new Error('inactive');
    const d = String(req.date || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || d > today() || (!owner && d < shift(today(), -62))) throw new Error('bad_date');
    const ex = rows('Entries').filter(function (r) { return String(r.tg_id) === target && dstr(r.date) === d; })[0];
    if (a === 'delEntry') { if (ex) sheet('Entries').deleteRow(ex._row); }
    else {
      const row = [target, c.name, d, num(req.hours, 0, 24), num(req.earned, 0, 100000), num(req.orders, 0, 500), new Date().toISOString()];
      if (ex) sheet('Entries').getRange(ex._row, 1, 1, row.length).setValues([row]);
      else sheet('Entries').appendRow(row);
    }
  } else if (owner && a === 'pay') {
    const c = findCourier(String(req.for));
    const amount = num(req.amount, 0, 1000000);
    if (!c || !amount) throw new Error('bad_payment');
    const d = /^\d{4}-\d{2}-\d{2}$/.test(String(req.date)) ? String(req.date) : today();
    sheet('Payments').appendRow([Utilities.getUuid().slice(0, 8), String(c.tg_id), c.name, d, amount, clean(req.note, 80)]);
  } else if (owner && a === 'delPay') {
    const p = rows('Payments').filter(function (r) { return String(r.pid) === String(req.pid); })[0];
    if (p) sheet('Payments').deleteRow(p._row);
  } else if (owner && a === 'active') {
    const c = findCourier(String(req.for));
    if (c) sheet('Couriers').getRange(c._row, 7).setValue(!!req.active);
  } else if (owner && a === 'info') {
    const c = findCourier(String(req.for));
    if (!c) throw new Error('not_registered');
    if (req.note !== undefined) setCol('Couriers', c._row, 'note', clean(req.note, 300));
    if (req.phone !== undefined) setCol('Couriers', c._row, 'phone', cleanPhone(req.phone));
    if (req.source !== undefined) setCol('Couriers', c._row, 'source', SOURCES.indexOf(req.source) >= 0 ? req.source : '');
    if (req.ref !== undefined) {
      const rc = req.ref ? findCourier(String(req.ref)) : null;
      setCol('Couriers', c._row, 'ref', rc && String(rc.tg_id) !== String(c.tg_id) ? String(rc.tg_id) : '');
      setCol('Couriers', c._row, 'ref_name', rc ? String(rc.name) : clean(req.refName, 40));
      if (rc || req.refName) setCol('Couriers', c._row, 'source', 'friend');
    }
  } else if (owner && a === 'refPaid') {
    const c = findCourier(String(req.for));
    if (c) setCol('Couriers', c._row, 'ref_paid', req.paid ? today() : '');
  }
  return payload(u, owner);
}

function payload(u, owner) {
  const me = String(u.id), since = shift(today(), -HISTORY_DAYS);
  const cs = rows('Couriers').map(courierOut);
  const es = rows('Entries').map(entryOut).filter(function (e) { return e.date >= since; });
  const ps = rows('Payments').map(payOut).filter(function (p) { return p.date >= since; });
  const res = {ok: true, owner: owner, me: {id: me, name: u.first_name || ''}, today: today(), config: {rate: RATE, rateHigh: RATE_HIGH, threshold: THRESHOLD}};
  res.courier = cs.filter(function (c) { return c.id === me; })[0] || null;
  if (owner) {
    res.couriers = cs; res.entries = es; res.payments = ps; res.sheetUrl = ss().getUrl();
  } else {
    res.entries = es.filter(function (e) { return e.id === me; });
    res.payments = ps.filter(function (p) { return p.id === me; });
    if (res.courier) delete res.courier.note;
    // кого привів цей курʼєр: тільки імʼя, дата старту, кількість змін і статус бонусу
    res.myRefs = cs.filter(function (c) { return c.ref === me; }).map(function (c) {
      const mine = es.filter(function (e) { return e.id === c.id && e.earned > 0; }).map(function (e) { return e.date; }).sort();
      return {name: c.name, joined: c.joined, first: mine[0] || '', shifts: mine.length, active: c.active, paid: c.refPaid};
    });
  }
  return res;
}

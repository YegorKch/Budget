/**
 * API приложения «Бюджет».
 * Скрипт привязан к таблице (Расширения → Apps Script) и опубликован как веб-приложение.
 *
 * Все расчёты — на листе «Свод». Скрипт только читает готовые значения по ключам
 * из колонки C «Свода» и отбирает строки для списков.
 *
 * GET               → { ok, summary, top, categories, accounts, who, expected, recent, serverTime }
 * POST (text/plain) → { action: 'add' | 'update' | 'delete', ... } → то же, что GET, плюс result
 */

const SHEET_TX = 'Транзакции';
const SHEET_SUMMARY = 'Свод';
const SHEET_REF = 'Справочники';

const TX_COLS = ['ID', 'Дата', 'Статус', 'Тип', 'Сумма', 'Категория', 'Счёт', 'Кто', 'Комментарий', 'Создано'];
const COL = Object.fromEntries(TX_COLS.map((name, i) => [name, i + 1]));

const EXPECTED_DAYS = 14;
const RECENT_COUNT = 20;
const STATUSES = ['План', 'Факт'];
const TYPES = ['Расход', 'Доход'];

// ---------- HTTP ----------

function doGet() {
  return respond_(() => snapshot_());
}

function doPost(e) {
  return respond_(() => {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    let result;
    try {
      if (req.action === 'add') result = addTx_(req.tx);
      else if (req.action === 'update') result = updateTx_(req.id, req.fields);
      else if (req.action === 'delete') result = deleteTx_(req.id);
      else throw new Error('Неизвестное действие: ' + req.action);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    return Object.assign(snapshot_(), { result: result });
  });
}

function respond_(fn) {
  let body;
  try {
    body = Object.assign({ ok: true }, fn());
  } catch (err) {
    body = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- Чтение ----------

function snapshot_() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const sv = readSummary_(ss, tz);
  const ref = readRefs_(ss);
  const rows = readTx_(ss);
  return {
    summary: sv.values,
    top: sv.top,
    categories: ref.categories,
    accounts: ref.accounts,
    who: ref.who,
    expected: expected_(rows, sv, tz),
    recent: recent_(rows, tz),
    serverTime: new Date().toISOString(),
  };
}

/** Ключ → значение из колонки C/B «Свода» и таблица по категориям. */
function readSummary_(ss, tz) {
  const sh = ss.getSheetByName(SHEET_SUMMARY);
  const data = sh.getRange(1, 1, sh.getLastRow(), 8).getValues();
  const values = {};
  const topRaw = [];
  const byCategory = {};
  let tableStart = -1;

  data.forEach((row, i) => {
    const key = row[2];
    if (typeof key === 'string' && /^[a-z][A-Za-z0-9]*$/.test(key)) {
      if (/^top\d+$/.test(key)) topRaw.push({ category: row[0], count: row[1] });
      else values[key] = isDate_(row[1]) ? ymd_(row[1], tz) : row[1];
    }
    if (row[0] === 'Категория' && row[1] === 'Тип') tableStart = i + 1;
  });
  if (tableStart > 0) {
    for (let i = tableStart; i < data.length; i++) {
      const [cat, type, plan, fact] = data[i];
      if (!cat) continue;
      byCategory[cat] = { type: type, plan: num_(plan), fact: num_(fact) };
    }
  }
  return {
    values: values,
    top: topRaw.filter((t) => t.category).map((t) => ({ category: t.category, count: num_(t.count) })),
    byCategory: byCategory,
  };
}

function readRefs_(ss) {
  const sh = ss.getSheetByName(SHEET_REF);
  const last = Math.max(sh.getLastRow(), 2);
  const data = sh.getRange(2, 1, last - 1, 6).getValues();
  return {
    categories: data.filter((r) => r[0]).map((r) => ({ name: String(r[0]).trim(), type: r[1] || 'Расход' })),
    accounts: data.map((r) => r[3]).filter(Boolean).map(String),
    who: data.map((r) => r[5]).filter(Boolean).map(String),
  };
}

function readTx_(ss) {
  const sh = ss.getSheetByName(SHEET_TX);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, TX_COLS.length).getValues()
    .map((r, i) => ({ row: i + 2, v: r }))
    .filter((x) => x.v[0] !== '');
}

/**
 * Плановые записи, ещё не закрытые фактом: от начала периода до сегодня + 14 дней.
 * Факт категории (из «Свода») гасит плановые строки этой категории по порядку дат —
 * так конверт «Поесть 140 000» уменьшается по мере трат, а оплаченный интернет исчезает.
 */
function expected_(rows, sv, tz) {
  const today = sv.values.today || ymd_(new Date(), tz);
  const from = sv.values.periodStart || today;
  const to = addDays_(today, EXPECTED_DAYS);
  const periodEnd = sv.values.periodEnd || to;

  const plans = rows
    .map((x) => txObj_(x.v, tz))
    .filter((t) => t.status === 'План' && t.date >= from && t.date <= periodEnd)
    .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));

  const factLeft = {};
  Object.keys(sv.byCategory).forEach((c) => { factLeft[c] = sv.byCategory[c].fact; });

  const out = [];
  plans.forEach((p) => {
    const covered = Math.min(p.amount, factLeft[p.category] || 0);
    factLeft[p.category] = (factLeft[p.category] || 0) - covered;
    const left = p.amount - covered;
    if (left > 0 && p.date <= to) {
      out.push(Object.assign(p, { left: left, overdue: p.date < today }));
    }
  });
  return out;
}

function recent_(rows, tz) {
  const out = [];
  for (let i = rows.length - 1; i >= 0 && out.length < RECENT_COUNT; i--) {
    const t = txObj_(rows[i].v, tz);
    if (t.status === 'Факт') out.push(t);
  }
  return out;
}

function txObj_(v, tz) {
  return {
    id: String(v[COL['ID'] - 1]),
    date: isDate_(v[COL['Дата'] - 1]) ? ymd_(v[COL['Дата'] - 1], tz) : String(v[COL['Дата'] - 1]),
    status: v[COL['Статус'] - 1],
    type: v[COL['Тип'] - 1],
    amount: num_(v[COL['Сумма'] - 1]),
    category: v[COL['Категория'] - 1],
    account: v[COL['Счёт'] - 1],
    who: v[COL['Кто'] - 1],
    comment: v[COL['Комментарий'] - 1],
  };
}

// ---------- Запись ----------

/** tx: { id, date: 'YYYY-MM-DD', status, type, amount, category, account, who, comment } */
function addTx_(tx) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_TX);
  tx = validate_(tx, true);

  // Повторная отправка из оффлайн-очереди не должна задвоить запись.
  if (findRow_(sh, tx.id)) return { id: tx.id, duplicate: true };

  ensureCategory_(ss, tx.category, tx.type);
  sh.appendRow([
    tx.id, parseYmd_(tx.date), tx.status, tx.type, tx.amount, tx.category,
    tx.account || '', tx.who || '', tx.comment || '', new Date(),
  ]);
  return { id: tx.id };
}

/** fields: любые из { date, status, type, amount, category, account, who, comment } */
function updateTx_(id, fields) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_TX);
  const row = findRow_(sh, id);
  if (!row) throw new Error('Запись не найдена: ' + id);

  const range = sh.getRange(row, 1, 1, TX_COLS.length);
  const cur = txObj_(range.getValues()[0], ss.getSpreadsheetTimeZone());
  const tx = validate_(Object.assign(cur, fields || {}, { id: id }), false);

  ensureCategory_(ss, tx.category, tx.type);
  sh.getRange(row, COL['Дата'], 1, 8).setValues([[
    parseYmd_(tx.date), tx.status, tx.type, tx.amount, tx.category,
    tx.account || '', tx.who || '', tx.comment || '',
  ]]);
  return { id: id };
}

function deleteTx_(id) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_TX);
  const row = findRow_(sh, id);
  if (!row) return { id: id, missing: true };
  sh.deleteRow(row);
  return { id: id, deleted: true };
}

function validate_(tx, isNew) {
  if (!tx) throw new Error('Нет данных записи');
  const t = Object.assign({}, tx);
  t.id = String(t.id || '').trim();
  if (!t.id) {
    if (!isNew) throw new Error('Нет ID');
    t.id = Utilities.getUuid();
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(t.date))) throw new Error('Дата должна быть в формате ГГГГ-ММ-ДД');
  t.status = t.status || 'Факт';
  t.type = t.type || 'Расход';
  if (STATUSES.indexOf(t.status) < 0) throw new Error('Статус: ' + t.status);
  if (TYPES.indexOf(t.type) < 0) throw new Error('Тип: ' + t.type);
  t.amount = Math.round(Number(t.amount));
  if (!(t.amount > 0)) throw new Error('Сумма должна быть больше нуля');
  t.category = String(t.category || '').trim();
  if (!t.category) throw new Error('Не указана категория');
  t.comment = String(t.comment || '').trim();
  return t;
}

/** Новая категория из приложения дописывается в первую пустую строку справочника. */
function ensureCategory_(ss, name, type) {
  const sh = ss.getSheetByName(SHEET_REF);
  const last = Math.max(sh.getLastRow(), 1);
  const col = sh.getRange(1, 1, last + 1, 1).getValues();
  let empty = -1;
  for (let i = 1; i < col.length; i++) {
    const v = String(col[i][0]).trim();
    if (v.toLowerCase() === name.toLowerCase()) return;
    if (!v && empty < 0) empty = i + 1;
  }
  if (empty < 0) empty = col.length + 1;
  sh.getRange(empty, 1, 1, 2).setValues([[name, type]]);
}

function findRow_(sh, id) {
  if (!id) return 0;
  const cell = sh.getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).findNext();
  return cell && cell.getRow() > 1 ? cell.getRow() : 0;
}

// ---------- Утилиты ----------

function ymd_(d, tz) {
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

// instanceof Date ненадёжен для значений из getValues(), поэтому проверяем по тегу.
function isDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

/**
 * Полночь даты в часовом поясе таблицы. Пояс скрипта не используем: если он отличается
 * от пояса таблицы, дата в ячейке съезжает на день.
 */
function parseYmd_(s) {
  if (isDate_(s)) return new Date(s.getTime());
  return Utilities.parseDate(String(s), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

/** Арифметика над строкой ГГГГ-ММ-ДД, без часовых поясов. */
function addDays_(s, n) {
  const [y, m, d] = String(s).split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10);
}

function num_(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

// ---------- Проверка из редактора ----------

/** Запустить из редактора: выводит в журнал то, что получит приложение. */
function debugSnapshot() {
  Logger.log(JSON.stringify(snapshot_(), null, 2));
}

/** Добавляет, меняет и удаляет тестовую запись. В таблице после выполнения ничего не остаётся. */
function debugWriteCycle() {
  const id = 'test-' + Date.now();
  const today = ymd_(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone());
  Logger.log(addTx_({ id: id, date: today, amount: 1, category: 'Прочие', account: 'Наличные', who: 'Егор', comment: 'тест' }));
  Logger.log(addTx_({ id: id, date: today, amount: 1, category: 'Прочие' })); // должен вернуть duplicate
  Logger.log(updateTx_(id, { amount: 2, comment: 'тест изменён' }));
  Logger.log(deleteTx_(id));
}

/**
 * API приложения «Бюджет».
 * Скрипт привязан к таблице (Расширения → Apps Script) и опубликован как веб-приложение.
 *
 * Все расчёты — на листе «Свод». Скрипт только читает готовые значения по ключам
 * из колонки C «Свода» и отбирает строки для списков.
 *
 * GET               → { ok, summary, top, categories, accounts, balances, balanceHistory, who, expected, recent, feed, serverTime }
 * POST (text/plain) → { action: 'add' | 'update' | 'delete' | 'setBalance', ... } → то же, что GET, плюс result
 */

const SHEET_TX = 'Транзакции';
const SHEET_SUMMARY = 'Свод';
const SHEET_REF = 'Справочники';
const SHEET_ACCOUNTS = 'Счета';
const SHEET_HISTORY = 'История счетов';
const HISTORY_DAYS = 60;      // сколько хранить историю сверок счетов

const TX_COLS = ['ID', 'Дата', 'Статус', 'Тип', 'Сумма', 'Категория', 'Счёт', 'Кто', 'Комментарий', 'Создано', 'На счёт'];
const COL = Object.fromEntries(TX_COLS.map((name, i) => [name, i + 1]));

const EXPECTED_DAYS = 14;
const RECENT_COUNT = 20;
const STATUSES = ['План', 'Факт'];
const TYPES = ['Расход', 'Доход', 'Перевод'];
const TRANSFER = 'Перевод';   // тип и категория перевода между счетами: не трата и не доход

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
      else if (req.action === 'setBalance') result = setBalance_(req.account, req.value, req.who);
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
    balances: readBalances_(ss),
    balanceHistory: readBalanceHistory_(ss, tz),
    who: ref.who,
    expected: expected_(rows, sv, tz),
    recent: recent_(rows, tz),
    feed: feed_(rows, sv, tz),
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

/**
 * Остатки по счетам с листа «Счета» (считаются формулами).
 * v3: A счёт, B остаток при сверке, C когда сверен, H сейчас. v2: G сейчас. Нет листа — пусто.
 */
function readBalances_(ss) {
  const sh = ss.getSheetByName(SHEET_ACCOUNTS);
  if (!sh || sh.getLastRow() < 2) return [];
  const v3 = isAccountsV3_(sh);
  const nowCol = v3 ? 7 : 6;
  return sh.getRange(2, 1, Math.min(sh.getLastRow(), 31) - 1, 8).getValues()
    .filter((r) => r[0] !== '' && r[nowCol] !== '' && r[0] !== 'Без счёта')
    .map((r) => ({ name: String(r[0]), start: num_(r[1]), now: num_(r[nowCol]) }));
}

function isAccountsV3_(sh) {
  return sh.getRange(1, 8).getValue() === 'Сейчас';
}

/** История сверок за HISTORY_DAYS, новые сверху. */
function readBalanceHistory_(ss, tz) {
  const sh = ss.getSheetByName(SHEET_HISTORY);
  if (!sh || sh.getLastRow() < 2) return [];
  const since = Date.now() - HISTORY_DAYS * 864e5;
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues()
    .filter((r) => isDate_(r[0]) && r[0].getTime() >= since)
    .map((r) => ({
      at: Utilities.formatDate(r[0], tz, 'yyyy-MM-dd HH:mm'),
      account: String(r[1]), before: num_(r[2]), after: num_(r[3]), who: String(r[5] || ''),
    }))
    .reverse();
}

/**
 * Сверка счёта: «сейчас на счёте value». Запоминает сумму и момент сверки на листе «Счета»
 * (дальше остаток = эта сумма + операции, записанные после), пишет строку в историю
 * и удаляет из истории записи старше HISTORY_DAYS.
 */
function setBalance_(account, value, who) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_ACCOUNTS);
  if (!sh || !isAccountsV3_(sh)) throw new Error('Сначала запусти migrateAccountsV3 в редакторе Apps Script');
  account = String(account || '').trim();
  const v = Math.round(Number(value) * 100) / 100;
  if (!account) throw new Error('Не указан счёт');
  if (!isFinite(v)) throw new Error('Неверная сумма');

  const names = sh.getRange(2, 1, 30, 1).getValues().map((r) => String(r[0]));
  const i = names.indexOf(account);
  if (i < 0) throw new Error('Счёт не найден: ' + account);
  const row = i + 2;
  const before = num_(sh.getRange(row, 8).getValue());
  const now = new Date();
  sh.getRange(row, 2, 1, 2).setValues([[v, now]]);

  let hist = ss.getSheetByName(SHEET_HISTORY);
  if (!hist) hist = createHistorySheet_(ss);
  hist.appendRow([now, account, before, v, v - before, who || '']);
  pruneHistory_(hist);
  return { account: account, before: before, after: v };
}

function createHistorySheet_(ss) {
  const hist = ss.insertSheet(SHEET_HISTORY);
  hist.getRange(1, 1, 1, 6).setValues([['Когда', 'Счёт', 'Было', 'Стало', 'Разница', 'Кто']])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#37474F');
  hist.getRange('A:A').setNumberFormat('dd.mm.yyyy hh:mm');
  hist.getRange('C:E').setNumberFormat('#,##0;-#,##0;0');
  hist.setColumnWidth(1, 140);
  hist.setColumnWidth(2, 140);
  hist.setFrozenRows(1);
  return hist;
}

/** Строки истории идут по времени, поэтому старые — подряд сверху. */
function pruneHistory_(hist) {
  const last = hist.getLastRow();
  if (last < 2) return;
  const since = Date.now() - HISTORY_DAYS * 864e5;
  const dates = hist.getRange(2, 1, last - 1, 1).getValues();
  let old = 0;
  while (old < dates.length && isDate_(dates[old][0]) && dates[old][0].getTime() < since) old++;
  if (old > 0) hist.deleteRows(2, old);
}

function readTx_(ss) {
  const sh = ss.getSheetByName(SHEET_TX);
  const last = sh.getLastRow();
  if (last < 2) return [];
  // До миграции колонки «На счёт» может не быть — читаем столько, сколько есть.
  const width = Math.min(TX_COLS.length, sh.getMaxColumns());
  return sh.getRange(2, 1, last - 1, width).getValues()
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

/**
 * Лента по дням: все планы и факты от начала периода (и как минимум за последнюю неделю)
 * до сегодня + 14 дней. Приложение раскладывает их по дням само.
 */
function feed_(rows, sv, tz) {
  const today = sv.values.today || ymd_(new Date(), tz);
  const weekAgo = addDays_(today, -7);
  const start = sv.values.periodStart && sv.values.periodStart < weekAgo ? sv.values.periodStart : weekAgo;
  const to = addDays_(today, EXPECTED_DAYS);
  const pick = (t) => ({
    id: t.id, date: t.date, type: t.type, amount: t.amount, category: t.category,
    account: t.account, toAccount: t.toAccount, who: t.who, comment: t.status === 'Факт' ? t.comment : '',
  });
  const plans = [];
  const facts = [];
  rows.forEach((x) => {
    const t = txObj_(x.v, tz);
    if (t.date < start || t.date > to) return;
    if (t.status === 'План') plans.push(pick(t));
    else if (t.status === 'Факт') facts.push(pick(t));
  });
  return { from: start, to: to, today: today, plans: plans, facts: facts };
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
    toAccount: v[COL['На счёт'] - 1] || '',
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

  if (tx.type !== TRANSFER) ensureCategory_(ss, tx.category, tx.type);
  sh.appendRow([
    tx.id, parseYmd_(tx.date), tx.status, tx.type, tx.amount, tx.category,
    tx.account || '', tx.who || '', tx.comment || '', new Date(), tx.toAccount || '',
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

  if (tx.type !== TRANSFER) ensureCategory_(ss, tx.category, tx.type);
  sh.getRange(row, COL['Дата'], 1, 8).setValues([[
    parseYmd_(tx.date), tx.status, tx.type, tx.amount, tx.category,
    tx.account || '', tx.who || '', tx.comment || '',
  ]]);
  sh.getRange(row, COL['На счёт']).setValue(tx.type === TRANSFER ? tx.toAccount : '');
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
  if (t.type === TRANSFER) {
    t.status = 'Факт';
    t.category = TRANSFER;
    t.account = String(t.account || '').trim();
    t.toAccount = String(t.toAccount || '').trim();
    if (!t.account || !t.toAccount) throw new Error('Для перевода нужны оба счёта');
    if (t.account === t.toAccount) throw new Error('Счета перевода совпадают');
  } else {
    t.toAccount = '';
  }
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

// ---------- Миграция ----------

/**
 * v1 → v2: остатки по счетам и переводы. Запустить один раз из редактора; повторный запуск безопасен.
 * 1. «Транзакции»: колонка K «На счёт», тип «Перевод» в выпадающем списке.
 * 2. «Справочники»: колонка E «Остаток на начало» рядом со счетами. Если пусто — весь остаток Егора
 *    ставится на «Kaspi Егор», Арины — на «Kaspi Арина»; разнести по счетам потом вручную.
 * 3. «Свод»: остатки на начало Егора и Арины считаются из «Справочников».
 * 4. Новый лист «Счета» с формулами остатка по каждому счёту.
 */
function migrateAccountsV2() {
  const ss = SpreadsheetApp.getActive();
  const tx = ss.getSheetByName(SHEET_TX);
  const ref = ss.getSheetByName(SHEET_REF);
  const sv = ss.getSheetByName(SHEET_SUMMARY);
  const log = [];

  // 1. Транзакции
  if (tx.getMaxColumns() < COL['На счёт']) tx.insertColumnsAfter(tx.getMaxColumns(), COL['На счёт'] - tx.getMaxColumns());
  const hdr = tx.getRange(1, COL['На счёт']);
  if (hdr.getValue() !== 'На счёт') {
    hdr.setValue('На счёт');
    tx.getRange(1, COL['Создано']).copyFormatToRange(tx, COL['На счёт'], COL['На счёт'], 1, 1);
    log.push('Транзакции: добавлена колонка «На счёт»');
  }
  const accList = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ref.getRange('$D$2:$D$100'), true).setAllowInvalid(true).build();
  const txRows = tx.getMaxRows() - 1;
  tx.getRange(2, COL['На счёт'], txRows, 1).setDataValidation(accList);
  tx.getRange(2, COL['Тип'], txRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(TYPES, true).setAllowInvalid(false).build());

  // 2. Справочники: остаток на начало по счетам (после v3 не нужен — там сверки)
  const svData = sv.getRange(1, 1, sv.getLastRow(), 3).getValues();
  const keyRow = (key) => svData.findIndex((r) => r[2] === key) + 1;
  const rEgor = keyRow('startBalance1');
  const rArina = keyRow('startBalance2');
  const accSheet = ss.getSheetByName(SHEET_ACCOUNTS);
  const alreadyV3 = accSheet && isAccountsV3_(accSheet);
  if (!alreadyV3 && ref.getRange('E1').getValue() !== 'Остаток на начало') {
    ref.getRange('E1').setValue('Остаток на начало');
    ref.getRange('D1').copyFormatToRange(ref, 5, 5, 1, 1);
    const n = Math.min(99, ref.getMaxRows() - 1);
    const accs = ref.getRange(2, 4, n, 1).getValues();
    const egor = rEgor ? num_(sv.getRange(rEgor, 2).getValue()) : 0;
    const arina = rArina ? num_(sv.getRange(rArina, 2).getValue()) : 0;
    const vals = accs.map((r) => [r[0] === 'Kaspi Егор' ? egor : r[0] === 'Kaspi Арина' ? arina : (r[0] ? 0 : '')]);
    ref.getRange(2, 5, n, 1).setValues(vals);
    ref.getRange(2, 5, n, 1).setNumberFormat('#,##0.##').setBackground('#E3F2FD');
    ref.setColumnWidth(5, 140);
    log.push('Справочники: остатки на начало — Kaspi Егор ' + egor + ', Kaspi Арина ' + arina + ', остальные 0');
  }

  // 3. Свод: Егор и Арина — из Справочников
  const accRng = "'Справочники'!$D$2:$D$100";
  const balRng = "'Справочники'!$E$2:$E$100";
  if (rEgor && !sv.getRange(rEgor, 2).getFormula()) {
    sv.getRange(rEgor, 1).setValue('Остаток на начало: все счета, кроме Kaspi Арина');
    sv.getRange(rEgor, 2).setFormula(`=SUM(${balRng})-SUMIF(${accRng},"Kaspi Арина",${balRng})`).setBackground(null);
    log.push('Свод: остаток Егора теперь считается из Справочников');
  }
  if (rArina && !sv.getRange(rArina, 2).getFormula()) {
    sv.getRange(rArina, 1).setValue('Остаток на начало: Kaspi Арина');
    sv.getRange(rArina, 2).setFormula(`=SUMIF(${accRng},"Kaspi Арина",${balRng})`).setBackground(null);
    log.push('Свод: остаток Арины теперь считается из Справочников');
  }

  // 4. Лист «Счета»
  if (!ss.getSheetByName(SHEET_ACCOUNTS)) {
    const sh = ss.insertSheet(SHEET_ACCOUNTS, 1);
    const ps = `'Свод'!$B$${keyRow('periodStart')}`;
    const pe = `'Свод'!$B$${keyRow('periodEnd')}`;
    const T = "'Транзакции'";
    const period = `${T}!$B:$B,">="&${ps},${T}!$B:$B,"<="&${pe},${T}!$C:$C,"Факт"`;
    const head = ['Счёт', 'Остаток на начало', 'Доходы', 'Расходы', 'Переводы сюда', 'Переводы отсюда', 'Сейчас'];
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#37474F');
    const rows = [];
    for (let i = 2; i <= 31; i++) {
      const a = `$A${i}`;
      const blank = `IF(${a}="","",`;
      rows.push([
        `=IF('Справочники'!D${i}="","",'Справочники'!D${i})`,
        `=${blank}N('Справочники'!E${i}))`,
        `=${blank}SUMIFS(${T}!$E:$E,${T}!$G:$G,${a},${T}!$D:$D,"Доход",${period}))`,
        `=${blank}SUMIFS(${T}!$E:$E,${T}!$G:$G,${a},${T}!$D:$D,"Расход",${period}))`,
        `=${blank}SUMIFS(${T}!$E:$E,${T}!$K:$K,${a},${T}!$D:$D,"${TRANSFER}",${period}))`,
        `=${blank}SUMIFS(${T}!$E:$E,${T}!$G:$G,${a},${T}!$D:$D,"${TRANSFER}",${period}))`,
        `=${blank}B${i}+C${i}-D${i}+E${i}-F${i})`,
      ]);
    }
    sh.getRange(2, 1, rows.length, 7).setFormulas(rows).setNumberFormat('#,##0;-#,##0;""');
    // Факты без счёта или со счётом не из справочника — чтобы сумма по счетам сходилась с «На счетах».
    sh.getRange(33, 1, 1, 7).setValues([['Без счёта', '', '', '', '', '',
      `=ROUND('Свод'!$B$${keyRow('balanceNow')}-SUM(G2:G31),0)`]]);
    sh.getRange(33, 7).setNumberFormat('#,##0;-#,##0;""');
    sh.getRange(35, 1).setValue('Остатки на начало периода вводятся на листе «Справочники», колонка E.')
      .setFontStyle('italic').setFontColor('#777777');
    sh.setColumnWidth(1, 160);
    sh.setFrozenRows(1);
    log.push('Создан лист «Счета»');
  }

  Logger.log(log.length ? log.join('; ') : 'Всё уже было сделано, изменений нет');
}

/**
 * v2 → v3: сверки счетов вместо остатка на начало периода. Запустить один раз; повторный запуск безопасен.
 * 1. Лист «Счета» перестраивается: A счёт, B остаток при сверке, C когда сверен, D–G операции после сверки,
 *    H сейчас. Первая сверка — остатки из «Справочников» (колонка E) на начало периода.
 * 2. Колонка E «Справочников» очищается — дальше остатки правятся сверкой из приложения.
 * 3. «Свод»: «На счетах» = сумма счетов; свободные деньги = на счетах + ещё ожидаемые доходы −
 *    ещё не оплаченные плановые расходы (та же величина, что раньше, но от реального остатка);
 *    остаток на начало периода больше не вводится, а считается (нужен «Календарю»).
 * 4. Создаётся лист «История счетов».
 */
function migrateAccountsV3() {
  migrateAccountsV2();
  withEnglishLocale_(() => {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_ACCOUNTS);
    const ref = ss.getSheetByName(SHEET_REF);
    const sv = ss.getSheetByName(SHEET_SUMMARY);
    const log = [];

    // 1. Лист «Счета»
    if (!isAccountsV3_(sh)) {
      const ps = sv.getRange(summaryKeyRow_(sv, 'periodStart'), 2).getValue();
      const anchor = isDate_(ps) ? ps : new Date();
      // Начальные сверки: остатки на начало периода из «Справочников» (колонка E)
      const startBy = {};
      ref.getRange(2, 4, 30, 2).getValues().forEach((r) => { if (r[0]) startBy[String(r[0])] = num_(r[1]); });

      sh.clear();
      const head = ['Счёт', 'Остаток при сверке', 'Когда сверен', 'Доходы после', 'Расходы после',
        'Переводы сюда', 'Переводы отсюда', 'Сейчас'];
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#37474F');
      writeAccountFormulas_(sh);
      SpreadsheetApp.flush();
      const anchors = sh.getRange(2, 1, 30, 1).getValues()
        .map((r) => (r[0] ? [startBy[String(r[0])] || 0, anchor] : ['', '']));
      sh.getRange(2, 2, 30, 2).setValues(anchors);
      sh.getRange(2, 2, 30, 1).setNumberFormat('#,##0.##').setBackground('#E3F2FD');
      sh.getRange(2, 3, 30, 1).setNumberFormat('dd.mm.yyyy hh:mm');
      sh.getRange(2, 4, 30, 5).setNumberFormat('#,##0;-#,##0;""');
      sh.getRange(33, 1).setValue('Остаток правится сверкой из приложения: тап по счёту в «На счетах». ' +
        'Можно и руками: вписать сумму в B и текущие дату-время в C.').setFontStyle('italic').setFontColor('#777777');
      sh.setColumnWidth(1, 160);
      sh.setColumnWidth(3, 140);
      sh.setFrozenRows(1);
      log.push('Лист «Счета» переведён на сверки (первая сверка — на начало периода)');

      // 2. Справочники: колонка E больше не нужна
      if (ref.getRange('E1').getValue() === 'Остаток на начало') {
        ref.getRange(1, 5, ref.getMaxRows(), 1).clear();
        log.push('Справочники: колонка E очищена');
      }
    }

    // 3. Свод
    const rEgor = summaryKeyRow_(sv, 'startBalance1');
    const rArina = summaryKeyRow_(sv, 'startBalance2');
    if (rEgor || rArina) {
      // Удаляем снизу вверх, чтобы номера строк не съехали; ссылки в формулах таблица поправит сама.
      [rEgor, rArina].filter(Boolean).sort((x, y) => y - x).forEach((r) => sv.deleteRow(r));
      log.push('Свод: убраны ручные остатки на начало Егора и Арины');
    }
    writeSummaryFormulas_(sv);
    log.push('Свод: свободные деньги и «На счетах» считаются от остатков по счетам');

    // 4. История
    if (!ss.getSheetByName(SHEET_HISTORY)) {
      createHistorySheet_(ss);
      log.push('Создан лист «История счетов»');
    }

    Logger.log(log.join('; '));
  });
}

/**
 * Перезаписывает формулы, которые вписывает скрипт (лист «Счета» и расчёт на «Своде»).
 * Данные — суммы и моменты сверок, транзакции — не трогает. Запускать, если там #ERROR!.
 */
function fixFormulas() {
  withEnglishLocale_(() => {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_ACCOUNTS);
    if (!sh || !isAccountsV3_(sh)) throw new Error('Сначала запусти migrateAccountsV3');
    writeAccountFormulas_(sh);
    writeSummaryFormulas_(ss.getSheetByName(SHEET_SUMMARY));
    SpreadsheetApp.flush();
    const bad = sh.getRange(2, 1, 30, 8).getDisplayValues().flat().filter((v) => /^#/.test(v)).length;
    Logger.log(bad ? `Формулы перезаписаны, но ошибок всё ещё ${bad} — пришли текст ошибки из ячейки «Счета»!A2`
      : 'Формулы перезаписаны, ошибок нет');
  });
}

/**
 * Формулы пишутся в английской нотации (запятые между аргументами). В русской локали таблица
 * ждёт «;» и не разбирает такие формулы (#ERROR!) — на время записи переключаем локаль и возвращаем.
 */
function withEnglishLocale_(fn) {
  const ss = SpreadsheetApp.getActive();
  const locale = ss.getSpreadsheetLocale();
  if (locale !== 'en_US') ss.setSpreadsheetLocale('en_US');
  try {
    return fn();
  } finally {
    SpreadsheetApp.flush();
    if (locale !== 'en_US') ss.setSpreadsheetLocale(locale);
  }
}

function summaryKeyRow_(sv, key) {
  return sv.getRange(1, 3, sv.getLastRow(), 1).getValues().findIndex((r) => r[0] === key) + 1;
}

/** «Счета»: A — счёт из справочника, D–G — операции после сверки, H — сейчас. B и C (сверка) не трогает. */
function writeAccountFormulas_(sh) {
  const T = "'Транзакции'";
  const colA = [];
  const colDH = [];
  for (let i = 2; i <= 31; i++) {
    const a = `$A${i}`;
    const c = `N($C${i})`;
    // Операция учитывается, если записана после сверки (по «Создано»);
    // у строк, внесённых в таблицу руками без «Создано», — если её дата позже дня сверки.
    const after = (accCol, type) =>
      `SUMIFS(${T}!$E:$E,${T}!${accCol}:${accCol},${a},${T}!$D:$D,"${type}",${T}!$C:$C,"Факт",${T}!$J:$J,">"&${c})` +
      `+SUMIFS(${T}!$E:$E,${T}!${accCol}:${accCol},${a},${T}!$D:$D,"${type}",${T}!$C:$C,"Факт",${T}!$J:$J,"",${T}!$B:$B,">"&INT(${c}))`;
    const blank = `IF(${a}="","",`;
    colA.push([`=IF('Справочники'!D${i}="","",'Справочники'!D${i})`]);
    colDH.push([
      `=${blank}${after('$G', 'Доход')})`,
      `=${blank}${after('$G', 'Расход')})`,
      `=${blank}${after('$K', TRANSFER)})`,
      `=${blank}${after('$G', TRANSFER)})`,
      `=${blank}N(B${i})+D${i}-E${i}+F${i}-G${i})`,
    ]);
  }
  sh.getRange(2, 1, 30, 1).setFormulas(colA);
  sh.getRange(2, 4, 30, 5).setFormulas(colDH);
}

/** «Свод»: на счетах, свободные деньги и вычисляемый остаток на начало периода. */
function writeSummaryFormulas_(sv) {
  const keyRow = (key) => summaryKeyRow_(sv, key);
  const data = sv.getRange(1, 1, sv.getLastRow(), 2).getValues();
  const head = data.findIndex((r) => r[0] === 'Категория' && r[1] === 'Тип') + 1;
  if (head < 1) throw new Error('На «Своде» не найдена таблица по категориям');
  const first = head + 1;
  const last = first + 99;
  const col = (c) => `$${c}$${first}:$${c}$${last}`;
  const cell = (key) => `$B$${keyRow(key)}`;
  const setF = (key, label, f, note) => {
    const r = keyRow(key);
    if (!r) return;
    if (label) sv.getRange(r, 1).setValue(label);
    sv.getRange(r, 2).setFormula(f).setBackground(null);
    if (note) sv.getRange(r, 4).setValue(note);
  };
  setF('balanceNow', 'На счетах сейчас (сумма по листу «Счета»)', `=SUM('Счета'!$H$2:$H$31)`);
  setF('freeNow', null,
    `=${cell('balanceNow')}+SUMIF(${col('B')},"Доход",${col('H')})-SUMIF(${col('B')},"Расход",${col('H')})`,
    'На счетах + ещё ожидаемые доходы − ещё не оплаченные плановые расходы');
  setF('freeMorning', null,
    `=${cell('freeNow')}+SUMIF(${col('B')},"Расход",${col('F')})-SUMIF(${col('B')},"Расход",${col('G')})`,
    'То же, но сегодняшние траты не учтены');
  setF('startBalance', 'Остаток на начало периода (считается из счетов)',
    `=${cell('balanceNow')}-SUMIF(${col('B')},"Доход",${col('D')})+SUMIF(${col('B')},"Расход",${col('D')})`,
    'Нужен «Календарю»: на счетах сейчас минус операции периода');
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

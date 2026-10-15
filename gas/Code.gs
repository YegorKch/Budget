/**
 * API приложения «Бюджет».
 * Скрипт привязан к таблице (Расширения → Apps Script) и опубликован как веб-приложение.
 *
 * Все расчёты — на листе «Свод». Скрипт только читает готовые значения по ключам
 * из колонки C «Свода» и отбирает строки для списков.
 *
 * GET               → { ok, summary, top, categories, categoryTable, accounts, balances, balanceHistory, who,
 *                       expected, recent, feed, serverTime }
 * POST (text/plain) → { action: 'add' | 'update' | 'delete' | 'setBalance', ... } → то же, что GET, плюс result
 */

const SHEET_TX = 'Транзакции';
const SHEET_SUMMARY = 'Свод';
const SHEET_REF = 'Справочники';
const SHEET_ACCOUNTS = 'Счета';
const SHEET_HISTORY = 'История счетов';
const SHEET_PLAN_CAL = 'Календарь план';
const SHEET_ARCHIVE = 'Архив';
// Период бюджета начинается в день зарплаты. Зарплата («Зп») приходит с 5-го по 10-е: период — с её прихода
// и до следующей зарплаты, но не позже 10-го следующего месяца. Нет зарплаты к 10-му — новый период с 11-го.
const SALARY_CATEGORY = 'Зп';
const SALARY_DAY_FROM = 5;
const SALARY_DAY_TO = 10;
// «Календарь план»: строка с датами, первая строка категорий, колонка «Надёжность» (F), первая колонка дней (G)
const PC_HEAD = 4;
const PC_FIRST = 7;
const PC_REL_COL = 6;
const PC_DAY_COL = 7;
const HISTORY_DAYS = 60;      // сколько хранить историю сверок счетов

const TX_COLS = ['ID', 'Дата', 'Статус', 'Тип', 'Сумма', 'Категория', 'Счёт', 'Кто', 'Комментарий', 'Создано', 'На счёт'];
const COL = Object.fromEntries(TX_COLS.map((name, i) => [name, i + 1]));

const EXPECTED_DAYS = 14;
const RECENT_COUNT = 20;
const STATUSES = ['План', 'Факт'];
const TYPES = ['Расход', 'Доход', 'Перевод'];
const TRANSFER = 'Перевод';   // тип и категория перевода между счетами: не трата и не доход
// Обязательные статьи (пометка «да» в «Справочниках», колонка C) — для прогноза «хватит ли до зарплаты»
const MANDATORY_DEFAULT = ['Рассрочки', 'Арина рассрочки', 'Каспи рэд (кредитка)', 'Арина учеба', 'Ком услуги',
  'Домашний интернет', 'Сбережения', 'Арина англ', 'Тренер Арине', 'Подписки/Тарифы'];
// Ежедневные статьи (пометка «да» в «Справочниках», колонка E) — по их темпу прогноз считает, хватит ли денег
const DAILY_DEFAULT = ['Поесть', 'Такси', 'Общественный транспорт', 'Продукты'];
// Надёжность доходов («Справочники», колонка G): гарантированные входят в прогноз, ожидаемые — отдельной строкой
const GUARANTEED = 'Гарантировано';
const EXPECTED = 'Ожидается';
const GUARANTEED_DEFAULT = ['Зп', 'Стипендия', 'Пособие'];
// Окно зарплаты в конце периода (с 5-го по 10-е): гарантированные доходы и обязательные платежи, что приходятся
// на эти дни, — уже следующий период (следующая зарплата, рассрочки с неё). Прогноз текущего их не учитывает.
const NEXT_INCOME_DAYS = SALARY_DAY_TO - SALARY_DAY_FROM + 1;
// Пометка ежедневных статей в «Календаре план»: они сверяются с планом день за днём, а не конвертом
const DAILY_MARK = 'ежедневная';

// ---------- HTTP ----------

function doGet() {
  return respond_(() => {
    ensurePeriodLocked_();
    return snapshot_();
  });
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
      ensurePeriod_(SpreadsheetApp.getActive());   // записали зарплату — новый период начинается сразу
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
    categoryTable: sv.table,
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

/**
 * Ключ → значение из колонки C/B «Свода» и таблица по категориям
 * (A категория, B тип, C план, D факт, H осталось по плану, M обязательная, N ежедневная, O надёжность дохода).
 */
function readSummary_(ss, tz) {
  const sh = ss.getSheetByName(SHEET_SUMMARY);
  const data = sh.getRange(1, 1, sh.getLastRow(), Math.min(15, sh.getMaxColumns())).getValues();
  const values = {};
  const topRaw = [];
  const byCategory = {};
  const table = [];
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
      // Одно имя может быть и доходом, и расходом («Прочие») — ключ с типом.
      byCategory[cat + '|' + type] = { type: type, plan: num_(plan), fact: num_(fact) };
      table.push({ name: String(cat), type: type, plan: num_(plan), fact: num_(fact),
        left: num_(data[i][7]), mandatory: data[i][12] === 'да', daily: data[i][13] === 'да',
        guaranteed: data[i][14] === GUARANTEED });
    }
  }
  return {
    values: values,
    top: topRaw.filter((t) => t.category).map((t) => ({ category: t.category, count: num_(t.count) })),
    byCategory: byCategory,
    table: table,
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
  Object.keys(sv.byCategory).forEach((k) => { factLeft[k] = sv.byCategory[k].fact; });

  const out = [];
  plans.forEach((p) => {
    const key = p.category + '|' + p.type;
    const covered = Math.min(p.amount, factLeft[key] || 0);
    factLeft[key] = (factLeft[key] || 0) - covered;
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
  if (/^[=+\-@]/.test(t.comment)) t.comment = "'" + t.comment;
  return t;
}

/**
 * Новая категория из приложения дописывается в первую пустую строку справочника.
 * Категория — это пара «имя + тип»: «Прочие» может быть и доходом, и расходом.
 * Имя сравнивается без учёта регистра — так же, как его сравнивают формулы таблицы.
 */
function ensureCategory_(ss, name, type) {
  const sh = ss.getSheetByName(SHEET_REF);
  const last = Math.max(sh.getLastRow(), 1);
  const rows = sh.getRange(1, 1, last + 1, 2).getValues();
  let empty = -1;
  for (let i = 1; i < rows.length; i++) {
    const v = String(rows[i][0]).trim();
    if (v.toLowerCase() === name.toLowerCase() && (rows[i][1] || 'Расход') === type) return;
    if (!v && empty < 0) empty = i + 1;
  }
  if (empty < 0) empty = rows.length + 1;
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
    const sv = ss.getSheetByName(SHEET_SUMMARY);
    writeSummaryFormulas_(sv);
    if (summaryKeyRow_(sv, 'lifeMoney')) writeForecastFormulas_(sv);
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
  // Исходную локаль запоминаем: если прошлый запуск оборвался по времени, таблица осталась в en_US,
  // и вернуть надо ту, что была до него.
  const props = PropertiesService.getScriptProperties();
  let locale = ss.getSpreadsheetLocale();
  if (locale === 'en_US' && props.getProperty('origLocale')) locale = props.getProperty('origLocale');
  else props.setProperty('origLocale', locale);
  if (ss.getSpreadsheetLocale() !== 'en_US') ss.setSpreadsheetLocale('en_US');
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

/**
 * v3 → v4: прогноз до зарплаты и исправления по итогам анализа. Запустить один раз; повторно — безопасно.
 * 1. «Справочники»: колонка C «Обязательная» («да» у рассрочек, учёбы, коммуналки и т.п.).
 * 2. «Справочники»: дописываются пары «категория + тип», которые есть в транзакциях, но нет в справочнике
 *    (например «Прочие» как доход) — иначе такие суммы попадали не в ту строку «Свода».
 * 3. «Транзакции»: комментарии, превратившиеся в формулу (#NAME?), возвращаются текстом.
 * 4. «Свод»: таблица по категориям считает с учётом типа (доход/расход); новые колонки
 *    L «Факт за 7 дней» и M «Обязательная»; блок прогноза до зарплаты.
 */
function migrateV4() {
  withEnglishLocale_(() => {
    const ss = SpreadsheetApp.getActive();
    const ref = ss.getSheetByName(SHEET_REF);
    const tx = ss.getSheetByName(SHEET_TX);
    const sv = ss.getSheetByName(SHEET_SUMMARY);
    const log = [];

    // 1. Обязательные статьи
    if (ref.getRange('C1').getValue() !== 'Обязательная') {
      ref.getRange('C1').setValue('Обязательная');
      ref.getRange('A1').copyFormatToRange(ref, 3, 3, 1, 1);
      const n = Math.max(ref.getLastRow() - 1, 1);
      const names = ref.getRange(2, 1, n, 2).getValues();
      ref.getRange(2, 3, n, 1).setValues(names.map((r) =>
        [r[1] === 'Расход' && MANDATORY_DEFAULT.indexOf(String(r[0])) >= 0 ? 'да' : '']));
      ref.getRange(2, 3, Math.min(ref.getMaxRows() - 1, 499), 1).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(['да'], true).setAllowInvalid(true).build());
      ref.setColumnWidth(3, 110);
      log.push('Справочники: колонка «Обязательная»');
    }

    // 1б. Ежедневные статьи
    if (ref.getRange('E1').getValue() !== 'Ежедневная') {
      ref.getRange('E1').setValue('Ежедневная');
      ref.getRange('A1').copyFormatToRange(ref, 5, 5, 1, 1);
      const n = Math.max(ref.getLastRow() - 1, 1);
      const names = ref.getRange(2, 1, n, 2).getValues();
      ref.getRange(2, 5, n, 1).setValues(names.map((r) =>
        [r[1] === 'Расход' && DAILY_DEFAULT.indexOf(String(r[0])) >= 0 ? 'да' : '']));
      log.push('Справочники: колонка «Ежедневная»');
    }

    // 2. Недостающие пары «категория + тип» из транзакций — справочник читаем один раз, дописываем пачкой
    const last = tx.getLastRow();
    const txv = last > 1 ? tx.getRange(2, 1, last - 1, COL['Комментарий']).getValues() : [];
    const refN = Math.max(ref.getLastRow() - 1, 1);
    const known = new Set(ref.getRange(2, 1, refN, 2).getValues()
      .filter((r) => r[0]).map((r) => String(r[0]).trim().toLowerCase() + '|' + (r[1] || 'Расход')));
    const added = [];
    txv.forEach((r) => {
      const type = r[COL['Тип'] - 1];
      const cat = String(r[COL['Категория'] - 1] || '').trim();
      if (!cat || type === TRANSFER || TYPES.indexOf(type) < 0) return;
      const k = cat.toLowerCase() + '|' + type;
      if (known.has(k)) return;
      known.add(k);
      added.push([cat, type]);
    });
    if (added.length) {
      // Первая пустая строка в колонке A (в C/D/F справочника могут быть другие данные)
      const colA = ref.getRange(1, 1, ref.getMaxRows(), 1).getValues();
      let r0 = colA.length;
      while (r0 > 1 && !colA[r0 - 1][0]) r0--;
      ref.getRange(r0 + 1, 1, added.length, 2).setValues(added);
      log.push('Справочники: добавлены ' + added.map((a) => a[0] + ' (' + a[1] + ')').join(', '));
    }

    // 3. Сломанные комментарии
    if (last > 1) {
      const rng = tx.getRange(2, COL['Комментарий'], last - 1, 1);
      const f = rng.getFormulas();
      let fixed = 0;
      f.forEach((r, i) => {
        if (r[0]) { tx.getRange(i + 2, COL['Комментарий']).setValue("'" + r[0]); fixed++; }
      });
      if (fixed) log.push('Транзакции: восстановлено комментариев — ' + fixed);
    }

    // 4. Период с 10-го по 9-е; прошедшее — в «Архив»
    const pr = summaryKeyRow_(sv, 'periodEnd');
    sv.getRange(pr, 1).setValue(`Конец периода (до следующей зарплаты, не позже ${SALARY_DAY_TO}-го)`);
    sv.getRange(summaryKeyRow_(sv, 'periodStart'), 1).setValue('Начало периода (сдвигается сам)');
    if (ensurePeriod_(ss)) log.push('Период переставлен, прошедшее записано в «Архив»');

    // 5. Свод
    writeCategoryTableFormulas_(sv);
    log.push('Свод: категории считаются с учётом типа, добавлены «Факт за 7 дней» и «Обязательная»');
    if (!summaryKeyRow_(sv, 'lifeMoney')) {
      sv.insertRowsAfter(summaryKeyRow_(sv, 'savingsForecast'), 8);
      log.push('Свод: добавлен блок прогноза до зарплаты');
    }
    writeForecastFormulas_(sv);
    SpreadsheetApp.flush();

    Logger.log(log.join('; '));
  });
}

/** Таблица по категориям на «Своде»: суммы с учётом типа + факт за 7 дней + пометка «обязательная». */
function writeCategoryTableFormulas_(sv) {
  const T = "'Транзакции'";
  const data = sv.getRange(1, 1, sv.getLastRow(), 2).getValues();
  const head = data.findIndex((r) => r[0] === 'Категория' && r[1] === 'Тип') + 1;
  if (head < 1) throw new Error('На «Своде» не найдена таблица по категориям');
  const ps = `$B$${summaryKeyRow_(sv, 'periodStart')}`;
  const pe = `$B$${summaryKeyRow_(sv, 'periodEnd')}`;
  sv.getRange(head, 12, 1, 3).setValues([['Факт за 7 дней', 'Обязательная', 'Ежедневная']])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#37474F').setWrap(true);
  const cde = [];
  const j = [];
  const lm = [];
  for (let k = 0; k < 100; k++) {
    const r = head + 1 + k;
    const i = k + 2;   // строка «Справочников»
    const a = `$A${r}`;
    const base = `${T}!$E:$E,${T}!$F:$F,${a},${T}!$D:$D,$B${r}`;
    // Окно зарплаты (колонка P): у гарантированных доходов и обязательных статей последние дни периода —
    // уже следующий период
    const period = `${T}!$B:$B,">="&${ps},${T}!$B:$B,"<="&(${pe}-N($P${r}))`;
    const blank = `IF(${a}="","",`;
    cde.push([
      `=${blank}SUMIFS(${base},${T}!$C:$C,"План",${period}))`,
      `=${blank}SUMIFS(${base},${T}!$C:$C,"Факт",${period}))`,
      `=${blank}SUMIFS(${base},${T}!$C:$C,"Факт",${period},${T}!$B:$B,"<"&TODAY()))`,
    ]);
    j.push([`=${blank}COUNTIFS(${T}!$F:$F,${a},${T}!$D:$D,$B${r},${T}!$C:$C,"Факт",` +
      `${T}!$B:$B,">="&(TODAY()-30),${T}!$B:$B,"<="&TODAY()))`]);
    lm.push([
      `=${blank}SUMIFS(${base},${T}!$C:$C,"Факт",${T}!$B:$B,">="&(TODAY()-7),${T}!$B:$B,"<"&TODAY()))`,
      `=${blank}IF('Справочники'!C${i}="да","да",""))`,
    ]);
  }
  sv.getRange(head + 1, 3, 100, 3).setFormulas(cde);
  sv.getRange(head + 1, 10, 100, 1).setFormulas(j);
  sv.getRange(head + 1, 12, 100, 2).setFormulas(lm);
  sv.getRange(head + 1, 12, 100, 1).setNumberFormat('#,##0;-#,##0;""');
  // «Ежедневная» — одной формулой массива на всю колонку
  sv.getRange(head + 1, 14, 100, 1).clearContent();
  sv.getRange(head + 1, 14).setFormula(`=ARRAYFORMULA(IF($A${head + 1}:$A${head + 100}="","",` +
    `IF('Справочники'!E2:E101="да","да","")))`);
  sv.setColumnWidth(12, 100);
  sv.setColumnWidth(13, 100);
  sv.setColumnWidth(14, 100);

  // O — надёжность дохода; P — сдвиг периода в днях; Q — факт в «окне» (уже деньги следующего периода).
  sv.getRange(head, 15, 1, 3).setValues([['Надёжность', 'Окно зарплаты, дн.', 'Факт для след. периода']])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#37474F').setWrap(true);
  sv.getRange(head + 1, 15, 100, 1).clearContent();
  sv.getRange(head + 1, 15).setFormula(`=ARRAYFORMULA(IF($B${head + 1}:$B${head + 100}<>"Доход","",` +
    `IF('Справочники'!G2:G101="${GUARANTEED}","${GUARANTEED}","${EXPECTED}")))`);
  const pq = [];
  for (let k = 0; k < 100; k++) {
    const r = head + 1 + k;
    const base = `${T}!$E:$E,${T}!$F:$F,$A${r},${T}!$D:$D,$B${r}`;
    pq.push([
      `=IF($A${r}="","",IF(OR($M${r}="да",$O${r}="${GUARANTEED}"),${NEXT_INCOME_DAYS},0))`,
      `=IF(OR($A${r}="",N($P${r})=0),"",SUMIFS(${base},${T}!$C:$C,"Факт",` +
      `${T}!$B:$B,">"&(${pe}-$P${r}),${T}!$B:$B,"<="&${pe}))`,
    ]);
  }
  sv.getRange(head + 1, 16, 100, 2).setFormulas(pq);
  sv.getRange(head + 1, 17, 100, 1).setNumberFormat('#,##0;-#,##0;""');
  sv.setColumnWidths(15, 3, 110);
}

/**
 * Блок прогноза под «Накоплениями»: хватит ли денег на ежедневное (еда, такси, транспорт) до конца периода.
 * «На ежедневное» = на счетах + ещё ожидаемые доходы − все ещё не оплаченные планы, кроме ежедневных статей.
 * Темп — средние траты за 7 полных дней по ежедневным статьям. Разовые траты вне плана в темп не входят:
 * они уменьшают остаток на счетах, когда случаются, и прогноз пересчитывается.
 */
function writeForecastFormulas_(sv) {
  // Адрес таблицы категорий — каждый раз заново: вставка строк прогноза сдвигает её вниз.
  const col = (c) => {
    const head = sv.getRange(1, 1, sv.getLastRow(), 2).getValues()
      .findIndex((r) => r[0] === 'Категория' && r[1] === 'Тип') + 1;
    return `$${c}$${head + 1}:$${c}$${head + 100}`;
  };
  const cell = (key) => `$B$${summaryKeyRow_(sv, key)}`;
  const start = summaryKeyRow_(sv, 'savingsForecast') + 1;
  const rows = [
    ['', '', '', ''],
    ['Прогноз до конца периода', '', '', ''],
    ['Начало следующего периода (не позже)', `=${cell('periodEnd')}+1`, 'payday',
      'Новый период начнётся в день зарплаты, а если её не будет — на следующий день после конца'],
    ['Окно зарплаты: с этой даты', '', 'nextIncomeFrom',
      `Зарплата и обязательные платежи с ${SALARY_DAY_FROM}-го по ${SALARY_DAY_TO}-е — уже следующий период`],
    ['Деньги на ежедневное до конца периода', '', 'lifeMoney',
      'На счетах + гарантированные доходы, что ещё придут, − все неоплаченные планы, кроме ежедневных'],
    ['Можно тратить в день на ежедневное', '', 'dailyBudget', 'Деньги на ежедневное / дней осталось'],
    ['Темп ежедневных трат за 7 дней, ₸/день', '', 'pace7',
      'Статьи с «да» в «Справочниках», колонка E «Ежедневная»; 7 полных дней до сегодня'],
    ['При таком темпе деньги закончатся', '', 'runOutDate', 'Пусто — темп нулевой'],
    ['Прогноз к концу периода: запас (+) / нехватка (−)', '', 'forecastGap', 'Деньги на ежедневное − темп × дней осталось'],
    ['Ожидается (не гарантировано) до конца периода', '', 'expectedLeft', 'Доходы с пометкой «Ожидается»: долги, подработки'],
    ['Можно в день, если придёт ожидаемое', '', 'dailyBudgetMax', '(Деньги на ежедневное + ожидаемое) / дней осталось'],
    ['Запас / нехватка, если придёт ожидаемое', '', 'forecastGapMax', ''],
  ];
  // Блок мог быть короче (до v5) — добавляем недостающие строки, чтобы не затереть «Топ-5» ниже.
  // Конец блока — последняя строка, что уже есть: при повторном запуске лишнее не вставляем.
  const oldEnd = summaryKeyRow_(sv, 'forecastGapMax') || summaryKeyRow_(sv, 'forecastGap');
  const have = oldEnd ? oldEnd - start + 1 : 0;
  if (have && have < rows.length) sv.insertRowsAfter(oldEnd, rows.length - have);
  sv.getRange(start, 1, rows.length, 4).setValues(rows.map((r) => [r[0], r[1] && !r[1].startsWith('=') ? r[1] : '', r[2], r[3]]));
  sv.getRange(start + 1, 1).setFontWeight('bold').setFontSize(12);
  const set = (key, f, fmt) => {
    const r = summaryKeyRow_(sv, key);
    sv.getRange(r, 2).setFormula(f).setNumberFormat(fmt);
    sv.getRange(r, 3).setBackground('#FFF8E1');
    sv.getRange(r, 4).setFontStyle('italic').setFontColor('#777777');
  };
  set('payday', `=${cell('periodEnd')}+1`, 'dd.mm.yyyy');
  set('nextIncomeFrom', `=${cell('periodEnd')}-${NEXT_INCOME_DAYS - 1}`, 'dd.mm.yyyy');
  writeCategoryTableFormulas_(sv);
  // Деньги периода: на счетах − доход, уже пришедший для следующего периода + обязательное, уже оплаченное за него
  // + гарантированные доходы, что ещё придут в этом периоде
  const money = `${cell('balanceNow')}-SUMIF(${col('B')},"Доход",${col('Q')})+SUMIF(${col('B')},"Расход",${col('Q')})` +
    `+SUMIFS(${col('H')},${col('B')},"Доход",${col('O')},"${GUARANTEED}")`;
  set('lifeMoney', `=ROUND(${money}-SUMIFS(${col('H')},${col('B')},"Расход",${col('N')},"<>да"),0)`, '#,##0');
  const fr = summaryKeyRow_(sv, 'freeNow');
  sv.getRange(fr, 2).setFormula(`=${money}-SUMIF(${col('B')},"Расход",${col('H')})`);
  sv.getRange(fr, 4).setValue('На счетах + гарантированные доходы, что ещё придут, − все неоплаченные планы');
  set('dailyBudget', `=ROUND(${cell('lifeMoney')}/${cell('daysLeft')},0)`, '#,##0');
  set('pace7', `=ROUND(SUMIFS(${col('L')},${col('B')},"Расход",${col('N')},"да")/7,0)`, '#,##0');
  set('runOutDate', `=IF(${cell('pace7')}<=0,"",IF(${cell('lifeMoney')}<=0,TODAY(),` +
    `TODAY()+INT(${cell('lifeMoney')}/${cell('pace7')})))`, 'dd.mm.yyyy');
  set('forecastGap', `=ROUND(${cell('lifeMoney')}-${cell('pace7')}*${cell('daysLeft')},0)`, '#,##0;-#,##0');
  set('expectedLeft', `=SUMIFS(${col('H')},${col('B')},"Доход",${col('O')},"${EXPECTED}")`, '#,##0');
  set('dailyBudgetMax', `=ROUND((${cell('lifeMoney')}+${cell('expectedLeft')})/${cell('daysLeft')},0)`, '#,##0');
  set('forecastGapMax', `=ROUND(${cell('lifeMoney')}+${cell('expectedLeft')}-${cell('pace7')}*${cell('daysLeft')},0)`,
    '#,##0;-#,##0');
}

/**
 * v4 → v5: надёжность доходов. Запустить один раз; повторно — безопасно.
 * 1. «Справочники»: колонка G «Надёжность» у доходов — «Гарантировано» (Зп, Стипендия, Пособие) или «Ожидается».
 * 2. «Свод»: колонки O–Q, прогноз и «Свободно» считают только гарантированные доходы; гарантированный доход
 *    за последние NEXT_INCOME_DAYS дней периода — деньги следующего периода; ожидаемое — отдельной строкой.
 */
function migrateV5() {
  withEnglishLocale_(() => {
    const ss = SpreadsheetApp.getActive();
    const ref = ss.getSheetByName(SHEET_REF);
    const sv = ss.getSheetByName(SHEET_SUMMARY);
    const log = [];
    sv.getRange(summaryKeyRow_(sv, 'periodStart'), 1).setValue('Начало периода (день зарплаты, ставится сам)');
    sv.getRange(summaryKeyRow_(sv, 'periodEnd'), 1)
      .setValue(`Конец периода (до следующей зарплаты, не позже ${SALARY_DAY_TO}-го)`);
    if (ensurePeriod_(ss)) log.push('Период: с дня зарплаты');
    if (ref.getRange('G1').getValue() !== 'Надёжность') {
      ref.getRange('G1').setValue('Надёжность');
      ref.getRange('A1').copyFormatToRange(ref, 7, 7, 1, 1);
      const n = Math.max(ref.getLastRow() - 1, 1);
      const names = ref.getRange(2, 1, n, 2).getValues();
      ref.getRange(2, 7, n, 1).setValues(names.map((r) => [r[1] !== 'Доход' ? ''
        : GUARANTEED_DEFAULT.indexOf(String(r[0])) >= 0 ? GUARANTEED : EXPECTED]));
      ref.getRange(2, 7, Math.min(ref.getMaxRows() - 1, 499), 1).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList([GUARANTEED, EXPECTED], true).setAllowInvalid(true).build());
      ref.setColumnWidth(7, 120);
      log.push('Справочники: колонка «Надёжность»');
    }
    writeCategoryTableFormulas_(sv);
    writeForecastFormulas_(sv);
    SpreadsheetApp.flush();
    log.push('Свод: прогноз и «Свободно» — по гарантированным доходам, ожидаемое отдельно');
    rebuildPlanCalendar();
    log.push('Календарь план: колонка «Надёжность»');
    Logger.log(log.join('; '));
  });
}

// ---------- Период и архив ----------

/**
 * Текущий период: начинается в день первой зарплаты («Зп»), пришедшей с SALARY_DAY_FROM по SALARY_DAY_TO число,
 * и длится до следующей зарплаты, но не дольше SALARY_DAY_TO-го следующего месяца. Нет зарплаты к SALARY_DAY_TO-му —
 * период начинается на следующий день. Пока в окне зарплата не пришла, действует прошлый период.
 */
function currentPeriod_(ss, tz) {
  const today = ymd_(new Date(), tz);
  const salary = readTx_(ss).map((x) => txObj_(x.v, tz))
    .filter((t) => t.status === 'Факт' && t.type === 'Доход' && t.date <= today &&
      String(t.category).trim().toLowerCase() === SALARY_CATEGORY.toLowerCase())
    .map((t) => t.date);
  const ymd = (y, m, d) => new Date(Date.UTC(y, m - 1, d)).toISOString().slice(0, 10);
  const startIn = (y, m) => {           // начало периода в месяце m или null, если окно ещё идёт и зарплаты нет
    const from = ymd(y, m, SALARY_DAY_FROM);
    const to = ymd(y, m, SALARY_DAY_TO);
    const hit = salary.filter((d) => d >= from && d <= to).sort()[0];
    if (hit) return hit;
    return today > to ? addDays_(to, 1) : null;
  };
  let [y, m, d] = today.split('-').map(Number);
  if (d < SALARY_DAY_FROM) m -= 1;
  let start = startIn(y, m);
  if (!start) start = startIn(y, m - 1);   // окно этого месяца идёт, зарплаты ещё нет — живём в прошлом периоде
  const [sy, sm] = start.split('-').map(Number);
  return { start: start, end: ymd(sy, sm + 1, SALARY_DAY_TO) };
}

/**
 * Если сегодня уже вне периода «Свода» — записать прошедшее в «Архив» и переставить период.
 * Возвращает true, если период менялся. Транзакции не трогает.
 */
function ensurePeriod_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const sv = ss.getSheetByName(SHEET_SUMMARY);
  const rs = summaryKeyRow_(sv, 'periodStart');
  const re = summaryKeyRow_(sv, 'periodEnd');
  const oldStart = sv.getRange(rs, 2).getValue();
  const oldEnd = sv.getRange(re, 2).getValue();
  const cur = currentPeriod_(ss, tz);
  if (isDate_(oldStart) && isDate_(oldEnd) && ymd_(oldStart, tz) === cur.start && ymd_(oldEnd, tz) === cur.end) return false;

  const os = isDate_(oldStart) ? ymd_(oldStart, tz) : '';
  if (os && os < cur.start) archiveRange_(ss, os, addDays_(cur.start, -1));
  else if (os && os > cur.start) trimArchive_(ss, cur.start);   // начало сдвинулось раньше — архив не должен его задевать
  sv.getRange(rs, 2).setValue(parseYmd_(cur.start));
  sv.getRange(re, 2).setValue(parseYmd_(cur.end));
  SpreadsheetApp.flush();
  try { rebuildPlanCalendar(); } catch (err) { /* сетку можно перестроить из меню */ }
  return true;
}

/** То же для веб-приложения: под блокировкой, чтобы два телефона не сдвинули период дважды. */
function ensurePeriodLocked_() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const sv = ss.getSheetByName(SHEET_SUMMARY);
  const start = sv.getRange(summaryKeyRow_(sv, 'periodStart'), 2).getValue();
  const end = sv.getRange(summaryKeyRow_(sv, 'periodEnd'), 2).getValue();
  const cur = currentPeriod_(ss, tz);
  if (isDate_(start) && isDate_(end) && ymd_(start, tz) === cur.start && ymd_(end, tz) === cur.end) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { ensurePeriod_(ss); } finally { lock.releaseLock(); }
}

/**
 * Итоги диапазона дат в «Архив»: по каждой статье план, факт, отклонение и итоговые строки.
 * Суммы считаются из «Транзакций» и записываются числами — архив не меняется задним числом.
 * Повторно тот же диапазон не пишется.
 */
function archiveRange_(ss, from, to) {
  const tz = ss.getSpreadsheetTimeZone();
  let sh = ss.getSheetByName(SHEET_ARCHIVE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ARCHIVE, ss.getSheets().length);
    sh.getRange(1, 1, 1, 8).setValues([['Период', 'С', 'По', 'Статья', 'Тип', 'План', 'Факт', 'Факт − План']])
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#37474F');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 120);
    sh.setColumnWidth(4, 190);
  }
  const label = `${from.slice(8, 10)}.${from.slice(5, 7)}.${from.slice(0, 4)}–${to.slice(8, 10)}.${to.slice(5, 7)}.${to.slice(0, 4)}`;
  if (sh.getLastRow() > 1 && sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().some((r) => r[0] === label)) return;

  const agg = {};
  readTx_(ss).forEach((x) => {
    const t = txObj_(x.v, tz);
    if (t.type === TRANSFER || t.date < from || t.date > to) return;
    const k = String(t.category) + '|' + t.type;
    agg[k] = agg[k] || { name: String(t.category), type: t.type, plan: 0, fact: 0 };
    if (t.status === 'План') agg[k].plan += t.amount;
    else if (t.status === 'Факт') agg[k].fact += t.amount;
  });
  const items = Object.values(agg).sort((a, b) => (a.type === b.type ? b.fact - a.fact : a.type === 'Доход' ? -1 : 1));
  const f = parseYmd_(from);
  const t2 = parseYmd_(to);
  const rows = items.map((i) => [label, f, t2, i.name, i.type, i.plan, i.fact, i.fact - i.plan]);
  ['Доход', 'Расход'].forEach((type) => {
    const list = items.filter((i) => i.type === type);
    const pl = list.reduce((a, i) => a + i.plan, 0);
    const fa = list.reduce((a, i) => a + i.fact, 0);
    rows.push([label, f, t2, type === 'Доход' ? 'ИТОГО доходы' : 'ИТОГО расходы', type, pl, fa, fa - pl]);
  });
  const r0 = sh.getLastRow() + 1;
  sh.getRange(r0, 1, rows.length, 8).setValues(rows);
  sh.getRange(r0, 2, rows.length, 2).setNumberFormat('dd.mm.yyyy');
  sh.getRange(r0, 6, rows.length, 3).setNumberFormat('#,##0;-#,##0;0');
  sh.getRange(r0 + rows.length - 2, 1, 2, 8).setFontWeight('bold').setBackground('#ECEFF1');
}

/**
 * Начало периода сдвинулось раньше конца уже записанного в «Архив» периода: блок архива, который задевает новый
 * период, удаляется и записывается заново — до дня перед новым началом.
 */
function trimArchive_(ss, newStart) {
  const sh = ss.getSheetByName(SHEET_ARCHIVE);
  if (!sh || sh.getLastRow() < 2) return;
  const tz = ss.getSpreadsheetTimeZone();
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  const labels = {};
  data.forEach((r, i) => {
    if (!isDate_(r[1]) || !isDate_(r[2]) || ymd_(r[2], tz) < newStart) return;
    (labels[r[0]] = labels[r[0]] || { from: ymd_(r[1], tz), rows: [] }).rows.push(i + 2);
  });
  Object.keys(labels).forEach((label) => {
    const b = labels[label];
    b.rows.sort((x, y) => y - x).forEach((row) => sh.deleteRow(row));
    if (b.from < newStart) archiveRange_(ss, b.from, addDays_(newStart, -1));
  });
}

// ---------- Календарь план ----------
//
// Редактируемая сетка «категория × день». Источник правды — плановые строки в «Транзакциях»:
// сетка из них строится (rebuildPlanCalendar), а правка ячейки сразу переписывает план этого дня
// (onEdit → syncPlanCells_). Поэтому приложение, «Свод» и прогноз видят изменения без доработок.

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Бюджет')
    .addItem('Обновить Календарь план', 'rebuildPlanCalendar')
    .addToUi();
  try { ensurePeriod_(SpreadsheetApp.getActive()); } catch (err) { /* не мешаем открыть таблицу */ }
  try { rebuildPlanCalendar(); } catch (err) { /* не мешаем открыть таблицу */ }
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET_PLAN_CAL) return;
  if (e.range.getLastRow() < PC_FIRST) return;
  if (e.range.getColumn() <= PC_REL_COL && e.range.getLastColumn() >= PC_REL_COL) syncReliability_(sh, e.range);
  if (e.range.getLastColumn() >= PC_DAY_COL) syncPlanCells_(sh, e.range);
}

/** Надёжность дохода, изменённая в «Календаре план», записывается в «Справочники» (колонка G). */
function syncReliability_(sh, range) {
  const ss = sh.getParent();
  const ref = ss.getSheetByName(SHEET_REF);
  const r0 = Math.max(range.getRow(), PC_FIRST);
  const rows = sh.getRange(r0, 1, range.getLastRow() - r0 + 1, PC_REL_COL).getValues();
  const n = Math.max(ref.getLastRow() - 1, 1);
  const refv = ref.getRange(2, 1, n, 2).getValues();
  rows.forEach((r) => {
    const v = r[PC_REL_COL - 1];
    if (r[1] !== 'Доход' || [GUARANTEED, EXPECTED].indexOf(v) < 0) return;
    const i = refv.findIndex((x) => String(x[0]).trim().toLowerCase() === String(r[0]).trim().toLowerCase() && x[1] === 'Доход');
    if (i >= 0) ref.getRange(i + 2, 7).setValue(v);
  });
}

/** Перестраивает «Календарь план» из плановых строк «Транзакций» на текущий период «Свода». */
function rebuildPlanCalendar() {
  {
    const ss = SpreadsheetApp.getActive();
    const tz = ss.getSpreadsheetTimeZone();
    const sv = ss.getSheetByName(SHEET_SUMMARY);
    const ps = sv.getRange(summaryKeyRow_(sv, 'periodStart'), 2).getValue();
    const pe = sv.getRange(summaryKeyRow_(sv, 'periodEnd'), 2).getValue();
    if (!isDate_(ps) || !isDate_(pe)) throw new Error('На «Своде» не заданы начало и конец периода');
    const days = [];
    for (let d = ymd_(ps, tz); d <= ymd_(pe, tz); d = addDays_(d, 1)) days.push(d);

    // Категории: сначала доходы, потом расходы — в порядке справочника
    const refRows = ss.getSheetByName(SHEET_REF).getRange(2, 1, 200, 7).getValues().filter((r) => r[0]);
    const cats = refRows.filter((r) => r[1] === 'Доход').concat(refRows.filter((r) => r[1] !== 'Доход'))
      .map((r) => {
        // Колонка «Пометка»: у доходов — надёжность (правится), у расходов — «ежедневная», если так в справочнике
        const rel = r[1] === 'Доход' ? (r[6] || EXPECTED) : (r[4] === 'да' ? DAILY_MARK : '');
        // Окно зарплаты — как на «Своде»: у гарантированных доходов и обязательных статей последние дни периода
        // уже относятся к следующему
        const shift = (r[1] === 'Доход' ? rel === GUARANTEED : r[2] === 'да') ? NEXT_INCOME_DAYS : 0;
        return { name: String(r[0]), type: r[1] || 'Расход', rel: rel, shift: shift };
      });

    // Планы и факты по дням (факты — для ежедневных статей, они сверяются день за днём)
    const plan = {};
    const fact = {};
    readTx_(ss).forEach((x) => {
      const t = txObj_(x.v, tz);
      if (t.type === TRANSFER) return;
      const k = String(t.category).toLowerCase() + '|' + t.type + '|' + t.date;
      if (t.status === 'План') plan[k] = (plan[k] || 0) + t.amount;
      else if (t.status === 'Факт') fact[k] = (fact[k] || 0) + t.amount;
    });

    let sh = ss.getSheetByName(SHEET_PLAN_CAL);
    if (!sh) sh = ss.insertSheet(SHEET_PLAN_CAL, ss.getSheets().length);
    const width = PC_DAY_COL - 1 + days.length;
    if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
    const height = PC_FIRST - 1 + cats.length;
    if (sh.getMaxRows() < height + 5) sh.insertRowsAfter(sh.getMaxRows(), height + 5 - sh.getMaxRows());
    sh.clear();
    sh.clearConditionalFormatRules();
    const F = formulaDialect_(sh);

    const lastCol = colLetter_(width);
    const firstDay = colLetter_(PC_DAY_COL);
    const n = cats.length;
    const lastRow = PC_FIRST + n - 1;

    sh.getRange(1, 1).setValue('Календарь план').setFontWeight('bold').setFontSize(14);
    sh.getRange(2, 1).setValue('Впиши плановые суммы в ячейки дней — они сразу попадут в план. ' +
      'Серые зачёркнутые — уже покрыты фактом, оранжевые — покрыты частично; ' +
      'ежедневные статьи сверяются день за днём: красный день — потратил больше плана. ' +
      '«Осталось» красным — по статье потрачено больше плана, заложи ещё. ' +
      '«Надёжность» у доходов: «Гарантировано» входит в прогноз, «Ожидается» — нет.').setFontStyle('italic').setFontColor('#777777');

    // Шапка
    const head = ['Категория', 'Тип', 'План', 'Потрачено', 'Осталось', 'Пометка'].concat(days.map((d) => parseYmd_(d)));
    sh.getRange(PC_HEAD, 1, 1, head.length).setValues([head])
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#37474F').setHorizontalAlignment('center');
    sh.getRange(PC_HEAD, PC_DAY_COL, 1, days.length).setNumberFormat('dd.mm');
    const wd = ['вс', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'];
    sh.getRange(PC_HEAD - 1, PC_DAY_COL, 1, days.length)
      .setValues([days.map((d) => wd[parseYmd_(d).getDay()])]).setFontColor('#777777').setHorizontalAlignment('center');

    // Итоги по дням
    const tot = [['Итого доходы', 'Доход'], ['Итого расходы', 'Расход']].map(([label, type]) =>
      [label, '', `=SUMIF($B$${PC_FIRST}:$B$${lastRow},"${type}",C${PC_FIRST}:C${lastRow})`,
        `=SUMIF($B$${PC_FIRST}:$B$${lastRow},"${type}",D${PC_FIRST}:D${lastRow})`,
        `=SUMIF($B$${PC_FIRST}:$B$${lastRow},"${type}",E${PC_FIRST}:E${lastRow})`, '']
        .concat(days.map((_, i) => {
          const c = colLetter_(PC_DAY_COL + i);
          return `=SUMIF($B$${PC_FIRST}:$B$${lastRow},"${type}",${c}${PC_FIRST}:${c}${lastRow})`;
        })));
    sh.getRange(PC_HEAD + 1, 1, 2, head.length).setFormulas(tot.map((r) => r.map((v) => (String(v).startsWith('=') ? F(v) : ''))));
    sh.getRange(PC_HEAD + 1, 1, 2, 1).setValues([['Итого доходы'], ['Итого расходы']]);
    sh.getRange(PC_HEAD + 1, 1, 2, head.length).setFontWeight('bold').setBackground('#ECEFF1');

    // Категории: имя, тип, формулы, планы по дням
    const T = "'Транзакции'";
    const rows = cats.map((c, i) => {
      const r = PC_FIRST + i;
      return [c.name, c.type,
        `=SUM(${firstDay}${r}:${lastCol}${r})`,
        `=SUMIFS(${T}!$E:$E,${T}!$F:$F,$A${r},${T}!$D:$D,$B${r},${T}!$C:$C,"Факт",` +
        `${T}!$B:$B,">="&$${firstDay}$${PC_HEAD},${T}!$B:$B,"<="&($${lastCol}$${PC_HEAD}-${c.shift}))`,
        `=C${r}-D${r}`, c.rel]
        .concat(days.map((d) => plan[c.name.toLowerCase() + '|' + c.type + '|' + d] || ''));
    });
    if (n) {
      sh.getRange(PC_FIRST, 1, n, 2).setValues(rows.map((r) => r.slice(0, 2)));
      sh.getRange(PC_FIRST, 3, n, 3).setFormulas(rows.map((r) => r.slice(2, 5).map(F)));
      sh.getRange(PC_FIRST, PC_REL_COL, n, 1).setValues(rows.map((r) => [r[5]]));
      // Выпадающий список надёжности — только у доходов
      const relRule = SpreadsheetApp.newDataValidation().requireValueInList([GUARANTEED, EXPECTED], true)
        .setAllowInvalid(false).build();
      cats.forEach((c, i) => {
        if (c.type === 'Доход') sh.getRange(PC_FIRST + i, PC_REL_COL).setDataValidation(relRule).setBackground('#FFF8E1');
      });
      sh.getRange(PC_FIRST, PC_DAY_COL, n, days.length).setValues(rows.map((r) => r.slice(6)));
      // Фон и зачёркивание. Ежедневные статьи — день за днём: прошедший день в плане — серый,
      // с перерасходом — красный; сегодня — оранжевый, если уже тратили; будущие дни не трогаем.
      const today = ymd_(new Date(), tz);
      const bg = [];
      const line = [];
      cats.forEach((c) => {
        const b = [];
        const l = [];
        days.forEach((d) => {
          let color = '#E3F2FD';
          let strike = 'none';
          if (c.rel === DAILY_MARK && d <= today) {
            const k = c.name.toLowerCase() + '|' + c.type + '|' + d;
            const p = plan[k] || 0;
            const f = fact[k] || 0;
            if (f > p) color = '#FFCDD2';
            else if (d < today && p > 0) { color = '#E0E0E0'; strike = 'line-through'; }
            else if (d === today && f > 0) color = '#FFE0B2';
          }
          b.push(color);
          l.push(strike);
        });
        bg.push(b);
        line.push(l);
      });
      sh.getRange(PC_FIRST, PC_DAY_COL, n, days.length).setBackgrounds(bg).setFontLines(line);
      cats.forEach((c, i) => {
        if (c.rel === DAILY_MARK) sh.getRange(PC_FIRST + i, PC_REL_COL).setFontColor('#777777');
      });
    }
    sh.getRange(PC_HEAD + 1, 3, n + 2, head.length - 2).setNumberFormat('#,##0;-#,##0;""');

    // Подсветка
    const grid = sh.getRange(PC_FIRST, PC_DAY_COL, Math.max(n, 1), days.length);
    const f = `${firstDay}${PC_FIRST}`;
    const cum = `SUM($${firstDay}${PC_FIRST}:${f})`;
    const rules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(F(`=AND(${f}<>"",$F${PC_FIRST}<>"${DAILY_MARK}",${cum}<=$D${PC_FIRST})`))
        .setBackground('#E0E0E0').setFontColor('#9E9E9E').setStrikethrough(true).setRanges([grid]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(F(`=AND(${f}<>"",$F${PC_FIRST}<>"${DAILY_MARK}",${cum}-${f}<$D${PC_FIRST},${cum}>$D${PC_FIRST})`))
        .setBackground('#FFE0B2').setRanges([grid]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(F(`=AND($B${PC_FIRST}="Расход",$E${PC_FIRST}<0)`))
        .setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true)
        .setRanges([sh.getRange(PC_FIRST, 5, Math.max(n, 1), 1)]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(F(`=${firstDay}$${PC_HEAD}=TODAY()`))
        .setBackground('#FFF59D').setRanges([sh.getRange(PC_HEAD, PC_DAY_COL, 1, days.length)]).build(),
    ];
    sh.setConditionalFormatRules(rules);

    sh.setFrozenRows(PC_HEAD + 2);
    sh.setFrozenColumns(PC_REL_COL);
    sh.setColumnWidth(1, 190);
    sh.setColumnWidth(2, 60);
    for (let c = 3; c <= 5; c++) sh.setColumnWidth(c, 85);
    sh.setColumnWidth(PC_REL_COL, 115);
    sh.setColumnWidths(PC_DAY_COL, days.length, 62);
    sh.getRange(PC_HEAD, 3, 1, 3).setNote('План / Потрачено / Осталось — за весь период со «Свода»');
  }
}

/**
 * Какой разделитель аргументов понимает таблица: формулы из скрипта в русской локали с «,» не разбираются.
 * Проверяем на пробной формуле, локаль не переключаем: смена локали перезагружает открытую таблицу,
 * а сетка перестраивается при каждом открытии — получилась бы бесконечная перезагрузка.
 */
function formulaDialect_(sh) {
  const probe = sh.getRange(1, 1);
  for (const sep of [',', ';']) {
    probe.setFormula(`=SUM(1${sep}2)`);
    SpreadsheetApp.flush();
    const ok = probe.getDisplayValue() === '3';
    probe.clearContent();
    if (ok) return sep === ',' ? (f) => f : (f) => f.replace(/,/g, ';');
  }
  throw new Error('Не удалось определить разделитель формул');
}

/**
 * Правка ячеек сетки: для каждой ячейки удаляет плановые строки «Транзакций» этой категории и дня
 * и, если в ячейке число, записывает одну новую. Даты берутся из шапки, категория и тип — из колонок A и B.
 */
function syncPlanCells_(sh, range) {
  const ss = sh.getParent();
  const tz = ss.getSpreadsheetTimeZone();
  // В простом триггере блокировка может быть недоступна — тогда работаем без неё.
  let lock = null;
  try { lock = LockService.getScriptLock(); lock.waitLock(20000); } catch (err) { lock = null; }
  try {
    const tx = ss.getSheetByName(SHEET_TX);
    const r0 = Math.max(range.getRow(), PC_FIRST);
    const c0 = Math.max(range.getColumn(), PC_DAY_COL);
    const r1 = range.getLastRow();
    const c1 = range.getLastColumn();
    const dates = sh.getRange(PC_HEAD, c0, 1, c1 - c0 + 1).getValues()[0];
    const cats = sh.getRange(r0, 1, r1 - r0 + 1, 2).getValues();
    const vals = sh.getRange(r0, c0, r1 - r0 + 1, c1 - c0 + 1).getValues();

    const wanted = {};   // ключ «категория|тип|дата» → сумма (0 — удалить)
    cats.forEach((c, i) => {
      if (!c[0]) return;
      dates.forEach((d, j) => {
        if (!isDate_(d)) return;
        const v = Math.round(num_(vals[i][j]));
        wanted[String(c[0]).toLowerCase() + '|' + (c[1] || 'Расход') + '|' + ymd_(d, tz)] =
          { name: String(c[0]), type: c[1] || 'Расход', date: ymd_(d, tz), amount: v > 0 ? v : 0 };
      });
    });

    // Удаляем старые планы этих ячеек снизу вверх
    const rows = readTx_(ss);
    for (let i = rows.length - 1; i >= 0; i--) {
      const t = txObj_(rows[i].v, tz);
      if (t.status !== 'План') continue;
      if (wanted[String(t.category).toLowerCase() + '|' + t.type + '|' + t.date]) tx.deleteRow(rows[i].row);
    }
    // Пишем новые
    const add = Object.values(wanted).filter((w) => w.amount > 0).map((w) => [
      'g-' + Utilities.getUuid().slice(0, 12), parseYmd_(w.date), 'План', w.type, w.amount, w.name,
      '', '', 'из «Календаря план»', new Date(), '']);
    if (add.length) tx.getRange(tx.getLastRow() + 1, 1, add.length, add[0].length).setValues(add);
  } finally {
    if (lock) lock.releaseLock();
  }
}

function colLetter_(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
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

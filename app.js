'use strict';

// ---------- хранилище ----------
const LS = {
  get(k, d) { try { const v = localStorage.getItem('b.' + k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem('b.' + k, JSON.stringify(v)); } catch {} },
};

const state = {
  url: LS.get('url', ''),
  who: LS.get('who', ''),
  snap: LS.get('snap', null),          // последний ответ сервера
  snapAt: LS.get('snapAt', 0),         // когда получен
  queue: LS.get('queue', []),          // неотправленные операции: {op:'add'|'update'|'delete', tx?, id, fields?}
  type: 'Расход',
  amount: '',
  category: '',
  plan: false,
  editing: null,                       // запись, которую сейчас меняем
  fromPlan: false,                     // ввод открыт тапом по ожидаемой позиции
};

const TRANSFER = 'Перевод';            // тип записи «перевод между счетами»: не трата и не доход
const $ = (id) => document.getElementById(id);
const fmt = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 0 });
const money = (n) => fmt.format(Math.round(Number(n) || 0)).replace(/ | /g, ' ');

function todayYmd() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
function shortDate(ymd) {
  const [, m, d] = ymd.split('-');
  return `${d}.${m}`;
}
function dateLabel(ymd) {
  const t = todayYmd();
  if (ymd === t) return 'Сегодня';
  const y = new Date(); y.setDate(y.getDate() - 1);
  const yy = `${y.getFullYear()}-${String(y.getMonth() + 1).padStart(2, '0')}-${String(y.getDate()).padStart(2, '0')}`;
  if (ymd === yy) return 'Вчера';
  return shortDate(ymd);
}
function uid() {
  return (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));
}

// ---------- экраны ----------
function show(id) {
  for (const s of document.querySelectorAll('.screen')) s.hidden = s.id !== id;
}

/**
 * Своя история переходов, чтобы жест «назад» на телефоне возвращал по экранам,
 * а не закрывал приложение. Корень — экран ввода; с него «назад» выходит.
 * sub: экран ввода, открытый из сводки (изменение записи или оплата ожидаемой позиции).
 */
function go(screen, sub) {
  history.pushState({ screen: screen, sub: !!sub }, '');
  show(screen);
}
function currentNav() { return history.state || { screen: 'entry' }; }

window.addEventListener('popstate', (e) => {
  const s = e.state || { screen: 'entry' };
  $('cat-sheet').hidden = true;
  $('acc-sheet').hidden = true;
  if (!s.sub && (state.editing || state.fromPlan)) resetEntry();   // ушли из режима правки без сохранения
  show(s.screen);
  if (s.screen === 'summary') renderSummary();
  if (s.screen === 'report') renderReport();
  // Выбор в листе делается после того, как лист закрылся: иначе два history.back() подряд.
  const after = state.afterSheet;
  state.afterSheet = null;
  if (after) after();
});

let toastTimer;
function toast(msg) {
  const t = $('toast');
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.hidden = true; }, 2200);
}

// ---------- сеть ----------
async function apiGet() {
  const r = await fetch(state.url, { cache: 'no-store' });
  const j = await r.json();
  if (!j.ok) throw new Error(j.error || 'Ошибка сервера');
  return j;
}
async function apiPost(body) {
  const r = await fetch(state.url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body),
  });
  const j = await r.json();
  if (!j.ok) { const e = new Error(j.error || 'Ошибка сервера'); e.server = true; throw e; }
  return j;
}

function applySnap(j) {
  state.snap = j;
  state.snapAt = Date.now();
  LS.set('snap', j);
  LS.set('snapAt', state.snapAt);
  renderAll();
}

async function refresh() {
  if (!state.url) return;
  try {
    applySnap(await apiGet());
  } catch {
    renderStamp(true);
  }
}

let flushing = false;
async function flush() {
  if (flushing || !state.queue.length || !state.url) return;
  flushing = true;
  try {
    while (state.queue.length) {
      const item = state.queue[0];
      let j;
      try {
        if (item.op === 'add') j = await apiPost({ action: 'add', tx: item.tx });
        else if (item.op === 'update') j = await apiPost({ action: 'update', id: item.id, fields: item.fields });
        else if (item.op === 'setBalance') j = await apiPost({ action: 'setBalance', account: item.account, value: item.value, who: item.who });
        else j = await apiPost({ action: 'delete', id: item.id });
      } catch (e) {
        if (e.server) {                // сервер отказал — повторять бессмысленно
          toast('Не записано: ' + e.message);
          state.queue.shift();
          LS.set('queue', state.queue);
          continue;
        }
        break;                         // нет сети — попробуем позже
      }
      state.queue.shift();
      LS.set('queue', state.queue);
      applySnap(j);
    }
  } finally {
    flushing = false;
    renderAll();
  }
}

function enqueue(item) {
  state.queue.push(item);
  LS.set('queue', state.queue);
  renderAll();
  flush();
}

// ---------- ввод ----------
function categoriesOf(type) {
  const all = (state.snap && state.snap.categories) || [];
  return all.filter((c) => c.type === type).map((c) => c.name);
}

/** 5 частых за 30 дней (по «Своду»), остальное добивается первыми из справочника. */
function chipCategories() {
  const list = categoriesOf(state.type);
  if (state.type === 'Доход') return list.slice(0, 7);
  const top = ((state.snap && state.snap.top) || []).filter((t) => t.count > 0).map((t) => t.category)
    .filter((c) => list.includes(c));
  const out = top.slice(0, 5);
  for (const c of list) { if (out.length >= 5) break; if (!out.includes(c)) out.push(c); }
  if (state.category && !out.includes(state.category)) out.unshift(state.category);
  return out;
}

function renderChips() {
  const box = $('chips');
  box.innerHTML = '';
  for (const c of chipCategories()) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'chip' + (c === state.category ? ' on' : '');
    b.textContent = c;
    b.onclick = () => { state.category = state.category === c ? '' : c; renderEntry(); };
    box.appendChild(b);
  }
  const more = document.createElement('button');
  more.type = 'button';
  more.className = 'chip more';
  more.textContent = 'Ещё…';
  more.onclick = openCategorySheet;
  box.appendChild(more);
}

function fillSelect(sel, accounts, cur) {
  sel.innerHTML = '';
  for (const a of accounts) {
    const o = document.createElement('option');
    o.value = o.textContent = a;
    sel.appendChild(o);
  }
  if (accounts.includes(cur)) sel.value = cur;
}

function renderAccounts() {
  const accounts = (state.snap && state.snap.accounts) || [];
  const def = LS.get('lastAccount', '') || (state.who === 'Арина' ? 'Kaspi Арина' : 'Kaspi Егор');
  fillSelect($('account'), accounts, $('account').value || def);
  const from = $('from-account').value || def;
  fillSelect($('from-account'), accounts, from);
  fillSelect($('to-account'), accounts, $('to-account').value || accounts.find((a) => a !== from) || '');
}

function renderEntry() {
  const v = state.amount ? money(state.amount) : '0';
  $('amount-value').textContent = v;
  $('amount').classList.toggle('empty', !state.amount);
  const isT = state.type === TRANSFER;
  $('amount').classList.toggle('income', state.type === 'Доход');
  $('amount').classList.toggle('transfer-mode', isT);
  for (const b of $('type-seg').children) b.classList.toggle('on', b.dataset.v === state.type);
  $('chips').hidden = isT;
  $('transfer').hidden = !isT;
  $('plan-toggle').hidden = isT;
  const d = $('date').value || todayYmd();
  $('date-label').textContent = dateLabel(d);
  $('date').parentElement.classList.toggle('changed', d !== todayYmd());
  $('plan-toggle').setAttribute('aria-pressed', String(state.plan));
  const ready = isT
    ? $('from-account').value && $('to-account').value && $('from-account').value !== $('to-account').value
    : state.category;
  $('save').disabled = !(Number(state.amount) > 0 && ready);
  $('save').textContent = state.editing ? 'Сохранить изменения' : isT ? 'Перевести' : state.plan ? 'Записать в план' : 'Записать';
  $('edit-bar').hidden = !state.editing;
  if (!isT) renderChips();
}

function press(k) {
  if (k === 'back') state.amount = state.amount.slice(0, -1);
  else if ((state.amount + k).replace(/^0+/, '').length <= 9) state.amount = (state.amount + k).replace(/^0+/, '');
  if (navigator.vibrate) navigator.vibrate(8);
  renderEntry();
}

function save(account) {
  if ($('save').disabled) return;
  const amount = Number(state.amount);
  const isT = state.type === TRANSFER;
  // Факт расхода/дохода: счёт выбирается в листе после «Записать». План и перевод — без листа.
  if (!isT && !state.plan && account === undefined) {
    if (!state.editing && !confirmNotDuplicate(amount)) return;
    return openAccountSheet();
  }
  const fields = isT ? {
    date: $('date').value || todayYmd(),
    status: 'Факт',
    type: TRANSFER,
    amount: amount,
    category: TRANSFER,
    account: $('from-account').value,
    toAccount: $('to-account').value,
    comment: $('comment').value.trim(),
  } : {
    date: $('date').value || todayYmd(),
    status: state.plan ? 'План' : 'Факт',
    type: state.type,
    amount: amount,
    category: state.category,
    account: state.plan ? '' : account,
    toAccount: '',
    comment: $('comment').value.trim(),
  };
  const label = isT ? `${fields.account} → ${fields.toAccount}` : fields.category;
  if (state.editing) {
    const id = state.editing.id;
    const queued = state.queue.find((q) => q.op === 'add' && q.tx.id === id);
    if (queued) {                      // ещё не ушло на сервер — правим прямо в очереди
      Object.assign(queued.tx, fields);
      LS.set('queue', state.queue);
      renderAll();
    } else {
      enqueue({ op: 'update', id: id, fields: fields });
    }
    toast(`Изменено: ${money(amount)} · ${label}`);
  } else {
    const tx = Object.assign({ id: uid(), who: state.who }, fields);
    if (!isT && tx.account) LS.set('lastAccount', tx.account);
    enqueue({ op: 'add', tx: tx, at: Date.now() });
    toast(`${isT ? 'Перевод' : tx.status === 'План' ? 'В план' : 'Записано'}: ${money(amount)} · ${label}`);
  }
  const back = currentNav().sub;
  resetEntry();
  if (back) history.back();
}

function resetEntry() {
  state.editing = null;
  state.fromPlan = false;
  state.amount = '';
  state.category = '';
  state.plan = false;
  if (state.type === TRANSFER) state.type = 'Расход';   // перевод — разовое действие
  $('comment').value = '';
  $('date').value = todayYmd();
  const last = LS.get('lastAccount', '');
  if (last && [...$('account').options].some((o) => o.value === last)) $('account').value = last;
  renderEntry();
}

function startEdit(t) {
  state.editing = t;
  state.type = t.type || 'Расход';
  state.category = t.category;
  state.amount = String(Math.round(t.amount));
  state.plan = t.status === 'План';
  $('comment').value = t.comment || '';
  $('date').value = t.date;
  const has = (sel, v) => v && [...sel.options].some((o) => o.value === v);
  if (has($('account'), t.account)) $('account').value = t.account;
  if (t.type === TRANSFER) {
    if (has($('from-account'), t.account)) $('from-account').value = t.account;
    if (has($('to-account'), t.toAccount)) $('to-account').value = t.toAccount;
  }
  go('entry', true);
  renderEntry();
}

function openCategorySheet() {
  const list = $('cat-list');
  list.innerHTML = '';
  const all = categoriesOf(state.type);
  const top = chipCategories().filter((c) => all.includes(c));
  const add = (text, cls, fn) => {
    const li = document.createElement('li');
    li.textContent = text;
    if (cls) li.className = cls;
    if (fn) li.onclick = fn;
    list.appendChild(li);
  };
  const pick = (c) => () => { state.category = c; closeSheet(); renderEntry(); };
  if (state.type === 'Расход' && top.length) {
    add('Частые', 'sep');
    top.forEach((c) => add(c, '', pick(c)));
    add('Все', 'sep');
  }
  all.filter((c) => !top.includes(c) || state.type === 'Доход').sort((a, b) => a.localeCompare(b, 'ru'))
    .forEach((c) => add(c, '', pick(c)));
  add('+ новая', 'new', () => {
    const name = (prompt('Новая категория') || '').trim();
    if (!name) return;
    // В справочник она попадёт на сервере вместе с первой записью.
    if (!all.some((c) => c.toLowerCase() === name.toLowerCase())) {
      state.snap = state.snap || { categories: [] };
      state.snap.categories.push({ name: name, type: state.type });
      LS.set('snap', state.snap);
    }
    state.category = all.find((c) => c.toLowerCase() === name.toLowerCase()) || name;
    closeSheet();
    renderEntry();
  });
  $('cat-sheet').hidden = false;
  history.pushState(Object.assign({}, currentNav(), { sheet: true }), '');
}
/** Та же сумма в той же категории за ±3 дня — спросить, не дубль ли это. */
function confirmNotDuplicate(amount) {
  const date = $('date').value || todayYmd();
  const dayNum = (s) => Date.UTC(...s.split('-').map((x, i) => Number(x) - (i === 1 ? 1 : 0)));
  const same = feedData().facts.find((t) => t.type === state.type && t.category === state.category &&
    Math.round(t.amount) === Math.round(amount) && Math.abs(dayNum(t.date) - dayNum(date)) <= 3 * 864e5);
  if (!same) return true;
  return confirm(`Похожая запись уже есть: ${money(same.amount)} · ${same.category} · ${shortDate(same.date)}` +
    `${same.who ? ' (' + same.who + ')' : ''}.\nВсё равно записать?`);
}

function openAccountSheet() {
  const list = $('acc-list');
  list.innerHTML = '';
  const accounts = (state.snap && state.snap.accounts) || [];
  const bal = currentBalances();
  const cur = (state.editing && state.editing.account) || LS.get('lastAccount', '') ||
    (state.who === 'Арина' ? 'Kaspi Арина' : 'Kaspi Егор');
  $('acc-sheet-title').textContent = state.type === 'Доход' ? 'На какой счёт?' : 'С какого счёта?';
  for (const a of accounts) {
    const li = document.createElement('li');
    li.className = 'acc-pick' + (a === cur ? ' on' : '');
    li.innerHTML = '<span></span><small></small>';
    li.querySelector('span').textContent = a;
    li.querySelector('small').textContent = a in bal ? money(bal[a]) : '';
    li.onclick = () => { state.afterSheet = () => save(a); history.back(); };
    list.appendChild(li);
  }
  $('acc-sheet').hidden = false;
  history.pushState(Object.assign({}, currentNav(), { sheet: true }), '');
}

function closeSheet() {
  if ($('cat-sheet').hidden) return;
  if (currentNav().sheet) history.back();   // popstate сам закроет лист
  else $('cat-sheet').hidden = true;
}

// ---------- сводка ----------
function renderStamp(offline) {
  const s = $('stamp');
  if (!state.snapAt) { s.textContent = ''; return; }
  const d = new Date(state.snapAt);
  const hm = d.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
  const day = d.toDateString() === new Date().toDateString() ? '' : d.toLocaleDateString('ru-RU', { day: '2-digit', month: '2-digit' }) + ' ';
  s.textContent = (offline || !navigator.onLine ? 'нет сети · данные на ' : 'обновлено ') + day + hm;
  s.classList.toggle('offline', !!offline || !navigator.onLine);
}

/** Неотправленные траты сразу уменьшают «остаток на сегодня», чтобы цифра не врала оффлайн. */
function pendingSpendToday() {
  const t = todayYmd();
  const start = state.snap && state.snap.summary && state.snap.summary.periodStart;
  if (start && t < start) return 0;   // сегодня ещё не в периоде — на остаток не влияет
  return state.queue
    .filter((q) => q.op === 'add' && q.tx.status === 'Факт' && q.tx.type === 'Расход' && q.tx.date === t)
    .reduce((s, q) => s + q.tx.amount, 0);
}

function renderSummary() {
  const sm = (state.snap && state.snap.summary) || {};
  const left = sm.todayLeft != null ? sm.todayLeft - pendingSpendToday() : null;
  const set = (id, v) => { $(id).textContent = v == null ? '—' : money(v); };
  set('s-today', left);
  set('today-left', left);
  $('s-today').classList.toggle('neg', left < 0);
  $('today-left').classList.toggle('neg', left < 0);
  $('s-limit').textContent = sm.dailyLimit != null ? `лимит на день ${money(sm.dailyLimit)}` : '';
  set('s-free', sm.freeNow);
  const bal = (state.snap && state.snap.balances) || [];
  const total = bal.length
    ? Object.values(currentBalances()).reduce((a, v) => a + v, 0)
    : (sm.balanceNow != null ? sm.balanceNow + pendingBalanceDelta().total : null);
  set('s-balance', total);
  renderBalances();
  $('s-days').textContent = sm.daysLeft != null ? sm.daysLeft : '—';
  renderStamp();
  renderForecast();
  renderEnvelopes();
  renderDays();
}

/** Неотправленные факты из очереди: на сколько они меняют каждый счёт и общий остаток. */
function pendingBalanceDelta() {
  const by = {};
  let total = 0;
  const add = (acc, v) => { by[acc || 'Без счёта'] = (by[acc || 'Без счёта'] || 0) + v; };
  state.queue.filter((q) => q.op === 'add' && q.tx.status === 'Факт').forEach(({ tx }) => {
    if (tx.type === 'Расход') { add(tx.account, -tx.amount); total -= tx.amount; }
    else if (tx.type === 'Доход') { add(tx.account, tx.amount); total += tx.amount; }
    else if (tx.type === TRANSFER) { add(tx.account, -tx.amount); add(tx.toAccount, tx.amount); }
  });
  return { by: by, total: total };
}

/**
 * Остатки по счетам с учётом очереди: сверка задаёт сумму, факты после неё её двигают.
 * Порядок важен — проходим очередь как есть.
 */
function currentBalances() {
  const cur = {};
  ((state.snap && state.snap.balances) || []).forEach((b) => { cur[b.name] = b.now; });
  const add = (acc, v) => { if (acc in cur) cur[acc] += v; };
  for (const q of state.queue) {
    if (q.op === 'setBalance') { cur[q.account] = q.value; continue; }
    if (q.op !== 'add' || q.tx.status !== 'Факт') continue;
    const tx = q.tx;
    if (tx.type === 'Расход') add(tx.account, -tx.amount);
    else if (tx.type === 'Доход') add(tx.account, tx.amount);
    else if (tx.type === TRANSFER) { add(tx.account, -tx.amount); add(tx.toAccount, tx.amount); }
  }
  return cur;
}

/** Раскрывающийся список остатков по счетам (лист «Счета») и история сверок. */
function renderBalances() {
  const open = !!LS.get('showBalances', false);
  $('accounts-panel').hidden = !open;
  $('balance-toggle').setAttribute('aria-expanded', String(open));
  if (!open) return;
  const ul = $('balances');
  ul.innerHTML = '';
  const list = (state.snap && state.snap.balances) || [];
  if (!list.length) emptyRow(ul, 'Нет данных по счетам — обнови Apps Script и запусти migrateAccountsV3');
  const cur = currentBalances();
  const pendingSet = new Set(state.queue.filter((q) => q.op === 'setBalance').map((q) => q.account));
  for (const b of list) {
    const now = cur[b.name];
    const li = document.createElement('li');
    li.className = 'tap acc' + (now < 0 ? ' neg' : '') + (!now ? ' dim' : '') + (pendingSet.has(b.name) ? ' pending' : '');
    li.innerHTML = '<span class="what"><span></span><small></small></span><span class="sum"></span>';
    li.querySelector('.what span').textContent = b.name;
    li.querySelector('.sum').textContent = money(now);
    li.onclick = () => editBalance(b.name, now);
    ul.appendChild(li);
  }
  renderBalanceHistory();
}

function editBalance(name, now) {
  const raw = prompt(`Сколько сейчас на счёте «${name}»?`, String(Math.round(now)));
  if (raw == null) return;
  const value = Number(String(raw).replace(/[\s\u00a0]/g, '').replace(',', '.'));
  if (!isFinite(value) || String(raw).trim() === '') return toast('Нужно число');
  if (Math.round(value) === Math.round(now)) return toast('Остаток не изменился');
  enqueue({ op: 'setBalance', account: name, value: value, who: state.who, at: Date.now() });
  toast(`${name}: ${money(now)} → ${money(value)}`);
}

function renderBalanceHistory() {
  const ul = $('balance-history');
  ul.innerHTML = '';
  const pending = state.queue.filter((q) => q.op === 'setBalance').reverse().map((q) => {
    const d = new Date(q.at || Date.now());
    const at = `${todayYmd()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
    return { at: at, account: q.account, before: null, after: q.value, who: q.who, pending: true };
  });
  const list = pending.concat((state.snap && state.snap.balanceHistory) || []);
  if (!list.length) return emptyRow(ul, 'Изменений пока не было');
  for (const h of list) {
    const li = document.createElement('li');
    li.className = 'hist' + (h.pending ? ' pending' : '');
    li.innerHTML = '<span class="when"></span><span class="what"><span></span><small></small></span><span class="sum"></span>';
    const [date, time] = h.at.split(' ');
    li.querySelector('.when').textContent = `${shortDate(date)}\n${time || ''}`;
    li.querySelector('.what span').textContent = h.account;
    li.querySelector('.what small').textContent =
      (h.before != null ? `${money(h.before)} → ${money(h.after)}` : `→ ${money(h.after)}`) + (h.who ? ` · ${h.who}` : '');
    const diff = h.before != null ? h.after - h.before : null;
    const sum = li.querySelector('.sum');
    sum.textContent = diff == null ? '' : (diff > 0 ? '+' : '') + money(diff);
    sum.classList.toggle('income', diff > 0);
    sum.classList.toggle('neg', diff < 0);
    ul.appendChild(li);
  }
}

function openTransfer() {
  resetEntry();
  state.type = TRANSFER;
  state.fromPlan = true;          // открыт из сводки: «назад» и сохранение возвращают туда
  go('entry', true);
  renderEntry();
}

const WEEKDAYS = ['вс', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'];
function dayTitle(ymd) {
  const [y, m, d] = ymd.split('-').map(Number);
  const wd = WEEKDAYS[new Date(y, m - 1, d).getDay()];
  const t = new Date(); const today = todayYmd();
  t.setDate(t.getDate() + 1);
  const tomorrow = `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`;
  const name = ymd === today ? 'Сегодня' : ymd === tomorrow ? 'Завтра' : '';
  return (name ? name + ' · ' : '') + `${wd} ${shortDate(ymd)}`;
}

/**
 * Планы и факты с учётом неотправленных правок из очереди.
 * Нет feed (старая версия Apps Script) — собираем из expected/recent.
 */
function feedData() {
  const snap = state.snap || {};
  const f = snap.feed || {
    plans: (snap.expected || []).map((x) => Object.assign({}, x, { amount: x.left })),
    facts: snap.recent || [],
  };
  const deleting = new Set(state.queue.filter((q) => q.op === 'delete').map((q) => q.id));
  const updates = {};
  state.queue.filter((q) => q.op === 'update').forEach((q) => { updates[q.id] = Object.assign(updates[q.id] || {}, q.fields); });
  const plans = [];
  const facts = [];
  const put = (t) => ((t.status || 'Факт') === 'План' ? plans : facts).push(t);
  f.plans.forEach((t) => { if (!deleting.has(t.id)) put(Object.assign({ status: 'План' }, t, updates[t.id] || {}, updates[t.id] ? { pending: true } : {})); });
  f.facts.forEach((t) => { if (!deleting.has(t.id)) put(Object.assign({ status: 'Факт' }, t, updates[t.id] || {}, updates[t.id] ? { pending: true } : {})); });
  state.queue.filter((q) => q.op === 'add').forEach((q) => put(Object.assign({ pending: true }, q.tx)));
  return { plans: plans, facts: facts };
}

/** Прогноз до конца периода — считается на «Своде», здесь только показ. */
function renderForecast() {
  const sm = (state.snap && state.snap.summary) || {};
  const box = $('forecast');
  if (sm.lifeMoney == null || sm.lifeMoney === '') { box.hidden = true; return; }
  box.hidden = false;
  const gap = Number(sm.forecastGap) || 0;
  const ok = gap >= 0;
  const runOut = sm.runOutDate ? shortDate(sm.runOutDate) : '';
  box.innerHTML = '';
  const add = (cls, html) => { const d = document.createElement('div'); d.className = cls; d.innerHTML = html; box.appendChild(d); };
  // Период кончается в день зарплаты — она придёт в окне nextIncomeFrom…periodEnd; считаем по худшему случаю
  const until = sm.nextIncomeFrom && sm.periodEnd
    ? `до зарплаты ${shortDate(sm.nextIncomeFrom)}–${shortDate(sm.periodEnd)}`
    : `до ${sm.payday ? shortDate(sm.payday) : 'конца периода'}`;
  add('f-title', `Еда, такси, транспорт · ${until} · ${sm.daysLeft} дн. · есть ${money(sm.lifeMoney)}`);
  add('f-row', `<span>Можно тратить в день</span><b>${money(sm.dailyBudget)}</b>`);
  add('f-row', `<span>Тратите сейчас (за 7 дней)</span><b>${money(sm.pace7)}</b>`);
  add('verdict ' + (ok ? 'ok' : 'bad'), ok
    ? `Хватит до зарплаты, запас ≈ ${money(gap)}`
    : `Не хватит ≈ ${money(-gap)}${runOut ? ' · деньги закончатся ~' + runOut : ''}`);
  // Ожидаемые (не гарантированные) доходы — отдельной строкой «а если придут»
  const exp = Number(sm.expectedLeft) || 0;
  if (exp > 0 && sm.dailyBudgetMax != null) {
    const gm = Number(sm.forecastGapMax) || 0;
    add('f-alt', `Если придёт ожидаемое (+${money(exp)}): ${money(sm.dailyBudgetMax)} в день · ` +
      (gm >= 0 ? `запас ≈ ${money(gm)}` : `не хватит ≈ ${money(-gm)}`));
  }
  // Гарантированный доход, который должен был прийти, но не пришёл
  const guaranteed = new Set(((state.snap && state.snap.categoryTable) || [])
    .filter((c) => c.type === 'Доход' && c.guaranteed).map((c) => c.name));
  ((state.snap && state.snap.expected) || [])
    .filter((x) => x.type === 'Доход' && x.overdue && guaranteed.has(x.category))
    .forEach((x) => add('f-late', `⚠ Ждём «${x.category}» с ${shortDate(x.date)} — ${money(x.left)}`));
}

/** Конверты: план и факт по статьям расходов за период («Свод»). */
function renderEnvelopes() {
  const open = !!LS.get('showEnvelopes', false);
  $('env-toggle').setAttribute('aria-expanded', String(open));
  $('envelopes').hidden = !open;
  if (!open) return;
  const ul = $('envelopes');
  ul.innerHTML = '';
  const days = Number((state.snap && state.snap.summary && state.snap.summary.daysLeft) || 1);
  const rows = ((state.snap && state.snap.categoryTable) || [])
    .filter((c) => c.type === 'Расход' && (c.plan > 0 || c.fact > 0))
    .map((c) => Object.assign({ ratio: c.plan ? c.fact / c.plan : Infinity }, c))
    .sort((a, b) => (a.mandatory - b.mandatory) || (b.ratio - a.ratio) || (b.fact - a.fact));
  if (!rows.length) return emptyRow(ul, 'Нет данных — обнови Apps Script и запусти migrateV4');
  for (const c of rows) {
    const li = document.createElement('li');
    const unplanned = !c.plan;
    // Отложить больше плана — не перерасход
    const over = c.plan && c.fact > c.plan && c.name !== 'Сбережения';
    li.className = 'env' + (c.mandatory ? ' mandatory' : '') + (unplanned ? ' unplanned' : over ? ' over' : c.ratio > 0.85 ? ' warn' : '');
    const left = c.plan - c.fact;
    const note = unplanned ? `вне плана ${money(c.fact)}`
      : over ? `перерасход ${money(-left)}`
        : left < 0 ? `сверх плана ${money(-left)}`
        : `осталось ${money(left)}${c.daily ? ' · ≈' + money(left / days) + '/день' : ''}`;
    li.innerHTML = '<div class="env-top"><span></span><small></small></div><div class="bar"><i></i></div>';
    li.querySelector('span').textContent = c.name + (c.plan ? ` · ${money(c.fact)} из ${money(c.plan)}` : '');
    li.querySelector('small').textContent = note;
    li.querySelector('.bar i').style.width = (unplanned ? 100 : Math.min(100, c.ratio * 100)) + '%';
    ul.appendChild(li);
  }
}

/**
 * Лента по дням: «Сегодня» всегда сверху, под ним переключатель «Прошедшие · Предстоящие».
 * В дне: планы, разделитель, факты, остаток плана (план расходов − факт расходов за день).
 */
function renderDays() {
  const box = $('days');
  box.innerHTML = '';
  const { plans, facts } = feedData();
  const today = todayYmd();
  const days = {};
  const day = (d) => (days[d] = days[d] || { plans: [], facts: [] });
  plans.forEach((t) => day(t.date).plans.push(t));
  facts.forEach((t) => day(t.date).facts.push(t));
  day(today);

  markPlanCoverage(plans);
  $('today-card').innerHTML = '';
  $('today-card').appendChild(dayCard(today, days[today]));

  const tab = LS.get('feedTab', 'past');
  for (const b of $('feed-seg').children) b.classList.toggle('on', b.dataset.v === tab);
  const dates = Object.keys(days).filter((d) => d !== today);
  const list = tab === 'past'
    ? dates.filter((d) => d < today).sort().reverse()
    : dates.filter((d) => d > today).sort();
  if (!list.length) {
    const ul = document.createElement('ul');
    ul.className = 'list';
    emptyRow(ul, tab === 'past' ? 'Прошедших дней пока нет' : 'Впереди ничего не запланировано');
    box.appendChild(ul);
  }
  list.forEach((d) => box.appendChild(dayCard(d, days[d])));
}

// ---------- отчёт ----------
function openReport() {
  go('report');
  renderReport();
}

/** Отчёт за период: итоги, план/факт по статьям, вне плана, недели, крупные траты, кто сколько. */
function renderReport() {
  const box = $('report-body');
  box.innerHTML = '';
  const snap = state.snap || {};
  const sm = snap.summary || {};
  const table = snap.categoryTable || [];
  if (!table.length) {
    box.innerHTML = '<p class="hint">Нет данных — обнови Apps Script и запусти migrateV4</p>';
    return;
  }
  const mand = {};
  table.forEach((c) => { mand[c.name + '|' + c.type] = c.mandatory; });
  const start = sm.periodStart || todayYmd();
  const today = todayYmd();
  const facts = feedData().facts.filter((t) => t.date >= start && t.date <= today);
  const sum = (arr, f) => arr.reduce((s, x) => s + (f ? f(x) : x), 0);
  const h = (text) => { const e = document.createElement('h2'); e.textContent = text; box.appendChild(e); };
  const tbl = (head, rows) => {
    const t = document.createElement('table');
    t.className = 'rep-table';
    const esc = (v) => String(v).replace(/[&<>]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[ch]));
    t.innerHTML = (head ? '<tr>' + head.map((x) => `<th>${esc(x)}</th>`).join('') + '</tr>' : '') +
      rows.map((r) => `<tr class="${r.cls || ''}">` +
        r.cells.map((x, i) => `<td class="${(r.tdCls && r.tdCls[i]) || ''}">${esc(x)}</td>`).join('') + '</tr>').join('');
    box.appendChild(t);
  };

  // Шапка и итоги
  const passed = Math.round((Date.parse(today) - Date.parse(start)) / 864e5) + 1;
  const total = passed + Number(sm.daysLeft || 1) - 1;
  const p = document.createElement('p');
  p.className = 'hint';
  p.textContent = `Период ${shortDate(start)} – ${sm.periodEnd ? shortDate(sm.periodEnd) : ''} · прошло ${passed} из ${total} дней`;
  box.appendChild(p);

  const inc = table.filter((c) => c.type === 'Доход');
  const exp = table.filter((c) => c.type === 'Расход');
  const savings = sum(exp.filter((c) => c.name === 'Сбережения'), (c) => c.fact);
  const mandatory = sum(exp.filter((c) => c.mandatory && c.name !== 'Сбережения'), (c) => c.fact);
  const variable = sum(exp.filter((c) => !c.mandatory), (c) => c.fact);
  h('Итоги');
  tbl(null, [
    { cells: ['Доходы', money(sum(inc, (c) => c.fact))], cls: 'total' },
    { cells: ['Расходы', money(sum(exp, (c) => c.fact))], cls: 'total' },
    { cells: ['обязательные', money(mandatory)], cls: 'sub' },
    { cells: ['в накопления', money(savings)], cls: 'sub' },
    { cells: ['остальные', money(variable)], cls: 'sub' },
    { cells: ['На счетах сейчас', money(sm.balanceNow)], cls: 'total' },
  ]);

  // Прогноз
  if (sm.lifeMoney != null && sm.lifeMoney !== '') {
    const gap = Number(sm.forecastGap) || 0;
    h('Прогноз до конца периода');
    tbl(null, [
      { cells: ['Деньги на ежедневное', money(sm.lifeMoney)] },
      { cells: ['Можно в день (еда, такси, транспорт)', money(sm.dailyBudget)] },
      { cells: ['Тратите сейчас (7 дней)', money(sm.pace7)] },
      { cells: [gap >= 0 ? 'Запас к концу периода' : 'Нехватка к концу периода', money(Math.abs(gap))],
        tdCls: [null, gap >= 0 ? 'ok' : 'over'], cls: 'total' },
    ]);
  }

  // План и факт по статьям
  h('Расходы: план и факт');
  const rows = exp.filter((c) => c.plan || c.fact).sort((a, b) => b.fact - a.fact).map((c) => {
    const pct = c.plan ? Math.round(c.fact / c.plan * 100) + '%' : 'вне плана';
    const bad = !c.plan || (c.fact > c.plan && c.name !== 'Сбережения');
    return { cells: [c.name, c.plan ? money(c.plan) : '—', money(c.fact), pct], tdCls: [null, null, null, bad ? 'over' : ''] };
  });
  rows.push({ cells: ['Итого', money(sum(exp, (c) => c.plan)), money(sum(exp, (c) => c.fact)), ''], cls: 'total' });
  tbl(['Статья', 'План', 'Факт', '%'], rows);
  const unplanned = sum(exp.filter((c) => !c.plan), (c) => c.fact);
  if (unplanned) {
    const q = document.createElement('p');
    q.className = 'hint';
    q.textContent = `Траты вне плана: ${money(unplanned)}`;
    box.appendChild(q);
  }

  h('Доходы: план и факт');
  tbl(['Источник', 'План', 'Факт', ''], inc.filter((c) => c.plan || c.fact).sort((a, b) => b.fact - a.fact)
    .map((c) => ({ cells: [c.name, c.plan ? money(c.plan) : '—', money(c.fact), c.guaranteed ? 'гарант.' : 'ожид.'] })));

  // Динамика по неделям: необязательные расходы
  const varFacts = facts.filter((t) => t.type === 'Расход' && !mand[t.category + '|' + t.type]);
  const weeks = {};
  varFacts.forEach((t) => {
    const w = Math.floor((Date.parse(t.date) - Date.parse(start)) / (7 * 864e5));
    weeks[w] = (weeks[w] || 0) + t.amount;
  });
  const wrows = Object.keys(weeks).map(Number).sort((a, b) => a - b).map((w) => {
    const from = new Date(Date.parse(start) + w * 7 * 864e5).toISOString().slice(0, 10);
    const toMs = Math.min(Date.parse(from) + 6 * 864e5, Date.parse(today));
    const to = new Date(toMs).toISOString().slice(0, 10);
    const n = Math.round((toMs - Date.parse(from)) / 864e5) + 1;
    return { cells: [`${shortDate(from)}–${shortDate(to)}`, money(weeks[w]), money(weeks[w] / n)] };
  });
  if (wrows.length) {
    h('Динамика: необязательные траты');
    tbl(['Неделя', 'Сумма', 'В день'], wrows);
  }

  // Крупные разовые траты
  const big = varFacts.filter((t) => t.amount >= 10000).sort((a, b) => b.amount - a.amount);
  if (big.length) {
    h('Крупные траты (от 10 000)');
    tbl(null, big.map((t) => ({ cells: [`${shortDate(t.date)} · ${t.category}${t.comment ? ' · ' + t.comment : ''}`, money(t.amount)] })));
  }

  // Кто сколько потратил
  const who = {};
  facts.filter((t) => t.type === 'Расход').forEach((t) => {
    const k = t.who || '—';
    who[k] = who[k] || { all: 0, v: 0 };
    who[k].all += t.amount;
    if (!mand[t.category + '|' + t.type]) who[k].v += t.amount;
  });
  h('Кто сколько потратил');
  tbl(['', 'Всего', 'Необязательные'], Object.keys(who).map((k) => ({ cells: [k, money(who[k].all), money(who[k].v)] })));
}

/**
 * Какая часть каждого плана уже покрыта фактом. Факт статьи за период («Свод», уже с учётом сдвига
 * для зарплаты и обязательных) гасит её планы по порядку дат — как в «Календаре план».
 * Оплатил рассрочки 06-го при плане на 10-е — план 10-го исчезает из ленты.
 */
function markPlanCoverage(plans) {
  // Ежедневные статьи (еда, такси, транспорт) — день за днём: план дня гасят только траты этого дня
  const daily = new Set(((state.snap && state.snap.categoryTable) || []).filter((c) => c.daily).map((c) => c.name + '|' + c.type));
  const dayFact = {};
  feedData().facts.forEach((t) => {
    const k = t.category + '|' + t.type + '|' + t.date;
    dayFact[k] = (dayFact[k] || 0) + t.amount;
  });
  plans.filter((p) => daily.has(p.category + '|' + p.type)).forEach((p) => {
    p.left = Math.max(0, p.amount - (dayFact[p.category + '|' + p.type + '|' + p.date] || 0));
  });

  const fact = {};
  ((state.snap && state.snap.categoryTable) || []).forEach((c) => { fact[c.name + '|' + c.type] = c.fact; });
  state.queue.filter((q) => q.op === 'add' && q.tx.status === 'Факт').forEach(({ tx }) => {
    const k = tx.category + '|' + tx.type;
    if (k in fact) fact[k] += tx.amount;
  });
  plans.slice().sort((a, b) => a.date.localeCompare(b.date)).forEach((p) => {
    const k = p.category + '|' + p.type;
    if (daily.has(k)) return;
    const covered = Math.min(p.amount, Math.max(0, fact[k] || 0));
    fact[k] = (fact[k] || 0) - covered;
    p.left = p.amount - covered;
  });
}

function dayCard(date, g) {
  const ul = document.createElement('ul');
  ul.className = 'list day-card';
  const head = document.createElement('li');
  head.className = 'day' + (date === todayYmd() ? ' is-today' : '');
  head.innerHTML = '<span class="what"></span>';
  head.querySelector('.what').textContent = dayTitle(date);
  ul.appendChild(head);

  const row = (t, isFact) => {
    const li = document.createElement('li');
    const isT = t.type === TRANSFER;
    li.className = 'tap item' + (isFact ? ' fact' : ' plan') + (isT ? ' transfer' : '') + (t.pending ? ' pending' : '');
    li.innerHTML = '<span class="what"><span></span><small></small></span><span class="sum"></span>';
    li.querySelector('.what span').textContent = isT ? `${t.account} → ${t.toAccount}` : t.category;
    const partial = !isFact && t.left != null && t.left < t.amount;
    li.querySelector('.what small').textContent = isFact ? (t.comment || '') : partial ? `из ${money(t.amount)}` : '';
    const sum = li.querySelector('.sum');
    sum.textContent = (t.type === 'Доход' ? '+' : '') + money(isFact || t.left == null ? t.amount : t.left);
    sum.classList.toggle('income', t.type === 'Доход');
    if (isFact) {
      const del = document.createElement('button');
      del.className = 'del';
      del.type = 'button';
      del.setAttribute('aria-label', 'Удалить');
      del.textContent = '✕';
      del.onclick = () => removeTx(t);
      li.appendChild(del);
      li.onclick = (e) => { if (!e.target.closest('.del')) startEdit(t); };
    } else {
      li.onclick = () => prefillFromPlan(t, g.facts);
    }
    ul.appendChild(li);
  };

  const open = g.plans.filter((t) => t.left == null || t.left > 0);   // оплаченные планы не показываем
  open.forEach((t) => row(t, false));
  if (open.length && g.facts.length) {
    const sep = document.createElement('li');
    sep.className = 'sep';
    ul.appendChild(sep);
  }
  g.facts.forEach((t) => row(t, true));
  if (!open.length && !g.facts.length) emptyRow(ul, g.plans.length ? 'План оплачен' : 'Ничего не запланировано');

  // Остаток плана дня — неоплаченная часть планов расходов (факты уже зачтены в покрытие)
  const openOut = open.filter((t) => t.type === 'Расход');
  if (openOut.length) {
    const left = openOut.reduce((s, t) => s + (t.left == null ? t.amount : t.left), 0);
    const foot = document.createElement('li');
    foot.className = 'total' + (left < 0 ? ' neg' : '');
    foot.innerHTML = '<span class="what">Остаток плана</span><span class="sum"></span>';
    foot.querySelector('.sum').textContent = money(left);
    ul.appendChild(foot);
  }
  return ul;
}

function prefillFromPlan(p, dayFacts) {
  // Предлагаем то, что по этой категории в этот день ещё не потрачено.
  const rest = p.left != null ? p.left
    : p.amount - (dayFacts || []).filter((f) => f.category === p.category).reduce((s, f) => s + f.amount, 0);
  state.editing = null;
  state.type = p.type;
  state.category = p.category;
  state.amount = String(rest > 0 ? rest : p.amount);
  state.plan = false;
  state.fromPlan = true;
  $('date').value = todayYmd();
  go('entry', true);
  renderEntry();
}

function removeTx(t) {
  const label = t.type === TRANSFER ? `перевод ${t.account} → ${t.toAccount}` : t.category;
  if (!confirm(`Удалить ${money(t.amount)} · ${label}?`)) return;
  const i = state.queue.findIndex((q) => q.op === 'add' && q.tx.id === t.id);
  if (i >= 0) {                      // ещё не ушло на сервер — просто убрать из очереди
    state.queue.splice(i, 1);
    LS.set('queue', state.queue);
    renderAll();
    return;
  }
  enqueue({ op: 'delete', id: t.id });
}

function emptyRow(ul, text) {
  const li = document.createElement('li');
  li.className = 'empty';
  li.textContent = text;
  ul.appendChild(li);
}

function renderAll() {
  renderSummary();
  renderAccounts();
  renderEntry();
}

// ---------- настройки ----------
function openSetup(first) {
  $('setup-url').value = state.url;
  for (const b of $('setup-who').children) b.classList.toggle('on', b.dataset.v === state.who);
  $('setup-cancel').hidden = first;
  $('setup-err').hidden = true;
  if (first) show('setup'); else go('setup');
}

async function saveSetup() {
  const url = $('setup-url').value.trim();
  const whoBtn = [...$('setup-who').children].find((b) => b.classList.contains('on'));
  const err = (m) => { $('setup-err').textContent = m; $('setup-err').hidden = false; };
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(url)) return err('Нужен адрес вида https://script.google.com/macros/s/…/exec');
  if (!whoBtn) return err('Выбери, кто пользуется телефоном');
  $('setup-save').disabled = true;
  $('setup-save').textContent = 'Проверяю…';
  const prev = state.url;
  state.url = url;
  try {
    const j = await apiGet();
    LS.set('url', url);
    state.who = whoBtn.dataset.v;
    LS.set('who', state.who);
    applySnap(j);
    if (currentNav().screen === 'setup') history.back(); else show('entry');
    flush();
  } catch (e) {
    state.url = prev;
    err('Не удалось подключиться: ' + e.message);
  } finally {
    $('setup-save').disabled = false;
    $('setup-save').textContent = 'Готово';
  }
}

// ---------- события ----------
function bind() {
  $('pad').addEventListener('click', (e) => { const k = e.target.dataset.k; if (k) press(k); });
  document.addEventListener('keydown', (e) => {
    if ($('entry').hidden || e.target.tagName === 'INPUT') return;
    if (/^\d$/.test(e.key)) press(e.key);
    else if (e.key === 'Backspace') press('back');
    else if (e.key === 'Enter' && !$('save').disabled && $('acc-sheet').hidden) save();
  });
  $('type-seg').addEventListener('click', (e) => {
    const v = e.target.dataset.v;
    if (!v || v === state.type) return;
    state.type = v;
    state.category = '';
    if (v === TRANSFER) state.plan = false;
    renderEntry();
  });
  $('from-account').addEventListener('change', renderEntry);
  $('to-account').addEventListener('change', renderEntry);
  $('balance-toggle').onclick = () => { LS.set('showBalances', !LS.get('showBalances', false)); renderBalances(); };
  $('transfer-btn').onclick = openTransfer;
  $('env-toggle').onclick = () => { LS.set('showEnvelopes', !LS.get('showEnvelopes', false)); renderEnvelopes(); };
  $('feed-seg').addEventListener('click', (e) => {
    const v = e.target.dataset.v;
    if (!v) return;
    LS.set('feedTab', v);
    renderDays();
  });
  $('report-btn').onclick = openReport;
  $('report-back').onclick = () => history.back();
  $('plan-toggle').onclick = () => { state.plan = !state.plan; renderEntry(); };
  $('date').addEventListener('change', renderEntry);
  $('comment').addEventListener('keydown', (e) => { if (e.key === 'Enter') e.target.blur(); });
  $('save').onclick = () => save();
  $('edit-cancel').onclick = () => history.back();
  $('today-chip').onclick = () => {
    if (currentNav().sub) history.back(); else { go('summary'); renderSummary(); }
    refresh();
  };
  $('back').onclick = () => history.back();
  $('settings').onclick = () => openSetup(false);
  $('setup-who').addEventListener('click', (e) => {
    if (!e.target.dataset.v) return;
    for (const b of $('setup-who').children) b.classList.toggle('on', b === e.target);
  });
  $('setup-save').onclick = saveSetup;
  $('setup-cancel').onclick = () => history.back();
  $('cat-sheet').addEventListener('click', (e) => { if (e.target.id === 'cat-sheet' || e.target.dataset.close != null) closeSheet(); });
  $('acc-sheet').addEventListener('click', (e) => {
    if ((e.target.id === 'acc-sheet' || e.target.dataset.close != null) && currentNav().sheet) history.back();
  });

  window.addEventListener('online', () => { flush(); refresh(); });
  window.addEventListener('offline', () => renderStamp(true));
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') { flush(); refresh(); }
  });
}

// ---------- старт ----------
bind();
history.replaceState({ screen: 'entry' }, '');
$('date').value = todayYmd();
if (!state.url || !state.who) {
  openSetup(true);
} else {
  show('entry');
  renderAll();                        // мгновенно из кеша
  flush().then(refresh);              // свежие цифры — в фоне
}

if ('serviceWorker' in navigator) {
  // Новая версия приложения встала — перезагрузиться один раз, чтобы не ждать следующего запуска.
  const hadController = !!navigator.serviceWorker.controller;
  let reloaded = false;
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (!hadController || reloaded) return;
    reloaded = true;
    location.reload();
  });
  window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
}

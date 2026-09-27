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
  showAllRecent: false,
};

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

function renderAccounts() {
  const sel = $('account');
  const accounts = (state.snap && state.snap.accounts) || [];
  const cur = sel.value || LS.get('lastAccount', '') || (state.who === 'Арина' ? 'Kaspi Арина' : 'Kaspi Егор');
  sel.innerHTML = '';
  for (const a of accounts) {
    const o = document.createElement('option');
    o.value = o.textContent = a;
    sel.appendChild(o);
  }
  if (accounts.includes(cur)) sel.value = cur;
}

function renderEntry() {
  const v = state.amount ? money(state.amount) : '0';
  $('amount-value').textContent = v;
  $('amount').classList.toggle('empty', !state.amount);
  $('amount').classList.toggle('income', state.type === 'Доход');
  for (const b of $('type-seg').children) b.classList.toggle('on', b.dataset.v === state.type);
  const d = $('date').value || todayYmd();
  $('date-label').textContent = dateLabel(d);
  $('date').parentElement.classList.toggle('changed', d !== todayYmd());
  $('plan-toggle').setAttribute('aria-pressed', String(state.plan));
  $('save').disabled = !(Number(state.amount) > 0 && state.category);
  $('save').textContent = state.editing ? 'Сохранить изменения' : state.plan ? 'Записать в план' : 'Записать';
  $('edit-bar').hidden = !state.editing;
  renderChips();
}

function press(k) {
  if (k === 'back') state.amount = state.amount.slice(0, -1);
  else if ((state.amount + k).replace(/^0+/, '').length <= 9) state.amount = (state.amount + k).replace(/^0+/, '');
  if (navigator.vibrate) navigator.vibrate(8);
  renderEntry();
}

function save() {
  const amount = Number(state.amount);
  if (!(amount > 0) || !state.category) return;
  const fields = {
    date: $('date').value || todayYmd(),
    status: state.plan ? 'План' : 'Факт',
    type: state.type,
    amount: amount,
    category: state.category,
    account: $('account').value,
    comment: $('comment').value.trim(),
  };
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
    toast(`Изменено: ${money(amount)} · ${fields.category}`);
    state.editing = null;
    show('summary');
  } else {
    const tx = Object.assign({ id: uid(), who: state.who }, fields);
    LS.set('lastAccount', tx.account);
    enqueue({ op: 'add', tx: tx, at: Date.now() });
    toast(`${tx.status === 'План' ? 'В план' : 'Записано'}: ${money(amount)} · ${tx.category}`);
  }
  resetEntry();
}

function resetEntry() {
  state.editing = null;
  state.amount = '';
  state.category = '';
  state.plan = false;
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
  if (t.account && [...$('account').options].some((o) => o.value === t.account)) $('account').value = t.account;
  show('entry');
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
}
function closeSheet() { $('cat-sheet').hidden = true; }

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
  set('s-balance', sm.balanceNow);
  $('s-days').textContent = sm.daysLeft != null ? sm.daysLeft : '—';
  renderStamp();
  renderExpected();
  renderRecent();
}

/**
 * Ожидаемые: разовые платежи — отдельными строками; то, что повторяется
 * 4+ раза за окно (такси, транспорт), — одной строкой на категорию.
 */
function renderExpected() {
  const ul = $('expected');
  ul.innerHTML = '';
  const items = (state.snap && state.snap.expected) || [];
  const byCat = {};
  items.forEach((it) => { (byCat[it.category] = byCat[it.category] || []).push(it); });
  const rows = [];
  for (const cat in byCat) {
    const list = byCat[cat];
    if (list.length >= 4) {
      rows.push({
        date: list[0].date, category: cat, type: list[0].type, left: list.reduce((s, x) => s + x.left, 0),
        note: `${list.length} раз, до ${shortDate(list[list.length - 1].date)}`, overdue: list[0].overdue, group: true,
      });
    } else {
      list.forEach((x) => rows.push(Object.assign({}, x, { note: x.left < x.amount ? `осталось из ${money(x.amount)}` : '' })));
    }
  }
  rows.sort((a, b) => a.date.localeCompare(b.date));
  if (!rows.length) return emptyRow(ul, 'Ничего не ожидается');

  for (const r of rows) {
    const li = document.createElement('li');
    li.className = 'tap' + (r.overdue ? ' overdue' : '');
    li.innerHTML = `<span class="when"></span><span class="what"><span></span><small></small></span><span class="sum"></span>`;
    li.querySelector('.when').textContent = r.group ? 'с ' + shortDate(r.date) : shortDate(r.date);
    li.querySelector('.what span').textContent = r.category;
    li.querySelector('.what small').textContent = r.note || (r.overdue ? 'просрочено' : '');
    const sum = li.querySelector('.sum');
    sum.textContent = (r.type === 'Доход' ? '+' : '') + money(r.left);
    sum.classList.toggle('income', r.type === 'Доход');
    // Тап по разовой позиции — подставить её в ввод как оплаченную сегодня.
    if (!r.group) li.onclick = () => prefillFromPlan(r);
    ul.appendChild(li);
  }
}

function prefillFromPlan(p) {
  state.editing = null;
  state.type = p.type;
  state.category = p.category;
  state.amount = String(p.left);
  state.plan = false;
  $('date').value = todayYmd();
  show('entry');
  renderEntry();
}

function renderRecent() {
  const ul = $('recent');
  ul.innerHTML = '';
  const pending = state.queue.filter((q) => q.op === 'add' && q.tx.status === 'Факт').map((q) => Object.assign({ pending: true }, q.tx)).reverse();
  const deleting = new Set(state.queue.filter((q) => q.op === 'delete').map((q) => q.id));
  const updates = {};
  state.queue.filter((q) => q.op === 'update').forEach((q) => { updates[q.id] = Object.assign(updates[q.id] || {}, q.fields); });
  const server = ((state.snap && state.snap.recent) || [])
    .filter((t) => !deleting.has(t.id))
    .map((t) => (updates[t.id] ? Object.assign({}, t, updates[t.id], { pending: true }) : t));
  const all = pending.concat(server);
  const list = state.showAllRecent ? all : all.slice(0, 5);
  $('recent-more').hidden = all.length <= 5;
  $('recent-more').textContent = state.showAllRecent ? 'Свернуть' : `Показать ещё ${all.length - 5}`;
  if (!list.length) return emptyRow(ul, 'Пока пусто');

  for (const t of list) {
    const li = document.createElement('li');
    li.className = 'tap' + (t.pending ? ' pending' : '');
    li.onclick = (e) => { if (!e.target.closest('.del')) startEdit(t); };
    li.innerHTML = `<span class="when"></span><span class="what"><span></span><small></small></span><span class="sum"></span><button class="del" type="button" aria-label="Удалить">✕</button>`;
    li.querySelector('.when').textContent = shortDate(t.date);
    li.querySelector('.what span').textContent = t.category;
    li.querySelector('.what small').textContent = [t.comment, t.account, t.who].filter(Boolean).join(' · ');
    const sum = li.querySelector('.sum');
    sum.textContent = (t.type === 'Доход' ? '+' : '') + money(t.amount);
    sum.classList.toggle('income', t.type === 'Доход');
    li.querySelector('.del').onclick = () => removeTx(t);
    ul.appendChild(li);
  }
}

function removeTx(t) {
  if (!confirm(`Удалить ${money(t.amount)} · ${t.category}?`)) return;
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
  show('setup');
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
    show('entry');
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
    else if (e.key === 'Enter' && !$('save').disabled) save();
  });
  $('type-seg').addEventListener('click', (e) => {
    const v = e.target.dataset.v;
    if (!v || v === state.type) return;
    state.type = v;
    state.category = '';
    renderEntry();
  });
  $('plan-toggle').onclick = () => { state.plan = !state.plan; renderEntry(); };
  $('date').addEventListener('change', renderEntry);
  $('comment').addEventListener('keydown', (e) => { if (e.key === 'Enter') e.target.blur(); });
  $('save').onclick = save;
  $('edit-cancel').onclick = () => { resetEntry(); show('summary'); };
  $('recent-more').onclick = () => { state.showAllRecent = !state.showAllRecent; renderRecent(); };
  $('today-chip').onclick = () => { if (state.editing) resetEntry(); show('summary'); renderSummary(); refresh(); };
  $('back').onclick = () => show('entry');
  $('settings').onclick = () => openSetup(false);
  $('setup-who').addEventListener('click', (e) => {
    if (!e.target.dataset.v) return;
    for (const b of $('setup-who').children) b.classList.toggle('on', b === e.target);
  });
  $('setup-save').onclick = saveSetup;
  $('setup-cancel').onclick = () => show('entry');
  $('cat-sheet').addEventListener('click', (e) => { if (e.target.id === 'cat-sheet' || e.target.dataset.close != null) closeSheet(); });

  window.addEventListener('online', () => { flush(); refresh(); });
  window.addEventListener('offline', () => renderStamp(true));
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') { flush(); refresh(); }
  });
}

// ---------- старт ----------
bind();
$('date').value = todayYmd();
if (!state.url || !state.who) {
  openSetup(true);
} else {
  show('entry');
  renderAll();                        // мгновенно из кеша
  flush().then(refresh);              // свежие цифры — в фоне
}

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
}

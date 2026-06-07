function escHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

const MAX_OP_CARDS = 20;

function getActivityScrollEl() {
  return document.getElementById('activityPanelInner')
    || document.getElementById('activityPanelScroll')
    || document.getElementById('activityPanel');
}

function refreshActivityEmptyHint() {
  const emptyHint = document.getElementById('emptyHint');
  const panel = getActivityScrollEl();
  const pop = document.getElementById('helpPrescriptionsPopover');
  if (!emptyHint || !panel) return;
  const hasCards = panel.querySelectorAll('.op-card').length > 0;
  const helpOpen = pop && pop.classList.contains('is-open');
  emptyHint.style.display = (hasCards || helpOpen) ? 'none' : '';
}

function createOpCard(title, subtitle) {
  const panel = getActivityScrollEl();
  const statusId = 'ls_' + Date.now() + '_' + Math.random().toString(36).slice(2);
  const card = document.createElement('div');
  card.className = 'op-card summary-card is-running';
  card.innerHTML = `
    <div class="summary-card-header">
      <div class="summary-card-title-wrap">
        <div class="summary-card-title">${escHtml(title)}</div>
        ${subtitle ? `<div class="summary-card-subtitle">${escHtml(subtitle)}</div>` : ''}
      </div>
      <span class="summary-card-meta">${new Date().toLocaleTimeString('ru')}</span>
      <button type="button" class="op-card-collapse" aria-label="Свернуть">▼</button>
    </div>
    <div class="op-card-body">
      <div class="live-status op-live" id="${statusId}">
        <div class="live-status-row">
          <span class="spinner"></span>
          <span class="live-status-text"></span>
        </div>
      </div>
    </div>
  `;
  panel.insertBefore(card, document.getElementById('emptyHint'));
  card.querySelector('.op-card-collapse').addEventListener('click', () => {
    card.classList.toggle('is-collapsed');
  });
  refreshActivityEmptyHint();
  const cards = panel.querySelectorAll('.op-card');
  if (cards.length > MAX_OP_CARDS) cards[0].remove();
  return { card, statusId };
}

function setCardStatus(statusId, msg, state) {
  const el = document.getElementById(statusId);
  if (!el) return;
  const card = el.closest('.op-card');
  el.classList.remove('done', 'error');
  const spinner = el.querySelector('.spinner');
  const txt = el.querySelector('.live-status-text');
  if (state === 'done') {
    el.classList.add('done');
    if (spinner) spinner.style.display = 'none';
    if (card) { card.classList.remove('is-running'); card.classList.add('is-done'); }
  } else if (state === 'error') {
    el.classList.add('error');
    if (spinner) spinner.style.display = 'none';
    if (card) { card.classList.remove('is-running'); card.classList.add('is-error'); }
  }
  if (txt && msg != null) txt.textContent = msg;
}

function finalizeOpCard(card, statusId, stats, details, options) {
  options = options || {};
  const liveEl = document.getElementById(statusId);
  if (liveEl) liveEl.remove();
  card.classList.remove('is-running');
  card.classList.add(options.errorState ? 'is-error' : 'is-done');

  const body = card.querySelector('.op-card-body') || card;
  const statsContainer = document.createElement('div');
  statsContainer.className = 'summary-stats';
  statsContainer.innerHTML = (stats || []).map(s =>
    `<span class="stat-chip stat-${s.color}">${escHtml(s.label)}</span>`
  ).join('');
  body.appendChild(statsContainer);

  if (details && details.length) {
    const detailId = 'det_' + Date.now();
    const toggleBtn = document.createElement('button');
    toggleBtn.className = 'detail-toggle';
    toggleBtn.type = 'button';
    const label = options.detailLabel || 'Подробности';
    toggleBtn.textContent = options.expandDetails ? `${label} ▲` : `${label} ▼`;
    toggleBtn.onclick = () => {
      const block = document.getElementById(detailId);
      block.classList.toggle('open');
      toggleBtn.textContent = block.classList.contains('open') ? `${label} ▲` : `${label} ▼`;
    };
    body.appendChild(toggleBtn);

    const detailBlock = document.createElement('div');
    detailBlock.className = 'detail-block log-timeline' + (options.expandDetails ? ' open' : '');
    detailBlock.id = detailId;
    details.forEach(d => {
      const row = document.createElement('div');
      row.className = 'detail-row' + (d.wrap ? ' wrap' : '');
      row.innerHTML = `
        <span class="detail-icon">${d.icon || ''}</span>
        <span class="detail-name wrap">${escHtml(d.name)}</span>
        ${d.badge ? `<span class="detail-badge ${d.badgeClass || ''}">${escHtml(d.badge)}</span>` : ''}
      `;
      detailBlock.appendChild(row);
    });
    body.appendChild(detailBlock);
  }
}

async function apiFetch(url, options, statusId) {
  const res = await fetch(url, options);
  let data = null;
  const ct = (res.headers.get('content-type') || '').toLowerCase();
  if (ct.includes('application/json')) {
    data = await res.json();
  } else {
    const raw = await res.text();
    try { data = raw ? JSON.parse(raw) : {}; } catch (_) { data = { detail: raw }; }
  }
  if (!res.ok) {
    const detail = (data && (data.detail || data.message)) || res.statusText;
    const err = new Error(typeof detail === 'string' ? detail : JSON.stringify(detail));
    if (statusId) setCardStatus(statusId, 'Ошибка: ' + err.message, 'error');
    throw err;
  }
  return data;
}

function buildUploadDetails(filenames) {
  return (filenames || []).map(name => ({
    icon: '📊',
    name,
    wrap: true,
    badge: 'загружен',
    badgeClass: 'db-ok',
  }));
}

function addCheckedFile(filename, downloadUrl) {
  const container = document.getElementById('checkedFiles');
  const noRes = container.querySelector('.no-results');
  if (noRes) noRes.remove();
  const item = document.createElement('div');
  item.className = 'result-item';
  item.innerHTML = `<span title="${escHtml(filename)}">${escHtml(filename)}</span><a href="${downloadUrl}" download>Скачать</a>`;
  container.appendChild(item);
}

async function refreshCheckedList() {
  const data = await apiFetch('/prescriptions/results');
  const el = document.getElementById('checkedFiles');
  el.innerHTML = data.files && data.files.length
    ? data.files.map(f =>
        `<div class="result-item">
           <span title="${escHtml(f)}">${escHtml(f)}</span>
           <a href="/download/prescriptions/${encodeURIComponent(f)}" download>Скачать</a>
         </div>`
      ).join('')
    : '<div class="no-results">Пока нет</div>';
}

async function uploadPrescriptions(input) {
  const fd = new FormData();
  for (const f of input.files) fd.append('files', f);
  const { card, statusId } = createOpCard('Загрузка Excel');
  setCardStatus(statusId, `Загружаю ${input.files.length} файл(ов)…`);
  try {
    const data = await apiFetch('/upload/prescriptions', { method: 'POST', body: fd }, statusId);
    finalizeOpCard(card, statusId, [
      { label: `Загружено: ${data.count}`, color: 'green' },
    ], buildUploadDetails(data.uploaded), { expandDetails: true, detailLabel: 'Файлы' });
  } catch (_) {}
  input.value = '';
}

async function clearPrescriptions() {
  const { card, statusId } = createOpCard('Очистка загрузки');
  setCardStatus(statusId, 'Удаляю файлы…');
  try {
    await apiFetch('/clear/prescriptions/uploads', { method: 'DELETE' }, statusId);
    finalizeOpCard(card, statusId, [{ label: 'Загрузка очищена', color: 'blue' }], null);
  } catch (_) {}
}

async function resetPrescriptions() {
  const { card, statusId } = createOpCard('Сброс');
  setCardStatus(statusId, 'Очищаю загрузку и результаты…');
  try {
    await apiFetch('/clear/prescriptions/all', { method: 'DELETE' }, statusId);
    document.getElementById('checkedFiles').innerHTML = '<div class="no-results">Пока нет</div>';
    finalizeOpCard(card, statusId, [{ label: 'Сброшено', color: 'blue' }], null);
  } catch (_) {}
}

async function checkPrescriptions() {
  const bar = document.getElementById('progressBar');
  if (bar) bar.style.width = '15%';

  const { card, statusId } = createOpCard(
    'Проверка предписаний',
    'Структура таблицы и обязательные поля в строках'
  );
  setCardStatus(statusId, 'Запускаю проверку…');

  const details = [];
  let okCount = 0;
  let warnCount = 0;
  let errCount = 0;

  const resp = await fetch('/check/prescriptions/stream', { method: 'POST' });
  if (!resp.ok) {
    setCardStatus(statusId, 'Ошибка запуска', 'error');
    finalizeOpCard(card, statusId, [{ label: 'Ошибка', color: 'red' }], null, { errorState: true });
    return;
  }

  const reader = resp.body.getReader();
  const decoder = new TextDecoder();
  let buf = '';

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += decoder.decode(value, { stream: true });
    const lines = buf.split('\n');
    buf = lines.pop();
    for (const line of lines) {
      if (!line.startsWith('data: ')) continue;
      let ev;
      try { ev = JSON.parse(line.slice(6)); } catch (_) { continue; }

      if (ev.type === 'start') {
        setCardStatus(statusId, ev.msg || 'Проверяю…');
      } else if (ev.type === 'info') {
        setCardStatus(statusId, ev.msg);
      } else if (ev.type === 'report') {
        const hasErr = ev.hasErrors;
        if (hasErr) errCount++; else if (ev.hasWarnings) warnCount++; else okCount++;
        details.push({
          icon: hasErr ? '⚠' : (ev.hasWarnings ? '◐' : '✓'),
          name: ev.msg || ev.filename,
          wrap: true,
          badge: hasErr ? 'ошибки' : (ev.hasWarnings ? 'замечания' : 'OK'),
          badgeClass: hasErr ? 'db-err' : (ev.hasWarnings ? 'db-warn' : 'db-ok'),
        });
        if (ev.download) addCheckedFile(ev.filename, ev.download);
        setCardStatus(statusId, ev.msg);
      } else if (ev.type === 'error') {
        errCount++;
        details.push({ icon: '✕', name: ev.msg, wrap: true, badge: 'ошибка', badgeClass: 'db-err' });
      } else if (ev.type === 'done') {
        if (bar) bar.style.width = '100%';
        setCardStatus(statusId, 'Проверка завершена', 'done');
        finalizeOpCard(card, statusId, [
          { label: `OK: ${okCount}`, color: 'green' },
          { label: `Замечания: ${warnCount}`, color: 'amber' },
          { label: `Ошибки: ${errCount}`, color: 'red' },
        ], details, { expandDetails: true, detailLabel: 'По файлам' });
        await refreshCheckedList();
      }
    }
  }
}

function downloadPrescriptionsZip() {
  window.location.href = '/download/prescriptions/all.zip';
}

function initHelpPopover() {
  const btn = document.getElementById('helpPrescriptionsBtn');
  const pop = document.getElementById('helpPrescriptionsPopover');
  const backdrop = document.getElementById('helpPrescriptionsBackdrop');
  const close = document.getElementById('helpPrescriptionsClose');
  if (!btn || !pop) return;

  function open() {
    pop.classList.add('is-open');
    if (backdrop) backdrop.classList.add('is-open');
    btn.setAttribute('aria-expanded', 'true');
    refreshActivityEmptyHint();
  }
  function shut() {
    pop.classList.remove('is-open');
    if (backdrop) backdrop.classList.remove('is-open');
    btn.setAttribute('aria-expanded', 'false');
    refreshActivityEmptyHint();
  }
  btn.addEventListener('click', () => (pop.classList.contains('is-open') ? shut() : open()));
  if (close) close.addEventListener('click', shut);
  if (backdrop) backdrop.addEventListener('click', shut);
}

document.addEventListener('DOMContentLoaded', () => {
  initHelpPopover();
  refreshCheckedList();
});

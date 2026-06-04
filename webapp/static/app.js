// ── Утилиты ──────────────────────────────────────────────────────────────

function escHtml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
// ── Activity panel API ────────────────────────────────────────────────────

const MAX_OP_CARDS = 20;

function getActivityTrackEl() {
  return document.getElementById('activityPanelScroll');
}

function getActivityScrollEl() {
  return document.getElementById('activityPanelInner')
    || getActivityTrackEl()
    || document.getElementById('activityPanel');
}

function refreshActivityEmptyHint() {
  const emptyHint = document.getElementById('emptyHint');
  const panel = getActivityScrollEl();
  const pop = document.getElementById('helpPreparePopover');
  if (!emptyHint || !panel) return;
  const scrollRoot = document.getElementById('activityPanelInner') || panel;
  const hasCards = scrollRoot.querySelectorAll('.op-card').length > 0;
  const helpOpen = pop && pop.classList.contains('is-open');
  emptyHint.style.display = (hasCards || helpOpen) ? 'none' : '';
}

const OP_SUBTITLES = {
  'Подготовка отчётов': 'Меняется содержимое загруженных .docx; имена файлов на диске те же.',
  'Дата в тексте болванок': 'Меняется текст внутри шаблонов; файлы на диске не переименовываются.',
  'Переименование готовых': 'Готовые сводные: новое имя на диске и дата в документе.',
  'Проверка отчётов': 'AI-проверка; правки в загрузку (temp) — сразу «Подготовить» и дальше. Скачать с именем _исправлен — панель «Исправленные».',
  'Сборка отчётов': 'Склейка по компаниям в папку результатов.',
};

function _trimHistory() {
  const panel = getActivityScrollEl();
  const cards = panel.querySelectorAll('.op-card');
  if (cards.length > MAX_OP_CARDS) {
    for (let i = 0; i < cards.length - MAX_OP_CARDS; i++) cards[i].remove();
  }
}

function collapseOtherOpCards(keepCard) {
  document.querySelectorAll('.op-card.summary-card').forEach(c => {
    if (c !== keepCard) c.classList.add('is-collapsed');
  });
}

function attachOpCardCollapse(card) {
  const btn = card.querySelector('.op-card-collapse');
  if (!btn) return;
  btn.addEventListener('click', () => {
    card.classList.toggle('is-collapsed');
    btn.textContent = card.classList.contains('is-collapsed') ? '▶' : '▼';
  });
}

function createOpCard(title, subtitle) {
  const panel = getActivityScrollEl();
  const statusId = 'ls_' + Date.now() + '_' + Math.random().toString(36).slice(2);
  const progressId = 'pr_' + Date.now() + '_' + Math.random().toString(36).slice(2);
  const sub = subtitle || OP_SUBTITLES[title] || '';
  const card = document.createElement('div');
  card.className = 'op-card summary-card is-running';
  card.innerHTML = `
    <div class="summary-card-header">
      <div class="summary-card-title-wrap">
        <div class="summary-card-title">${escHtml(title)}</div>
        ${sub ? `<div class="summary-card-subtitle">${escHtml(sub)}</div>` : ''}
      </div>
      <span class="summary-card-meta">${new Date().toLocaleTimeString('ru')}</span>
      <button type="button" class="op-card-collapse" aria-label="Свернуть карточку">▼</button>
    </div>
    <div class="op-card-body">
      <div class="live-status op-live" id="${statusId}">
        <div class="live-status-row">
          <span class="spinner"></span>
          <span class="live-status-text"></span>
        </div>
        <div class="op-progress-wrap" id="${progressId}">
          <div class="op-progress-track"><div class="op-progress-fill"></div></div>
          <div class="op-progress-caption">
            <span class="op-progress-label">Ожидание…</span>
            <span class="op-progress-pct">0%</span>
          </div>
        </div>
      </div>
    </div>
  `;
  panel.insertBefore(card, document.getElementById('emptyHint'));
  attachOpCardCollapse(card);
  collapseOtherOpCards(card);
  refreshActivityEmptyHint();
  const track = getActivityTrackEl();
  if (track) track.scrollTop = 0;
  _trimHistory();
  return { card, statusId, progressId };
}

function setCardProgress(statusId, progressId, current, total, msg, state) {
  const el = document.getElementById(statusId);
  if (!el) return;
  const card = el.closest('.op-card');
  el.classList.remove('done', 'error');
  const spinner = el.querySelector('.spinner');
  const txt = el.querySelector('.live-status-text');
  const wrap = progressId ? document.getElementById(progressId) : null;
  const fill = wrap ? wrap.querySelector('.op-progress-fill') : null;
  const label = wrap ? wrap.querySelector('.op-progress-label') : null;
  const pctEl = wrap ? wrap.querySelector('.op-progress-pct') : null;

  if (state === 'done') {
    el.classList.add('done');
    if (spinner) spinner.style.display = 'none';
    if (card) { card.classList.remove('is-running'); card.classList.add('is-done'); }
  } else if (state === 'error') {
    el.classList.add('error');
    if (spinner) spinner.style.display = 'none';
    if (card) { card.classList.remove('is-running'); card.classList.add('is-error'); }
  } else {
    if (spinner) spinner.style.display = '';
    if (card) card.classList.add('is-running');
  }
  if (txt && msg != null) txt.textContent = msg;

  const totalN = total > 0 ? total : 0;
  const cur = totalN ? Math.min(current, totalN) : 0;
  const pct = totalN ? Math.round((cur / totalN) * 100) : (state === 'done' ? 100 : 0);
  if (fill) fill.style.width = pct + '%';
  if (pctEl) pctEl.textContent = totalN ? `${pct}% (${cur}/${totalN})` : (pct ? pct + '%' : '—');
  if (label && totalN) label.textContent = msg || 'Выполняется…';
}

function setCardStatus(statusId, msg, state, progressId, progress) {
  const p = progress || {};
  setCardProgress(
    statusId,
    progressId || null,
    p.current || 0,
    p.total || 0,
    msg,
    state
  );
}

function formatLogLine(line) {
  const body = line.replace(/^\[(OK|ERR)\]\s*/, '');
  if (body.includes('имя файла на диске не менялось')) {
    return body.replace(
      /;\s*имя файла на диске не менялось/,
      ' · файл на диске: то же имя'
    );
  }
  if (body.includes('на диске') && body.includes('в документе')) {
    return body;
  }
  return body;
}

function finalizeOpCard(card, statusId, stats, details, fileCards, options) {
  options = options || {};
  const liveEl = document.getElementById(statusId);
  if (liveEl) liveEl.remove();
  card.classList.remove('is-running');
  if (options.errorState) card.classList.add('is-error');
  else card.classList.add('is-done');

  const body = card.querySelector('.op-card-body') || card;
  const statsContainer = document.createElement('div');
  statsContainer.className = 'summary-stats';
  statsContainer.innerHTML = stats.map(s =>
    `<span class="stat-chip stat-${s.color}">${escHtml(s.label)}</span>`
  ).join('');
  body.appendChild(statsContainer);

  if (options.meta) {
    const metaEl = document.createElement('div');
    metaEl.style.cssText = 'font-size:12px;color:#6b7280;margin-top:2px';
    metaEl.textContent = options.meta;
    body.appendChild(metaEl);
  }

  const hasFileCards = Array.isArray(fileCards) && fileCards.length > 0;
  const hasDetails = (details && details.length) || hasFileCards;
  if (hasDetails) {
    const detailId = 'det_' + Date.now() + '_' + Math.random().toString(36).slice(2);
    const toggleBtn = document.createElement('button');
    toggleBtn.className = 'detail-toggle';
    toggleBtn.type = 'button';
    const detailLabel = options.detailLabel || 'Подробный лог';
    toggleBtn.textContent = options.expandDetails ? `${detailLabel} ▲` : `${detailLabel} ▼`;
    toggleBtn.onclick = () => {
      const block = document.getElementById(detailId);
      block.classList.toggle('open');
      toggleBtn.textContent = block.classList.contains('open') ? `${detailLabel} ▲` : `${detailLabel} ▼`;
    };
    body.appendChild(toggleBtn);

    const detailBlock = document.createElement('div');
    detailBlock.className = 'detail-block log-timeline' + (options.expandDetails ? ' open' : '');
    detailBlock.id = detailId;

    if (details && details.length) {
      details.forEach(d => {
        const row = document.createElement('div');
        row.className = 'detail-row' + (d.indent ? ' indent' : '') + (d.fileRow ? ' file-row' : '');
        const nameCls = 'detail-name' + (d.wrap ? ' wrap' : '');
        row.innerHTML = `
          <span class="detail-icon">${d.icon || ''}</span>
          <span class="${nameCls}" title="${escHtml(d.name)}">${escHtml(d.name)}</span>
          ${d.badge ? `<span class="detail-badge ${d.badgeClass || ''}">${escHtml(d.badge)}</span>` : ''}
        `;
        detailBlock.appendChild(row);
      });
    }

    if (hasFileCards) {
      fileCards.forEach(fc => detailBlock.appendChild(fc));
    }

    body.appendChild(detailBlock);
  }
}

function buildFileCardEl(filename, hasErrors, reportText, downloadUrl) {
  const bodyId = 'rb_' + Date.now() + '_' + Math.random().toString(36).slice(2);
  const badgeHtml = hasErrors
    ? '<span class="badge badge-err">⚠ Проблемы</span>'
    : '<span class="badge badge-ok">✓ ОК</span>';
  const bodyHtml = reportText
    ? `<div class="report-body" id="${bodyId}">${escHtml(reportText)}</div>`
    : '';
  const toggleHtml = reportText
    ? `<button class="detail-toggle" onclick="toggleReportBody('${bodyId}', this)" style="margin-top:6px">Показать отчёт ▼</button>`
    : '';
  const dlHtml = downloadUrl
    ? `<a href="${downloadUrl}" download>Скачать исправленный</a>`
    : '';
  const el = document.createElement('div');
  el.className = 'file-card ' + (hasErrors ? 'file-card-warn' : 'file-card-ok');
  el.style.marginTop = '8px';
  el.dataset.filename = filename;
  el.innerHTML = `
    <div class="file-card-head">
      <span class="file-card-name" title="${escHtml(filename)}">${escHtml(filename)}</span>
      ${badgeHtml}
    </div>
    ${toggleHtml}
    ${bodyHtml}
    ${dlHtml ? `<div class="file-card-actions">${dlHtml}</div>` : ''}
  `;
  return el;
}

function toggleReportBody(id, btn) {
  const el = document.getElementById(id);
  const open = el.style.display === 'block';
  el.style.display = open ? 'none' : 'block';
  btn.textContent = open ? 'Показать отчёт ▼' : 'Скрыть отчёт ▲';
}

async function apiFetch(url, options = {}, statusId) {
  try {
    const res = await fetch(url, options);
    const contentType = (res.headers.get('content-type') || '').toLowerCase();
    let data = null;
    let rawText = '';

    if (contentType.includes('application/json')) {
      data = await res.json();
    } else {
      rawText = await res.text();
      try {
        data = rawText ? JSON.parse(rawText) : {};
      } catch (_) {
        data = { detail: rawText || res.statusText || 'Ошибка сервера' };
      }
    }

    if (!res.ok) {
      const detail = (data && (data.detail || data.message)) || rawText || res.statusText;
      throw new Error(typeof detail === 'string' ? detail : JSON.stringify(detail));
    }
    return data;
  } catch (e) {
    if (statusId) setCardStatus(statusId, 'Ошибка: ' + e.message, 'error');
    throw e;
  }
}
// ── Файлы ─────────────────────────────────────────────────────────────────

let fixedFiles = [];

function addFixed(filename, downloadUrl) {
  if (fixedFiles.includes(downloadUrl)) return;
  fixedFiles.push(downloadUrl);
  const container = document.getElementById('fixedFiles');
  const noRes = container.querySelector('.no-results');
  if (noRes) noRes.remove();
  const item = document.createElement('div');
  item.className = 'result-item';
  item.innerHTML = `<span title="${escHtml(filename)}">${escHtml(filename)}</span><a href="${downloadUrl}" download>Скачать</a>`;
  container.appendChild(item);
}

async function refreshReportsList() {
  const data = await apiFetch('/files/reports');
  const files = data.files || [];
  const el = document.getElementById('list-reports');
  el.innerHTML = files.length
    ? files.map(f => `<div>${escHtml(f)}</div>`).join('')
    : '<div style="color:#aaa;font-style:italic">—</div>';
}

async function refreshResults() {
  const data = await apiFetch('/results');
  const el = document.getElementById('results-list');
  el.innerHTML = data.files && data.files.length
    ? data.files.map(f =>
        `<div class="result-item">
           <span title="${escHtml(f)}">${escHtml(f)}</span>
           <a href="/download/${encodeURIComponent(f)}" download>Скачать</a>
         </div>`
      ).join('')
    : '<div class="no-results">Пока нет</div>';
}
// ── Загрузка отчётов ──────────────────────────────────────────────────────

async function uploadReports(input) {
  const fd = new FormData();
  for (const f of input.files) fd.append('files', f);
  const { card, statusId } = createOpCard('Загрузка файлов');
  setCardStatus(statusId, `Загружаю ${input.files.length} файл(ов)...`);
  try {
    const data = await apiFetch('/upload/reports', { method: 'POST', body: fd }, statusId);
    finalizeOpCard(card, statusId, [{ label: `Загружено: ${data.count}`, color: 'green' }], null, null);
    refreshReportsList();
  } catch (_) {}
}

async function clearReports() {
  const { card, statusId } = createOpCard('Очистка отчётов');
  setCardStatus(statusId, 'Удаляю загруженные отчёты...');
  try {
    await apiFetch('/clear/reports', { method: 'DELETE' }, statusId);
    finalizeOpCard(card, statusId, [{ label: 'Удалено', color: 'blue' }], null, null);
    refreshReportsList();
  } catch (_) {}
}
// ── Проверка и исправление (SSE) ──────────────────────────────────────────

async function startCheck() {
  const btn = document.getElementById('btnCheck');
  btn.disabled = true;
  btn.textContent = 'Проверяю...';
  document.getElementById('fixedFiles').innerHTML = '<div class="no-results">Пока нет</div>';
  fixedFiles = [];

  const bar = document.getElementById('progressBar');
  bar.style.width = '10%';

  const { card, statusId, progressId } = createOpCard('Проверка отчётов');
  let total = 0;
  setCardProgress(statusId, progressId, 0, 0, 'Запускаю проверку…');

  const resp = await fetch('/check/descriptions/stream', { method: 'POST' });
  if (!resp.ok) {
    setCardProgress(statusId, progressId, 0, 0, 'Ошибка запуска проверки', 'error');
    finalizeOpCard(card, statusId, [{ label: 'Ошибка', color: 'red' }], null, null, { errorState: true });
    btn.disabled = false;
    btn.textContent = 'Проверить и исправить';
    return;
  }

  const reader = resp.body.getReader();
  const decoder = new TextDecoder();
  let buf = '';
  let processed = 0;
  const collectedFileCards = [];
  const downloadMap = {};

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += decoder.decode(value, { stream: true });
    const lines = buf.split('\n');
    buf = lines.pop();
    for (const line of lines) {
      if (!line.startsWith('data: ')) continue;
      let ev;
      try { ev = JSON.parse(line.slice(6)); } catch { continue; }

      if (ev.type === 'start') {
        total = ev.total || 0;
        setCardProgress(statusId, progressId, 0, total, ev.msg);
        bar.style.width = total ? '12%' : '20%';
      } else if (ev.type === 'info') {
        setCardProgress(statusId, progressId, processed, total, ev.msg);
      } else if (ev.type === 'report') {
        processed++;
        const pctBar = total ? Math.min(12 + (processed / total) * 78, 90) : Math.min(20 + processed * 30, 85);
        bar.style.width = pctBar + '%';
        const shortMsg = ev.hasErrors
          ? `${ev.filename}: замечания по описаниям`
          : `${ev.filename}: без замечаний`;
        setCardProgress(statusId, progressId, processed, total, shortMsg);
        collectedFileCards.push({ filename: ev.filename, hasErrors: ev.hasErrors, reportText: ev.result?.report || '', downloadUrl: null });
      } else if (ev.type === 'fixed') {
        addFixed(ev.filename, ev.download);
        downloadMap[ev.filename] = ev.download;
      } else if (ev.type === 'done') {
        bar.style.width = '100%';
        const s = ev.summary || {};
        const total = s.total || 0;
        const errors = s.errors || 0;
        const promoted = s.promoted || 0;
        const fileCardEls = collectedFileCards.map(fc =>
          buildFileCardEl(fc.filename, fc.hasErrors, fc.reportText, downloadMap[fc.filename] || null)
        );
        setCardProgress(statusId, progressId, total || processed, total || processed, 'Проверка завершена', 'done');
        finalizeOpCard(card, statusId, [
          { label: `Файлов: ${total}`, color: 'blue' },
          { label: `Без замечаний: ${total - errors}`, color: 'green' },
          ...(errors > 0 ? [{ label: `С замечаниями: ${errors}`, color: 'amber' }] : []),
          ...(promoted > 0 ? [{ label: `В загрузке обновлено: ${promoted}`, color: 'green' }] : []),
        ], null, fileCardEls, {
          expandDetails: true,
          detailLabel: 'Отчёты по файлам',
        });
        if (promoted > 0) refreshReportsList();
        setTimeout(() => { bar.style.width = '0%'; }, 2000);
        btn.disabled = false;
        btn.textContent = 'Проверить и исправить';
      } else if (ev.type === 'error') {
        setCardProgress(statusId, progressId, processed, total, ev.msg, 'error');
      }
    }
  }
  btn.disabled = false;
  btn.textContent = 'Проверить и исправить';
}
// ── Руководитель ──────────────────────────────────────────────────────────

async function switchLeader(leader) {
  const names = {
    'aniskov': 'Аниськов Владимир Иванович',
    'mandzhiev': 'Манджиев Игорь Александрович (И.О.)'
  };
  const { card, statusId } = createOpCard(`Руководитель: ${names[leader]}`);
  setCardStatus(statusId, `Переключаю руководителя на: ${names[leader]}...`);
  try {
    const data = await apiFetch(`/switch-leader-ai/${leader}`, { method: 'POST' }, statusId);
    if (data.ok) {
      const lines = data.message.split('\n').filter(l => l.trim());
      const details = lines.map(l => ({
        icon: l.startsWith('→') ? '✓' : '',
        name: l.replace(/^→\s*/, ''),
        badge: l.includes('замен') ? 'изменено' : null,
        badgeClass: 'db-ok',
      }));
      finalizeOpCard(card, statusId, [
        { label: `Файлов: ${lines.filter(l => l.startsWith('→')).length}`, color: 'green' },
      ], details, null);
    } else {
      setCardStatus(statusId, `Ошибка: ${data.message || 'Неизвестная ошибка'}`, 'error');
      finalizeOpCard(card, statusId, [{ label: 'Ошибка', color: 'red' }], null, null);
    }
  } catch (_) {}
}
// ── Макросы ───────────────────────────────────────────────────────────────

(function initMacroDate() {
  const el = document.getElementById('macroReportDate');
  if (el) el.value = new Date().toISOString().slice(0, 10);
})();

(function initResultsPanelResize() {
  const handle = document.getElementById('resultsResizeHandle');
  const panel = document.getElementById('resultsPanel');
  if (!handle || !panel) return;

  const MIN = 280;
  const maxWidth = () => Math.min(560, Math.floor(window.innerWidth * 0.55));

  try {
    const saved = parseInt(localStorage.getItem('resultsPanelWidth'), 10);
    if (saved >= MIN && saved <= maxWidth()) panel.style.width = saved + 'px';
  } catch (_) {}

  handle.addEventListener('mousedown', (e) => {
    e.preventDefault();
    const startX = e.clientX;
    const startW = panel.offsetWidth;
    handle.classList.add('is-dragging');
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';

    function onMove(ev) {
      const w = Math.min(maxWidth(), Math.max(MIN, startW + (startX - ev.clientX)));
      panel.style.width = w + 'px';
    }
    function onUp() {
      handle.classList.remove('is-dragging');
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      try { localStorage.setItem('resultsPanelWidth', String(panel.offsetWidth)); } catch (_) {}
    }
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  });
})();

(function initPrepareBlockHelp() {
  const btn = document.getElementById('helpPrepareBtn');
  const closeBtn = document.getElementById('helpPrepareClose');
  const pop = document.getElementById('helpPreparePopover');
  const backdrop = document.getElementById('helpPrepareBackdrop');
  if (!btn || !pop) return;

  function setOpen(open) {
    pop.classList.toggle('is-open', open);
    if (backdrop) {
      backdrop.classList.toggle('is-open', open);
      backdrop.setAttribute('aria-hidden', open ? 'false' : 'true');
    }
    btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    refreshActivityEmptyHint();
  }

  function toggle() {
    setOpen(!pop.classList.contains('is-open'));
  }

  window.closePrepareHelp = () => setOpen(false);

  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    toggle();
  });
  if (closeBtn) {
    closeBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      setOpen(false);
    });
  }
  if (backdrop) {
    backdrop.addEventListener('click', () => setOpen(false));
  }

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && pop.classList.contains('is-open')) setOpen(false);
  });
})();
function getMacroReportDate() {
  const dateVal = document.getElementById('macroReportDate')?.value;
  if (!dateVal) {
    throw new Error('Выберите дату в календаре (блок «Подготовка к сборке»)');
  }
  return dateVal;
}


function logLinesToDetails(log, extra) {
  const details = [];
  extra = extra || {};
  if (extra.meta) {
    details.push({ icon: 'ℹ', name: extra.meta, wrap: true, badge: 'Параметры', badgeClass: '' });
  }
  if (extra.grid_cols && extra.grid_cols.length) {
    details.push({
      icon: '⊞',
      name: 'Колонки (dxa): ' + extra.grid_cols.join(' + '),
      wrap: true,
      badge: 'Сетка',
      badgeClass: '',
    });
  }
  for (const line of (Array.isArray(log) ? log : [])) {
    const isErr = line.startsWith('[ERR]');
    const body = line.replace(/^\[(OK|ERR)\]\s*/, '');
    const colon = body.indexOf(':');
    if (colon > 0) {
      const file = body.slice(0, colon).trim();
      const rest = body.slice(colon + 1).trim();
      details.push({
        icon: isErr ? '✗' : '📄',
        name: file,
        badge: isErr ? 'Ошибка' : 'Файл',
        badgeClass: isErr ? 'db-err' : 'db-ok',
        fileRow: true,
      });
      if (rest) {
        rest.split(/,\s*/).forEach((part, i) => {
          const p = part.trim();
          if (!p) return;
          const warn = p.includes('⚠');
          details.push({
            icon: warn ? '⚠' : '→',
            name: p,
            badge: warn ? 'в документе' : 'в документе',
            badgeClass: warn ? 'db-warn' : '',
            wrap: true,
            indent: true,
          });
        });
      }
    } else {
      details.push({
        icon: isErr ? '✗' : '✓',
        name: body,
        badge: isErr ? 'Ошибка' : 'OK',
        badgeClass: isErr ? 'db-err' : 'db-ok',
        wrap: true,
      });
    }
  }
  return details;
}

async function runBolvankaDateUpdate(dateVal) {
  const { card, statusId, progressId } = createOpCard('Дата в тексте болванок');
  setCardProgress(statusId, progressId, 0, 0, `Обновляю текст в шаблонах на ${dateVal}…`);
  try {
    const data = await apiFetch('/rename/templates', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ date: dateVal }),
    }, statusId);
    const lines = Array.isArray(data.log) ? data.log : [];
    const details = lines.map(line => ({
      icon: line.startsWith('[ERR]') ? '✗' : '✓',
      name: formatLogLine(line),
      badge: line.startsWith('[ERR]') ? 'Ошибка' : 'в документе',
      badgeClass: line.startsWith('[ERR]') ? 'db-err' : 'db-ok',
      wrap: true,
    }));
    const errCount = lines.filter(l => l.startsWith('[ERR]')).length;
    setCardProgress(statusId, progressId, lines.length, lines.length, 'Готово', 'done');
    finalizeOpCard(card, statusId, [
      { label: `Дата в тексте: ${data.date || dateVal}`, color: 'blue' },
      { label: 'Содержимое .docx обновлено', color: 'green' },
      { label: `Шаблонов: ${lines.length - errCount}`, color: 'green' },
      ...(errCount ? [{ label: `Ошибок: ${errCount}`, color: 'red' }] : []),
    ], details, null, { expandDetails: true, detailLabel: 'По шаблонам' });
  } catch (e) {
    finalizeOpCard(card, statusId, [{ label: 'Ошибка', color: 'red' }], [
      { icon: '✗', name: e.message, wrap: true, badge: 'Ошибка', badgeClass: 'db-err' },
    ], null, { expandDetails: true });
  }
}

async function prepareReports() {
  const btn = document.getElementById('btnPrepare');
  if (!btn) return;
  let dateVal;
  try {
    dateVal = getMacroReportDate();
  } catch (e) {
    alert(e.message);
    return;
  }
  btn.disabled = true;
  const { card, statusId, progressId } = createOpCard('Подготовка отчётов');
  setCardProgress(statusId, progressId, 0, 0, 'Правки внутри загруженных файлов…');
  try {
    const data = await apiFetch('/macro/prepare', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ date: dateVal }),
    }, statusId);
    const lines = Array.isArray(data.log) ? data.log : [`[ERR] Нет log в ответе: ${JSON.stringify(data)}`];
    const meta = `Дата: ${data.date || '—'} · ${data.template || 'сетка'}`;
    const details = logLinesToDetails(lines, { meta, grid_cols: data.grid_cols });
    const errCount = lines.filter(l => l.startsWith('[ERR]')).length;
    setCardProgress(statusId, progressId, lines.length, lines.length, 'Готово', 'done');
    finalizeOpCard(card, statusId, [
      { label: 'Содержимое .docx обновлено', color: 'green' },
      { label: 'Имена на диске без изменений', color: 'blue' },
      { label: `Файлов: ${lines.length - errCount}`, color: 'green' },
      ...(errCount ? [{ label: `Ошибок: ${errCount}`, color: 'red' }] : []),
    ], details, null, { expandDetails: true, detailLabel: 'По файлам' });
  } catch (e) {
    finalizeOpCard(card, statusId, [{ label: 'Ошибка', color: 'red' }], [
      { icon: '✗', name: e.message, wrap: true, badge: 'Ошибка', badgeClass: 'db-err' },
    ], null, { expandDetails: true });
  }
  await runBolvankaDateUpdate(dateVal);
  btn.disabled = false;
}

async function runRenameResults(dateVal) {
  const { card, statusId, progressId } = createOpCard('Переименование готовых');
  setCardProgress(statusId, progressId, 0, 0, `Дата ${dateVal}: переименование на диске и в документе…`);
  try {
    const data = await apiFetch('/rename/results', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ date: dateVal }),
    }, statusId);
    const lines = Array.isArray(data.log) ? data.log : [];
    const details = lines.map(line => ({
      icon: line.startsWith('[ERR]') ? '✗' : '✓',
      name: formatLogLine(line),
      badge: line.startsWith('[ERR]') ? 'Ошибка' : 'диск + документ',
      badgeClass: line.startsWith('[ERR]') ? 'db-err' : 'db-ok',
      wrap: true,
    }));
    const errCount = lines.filter(l => l.startsWith('[ERR]')).length;
    setCardProgress(statusId, progressId, lines.length, lines.length, 'Готово', 'done');
    finalizeOpCard(card, statusId, [
      { label: `Дата: ${data.date || dateVal}`, color: 'blue' },
      { label: `Готовых файлов: ${lines.length - errCount}`, color: 'green' },
      ...(errCount ? [{ label: `Ошибок: ${errCount}`, color: 'red' }] : []),
    ], details.length ? details : null, null, { expandDetails: true, detailLabel: 'По файлам' });
  } catch (e) {
    finalizeOpCard(card, statusId, [{ label: 'Ошибка', color: 'red' }], [
      { icon: '✗', name: e.message, wrap: true, badge: 'Ошибка', badgeClass: 'db-err' },
    ], null, { expandDetails: true });
  }
}
// ── Сборка всех отчётов (SSE) ─────────────────────────────────────────────

async function mergeAll() {
  let dateVal;
  try {
    dateVal = getMacroReportDate();
  } catch (e) {
    alert(e.message);
    return;
  }

  const { card, statusId, progressId } = createOpCard('Сборка отчётов');
  let mergeTotal = 0;
  let mergeStep = 0;
  setCardProgress(statusId, progressId, 0, 0, 'Формирую сводные отчёты…');
  const bar = document.getElementById('progressBar');
  bar.style.width = '10%';
  let mergeFinished = false;

  try {
    const response = await fetch('/merge/all/stream');
    if (!response.ok) throw new Error(response.statusText);

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    let successCount = 0;

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n\n');
      buffer = lines.pop();

      for (const line of lines) {
        if (!line.startsWith('data: ')) continue;
        try {
          const ev = JSON.parse(line.slice(6));
          if (ev.type === 'start') {
            mergeTotal = ev.total || 0;
            setCardProgress(statusId, progressId, 0, mergeTotal, ev.msg);
            bar.style.width = '15%';
          } else if (ev.type === 'progress') {
            mergeStep = ev.current || mergeStep;
            setCardProgress(statusId, progressId, mergeStep, ev.total || mergeTotal, ev.msg);
            const t = ev.total || mergeTotal || 1;
            bar.style.width = Math.min(15 + (mergeStep / t) * 75, 90) + '%';
          } else if (ev.type === 'info') {
            if (!ev.msg.includes('Найдено 0 отчётов')) {
              setCardProgress(statusId, progressId, mergeStep, mergeTotal, ev.msg);
            }
          } else if (ev.type === 'success') {
            successCount++;
            setCardProgress(statusId, progressId, mergeStep, mergeTotal, ev.msg);
            bar.style.width = Math.min(20 + successCount * 10, 85) + '%';
          } else if (ev.type === 'warning') {
            setCardProgress(statusId, progressId, mergeStep, mergeTotal, ev.msg);
          } else if (ev.type === 'error') {
            setCardProgress(statusId, progressId, mergeStep, mergeTotal, ev.msg, 'error');
          } else if (ev.type === 'done') {
            mergeFinished = true;
            bar.style.width = '100%';
            setCardProgress(statusId, progressId, mergeTotal || mergeStep, mergeTotal || mergeStep, 'Сборка завершена', 'done');
            const merged = ev.results || [];
            const unmatched = ev.unmatched || [];
            const unknown = ev.unmatched_unknown || [];
            const errCount = (ev.errors || []).length;

            const details = [
              ...merged.map(r => ({ icon: '✓', name: `${r.company} (${r.inserted} отч.)`, badge: 'Собран', badgeClass: 'db-ok' })),
              ...unmatched.map(f => ({ icon: '⚠', name: `${f.file} — нет болванки`, badge: f.company, badgeClass: 'db-warn' })),
              ...unknown.map(f => ({ icon: '?', name: f, badge: 'Не распознан', badgeClass: 'db-err' })),
            ];

            finalizeOpCard(card, statusId, [
              { label: `Собрано: ${merged.length}`, color: 'green' },
              ...(unmatched.length > 0 ? [{ label: `Без болванки: ${unmatched.length}`, color: 'amber' }] : []),
              ...(unknown.length > 0 ? [{ label: `Не распознано: ${unknown.length}`, color: 'red' }] : []),
              ...(errCount > 0 ? [{ label: `Ошибок: ${errCount}`, color: 'red' }] : []),
            ], details.length ? details : null, null, { expandDetails: true, detailLabel: 'По компаниям' });

            setTimeout(() => { bar.style.width = '0%'; }, 2000);
          }
        } catch (e) {
          console.error('Parse error:', e);
        }
      }
    }
  } catch (e) {
    setCardStatus(statusId, `Ошибка: ${e.message}`, 'error');
    finalizeOpCard(card, statusId, [{ label: 'Ошибка', color: 'red' }], null, null);
  }

  if (mergeFinished) {
    await runRenameResults(dateVal);
  }
  refreshResults();
}
// ── Скачать ZIP ───────────────────────────────────────────────────────────

function _triggerDownload(url, filename) {
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

async function downloadAll() {
  const { card, statusId } = createOpCard('Скачать всё ZIP');
  setCardStatus(statusId, 'Подготавливаю архив...');
  _triggerDownload('/download/all.zip', 'отчёты.zip');
  finalizeOpCard(card, statusId, [{ label: 'Загрузка началась', color: 'green' }], null, null);
}

// ── Сброс ─────────────────────────────────────────────────────────────────

async function resetAll() {
  if (!confirm('Сбросить всё? Удалятся отчёты и результаты.')) return;
  const { card, statusId } = createOpCard('Сброс всех данных');
  setCardStatus(statusId, 'Удаляю все файлы...');
  try {
    await apiFetch('/clear/all', { method: 'DELETE' }, statusId);
    document.getElementById('fixedFiles').innerHTML = '<div class="no-results">Пока нет</div>';
    fixedFiles = [];
    finalizeOpCard(card, statusId, [{ label: 'Сброшено', color: 'blue' }], null, null);
    refreshReportsList();
    refreshResults();
  } catch (_) {}
}
// ── Drag & Drop ───────────────────────────────────────────────────────────

const uploadZone = document.getElementById('inp-reports').closest('.upload-zone');
uploadZone.addEventListener('dragover', e => { e.preventDefault(); uploadZone.style.borderColor = '#1a56a0'; });
uploadZone.addEventListener('dragleave', () => { uploadZone.style.borderColor = '#aac4e8'; });
uploadZone.addEventListener('drop', e => {
  e.preventDefault();
  uploadZone.style.borderColor = '#aac4e8';
  const input = document.getElementById('inp-reports');
  const dt = new DataTransfer();
  for (const f of e.dataTransfer.files) dt.items.add(f);
  input.files = dt.files;
  uploadReports(input);
});
// ── Инициализация ─────────────────────────────────────────────────────────

refreshReportsList();
refreshResults();
// Global handlers for inline onclick attributes
window.uploadReports = uploadReports;
window.startCheck = startCheck;
window.switchLeader = switchLeader;
window.prepareReports = prepareReports;
window.mergeAll = mergeAll;
window.downloadAll = downloadAll;
window.clearReports = clearReports;
window.resetAll = resetAll;
window.toggleReportBody = toggleReportBody;

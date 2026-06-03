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
  'Проверка отчётов': 'AI-проверка описаний; при необходимости правки вставляются в документ.',
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

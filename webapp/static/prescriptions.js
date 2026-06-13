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

function finalizeOpCard(card, statusId, stats, details, fileCards, options) {
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

  const hasFileCards = Array.isArray(fileCards) && fileCards.length > 0;
  const hasDetails = (details && details.length) || hasFileCards;
  if (hasDetails) {
    const detailId = 'det_' + Date.now() + '_' + Math.random().toString(36).slice(2);
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

    if (details && details.length) {
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
    }

    if (hasFileCards) {
      fileCards.forEach(fc => detailBlock.appendChild(fc));
    }

    body.appendChild(detailBlock);
  }
}

function toggleReportBody(id, btn) {
  const el = document.getElementById(id);
  const open = el.style.display === 'block';
  el.style.display = open ? 'none' : 'block';
  btn.textContent = open ? 'Показать отчёт ▼' : 'Скрыть отчёт ▲';
}

function buildPrescriptionTextBlock(title, text) {
  const body = (text || '').trim();
  if (!body) return '';
  return `<div class="prescription-text-block">`
    + `<div class="prescription-block-title">${escHtml(title)}</div>`
    + `<div class="prescription-text-body">${escHtml(body)}</div>`
    + `</div>`;
}

function buildPrescriptionFileCard(filename, meta) {
  const {
    hasErrors, hasWarnings, reportText, downloadUrl,
    normativeSource, questions, draftLetter, issues,
    engineerContent, agentContent, finalContent,
  } = meta;
  const bodyId = 'rb_' + Date.now() + '_' + Math.random().toString(36).slice(2);
  const badgeHtml = hasErrors
    ? '<span class="badge badge-err">⚠ Ошибки</span>'
    : (hasWarnings
      ? '<span class="badge badge-warn">◐ Замечания</span>'
      : '<span class="badge badge-ok">✓ OK</span>');

  let sourceHtml = '';
  if (normativeSource && normativeSource.label) {
    const srcClass = normativeSource.source === 'techexpert'
      ? 'source-te'
      : (normativeSource.source === 'internet' ? 'source-net' : 'source-unknown');
    const status = normativeSource.status === 'получен' ? '' : ` (${normativeSource.status})`;
    sourceHtml = `<div class="prescription-meta-row">`
      + `<span class="prescription-source-chip ${srcClass}">Нормативка: ${escHtml(normativeSource.label)}${escHtml(status)}</span>`
      + `</div>`;
  }

  let questionsHtml = '';
  if (questions && questions.length) {
    questionsHtml = `<div class="prescription-questions-block">`
      + `<div class="prescription-block-title">Вопросы инженеру (для письма)</div>`
      + `<ol class="prescription-questions">${questions.map(q => `<li>${escHtml(q)}</li>`).join('')}</ol>`
      + `</div>`;
  }

  let letterHtml = '';
  if (draftLetter) {
    letterHtml = `<div class="prescription-letter-block">`
      + `<div class="prescription-block-title">Черновик письма</div>`
      + `<div class="prescription-letter-text">${escHtml(draftLetter)}</div>`
      + `</div>`;
  }

  const warnLines = (issues || [])
    .filter(i => i.level === 'warn' || i.level === 'error')
    .map(i => i.message || i);
  const issuesHtml = warnLines.length
    ? `<ul class="prescription-issues">${warnLines.map(l => `<li>${escHtml(l)}</li>`).join('')}</ul>`
    : '';

  let b18Html = '';
  if (engineerContent || agentContent) {
    b18Html = buildPrescriptionTextBlock('B18 — исходник инженера', engineerContent)
      + buildPrescriptionTextBlock('B18 — переработка агента', agentContent || '(без переработки)');
    if (finalContent && finalContent !== agentContent && finalContent !== engineerContent) {
      b18Html += buildPrescriptionTextBlock('B18 — записано в файл', finalContent);
    }
  }

  const bodyHtml = reportText
    ? `<div class="report-body" id="${bodyId}">${escHtml(reportText)}</div>`
    : '';
  const toggleHtml = reportText
    ? `<button class="detail-toggle" type="button" onclick="toggleReportBody('${bodyId}', this)" style="margin-top:6px">Полный отчёт ▼</button>`
    : '';
  const dlHtml = downloadUrl
    ? `<a href="${downloadUrl}" download>Скачать _проверен</a>`
    : '';
  const el = document.createElement('div');
  el.className = 'file-card ' + (hasErrors ? 'file-card-warn' : (hasWarnings ? 'file-card-warn' : 'file-card-ok'));
  el.style.marginTop = '8px';
  el.innerHTML = `
    <div class="file-card-head">
      <span class="file-card-name" title="${escHtml(filename)}">${escHtml(filename)}</span>
      ${badgeHtml}
    </div>
    ${sourceHtml}
    ${b18Html}
    ${questionsHtml}
    ${letterHtml}
    ${issuesHtml}
    ${toggleHtml}
    ${bodyHtml}
    ${dlHtml ? `<div class="file-card-actions">${dlHtml}</div>` : ''}
  `;
  return el;
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
    ], buildUploadDetails(data.uploaded), null, { expandDetails: true, detailLabel: 'Файлы' });
  } catch (_) {}
  input.value = '';
}

async function clearPrescriptions() {
  const { card, statusId } = createOpCard('Очистка загрузки');
  setCardStatus(statusId, 'Удаляю файлы…');
  try {
    await apiFetch('/clear/prescriptions/uploads', { method: 'DELETE' }, statusId);
    finalizeOpCard(card, statusId, [{ label: 'Загрузка очищена', color: 'blue' }], null, null);
  } catch (_) {}
}

async function resetPrescriptions() {
  const { card, statusId } = createOpCard('Сброс');
  setCardStatus(statusId, 'Очищаю загрузку и результаты…');
  try {
    await apiFetch('/clear/prescriptions/all', { method: 'DELETE' }, statusId);
    document.getElementById('checkedFiles').innerHTML = '<div class="no-results">Пока нет</div>';
    finalizeOpCard(card, statusId, [{ label: 'Сброшено', color: 'blue' }], null, null);
  } catch (_) {}
}

async function checkPrescriptions() {
  const bar = document.getElementById('progressBar');
  if (bar) bar.style.width = '15%';

  const { card, statusId } = createOpCard(
    'Проверка предписаний',
    'B18 — замечание, B19 — нормативка; отчёт по каждому файлу'
  );
  setCardStatus(statusId, 'Запускаю проверку…');

  const fileCards = [];
  let okCount = 0;
  let warnCount = 0;
  let errCount = 0;

  const resp = await fetch('/check/prescriptions/stream', { method: 'POST' });
  if (!resp.ok) {
    setCardStatus(statusId, 'Ошибка запуска', 'error');
    finalizeOpCard(card, statusId, [{ label: 'Ошибка', color: 'red' }], null, null, { errorState: true });
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
        const hasWarn = ev.hasWarnings;
        if (hasErr) errCount++;
        else if (hasWarn) warnCount++;
        else okCount++;

        const issues = (ev.result && ev.result.issues) || [];
        const reportText = (ev.result && (ev.result.review_display || ev.result.report)) || '';

        const result = ev.result || {};
        const fields = result.fields || {};
        const modelCorrected = result.model_corrected || {};
        const corrected = result.corrected || {};
        fileCards.push(
          buildPrescriptionFileCard(ev.filename || ev.msg, {
            hasErrors: hasErr,
            hasWarnings: hasWarn,
            reportText,
            downloadUrl: ev.download,
            normativeSource: result.normative_source,
            questions: result.engineer_questions || [],
            draftLetter: result.draft_letter || '',
            issues,
            engineerContent: fields.content || '',
            agentContent: modelCorrected.content || '',
            finalContent: corrected.content || '',
          })
        );
        if (ev.download) addCheckedFile(ev.filename, ev.download);
        setCardStatus(statusId, ev.msg);
      } else if (ev.type === 'error') {
        errCount++;
        fileCards.push(
          buildPrescriptionFileCard(ev.msg || 'Ошибка', {
            hasErrors: true,
            hasWarnings: false,
            reportText: ev.msg || '',
            downloadUrl: null,
            issues: [{ level: 'error', message: ev.msg }],
          })
        );
      } else if (ev.type === 'done') {
        if (bar) bar.style.width = '100%';
        setCardStatus(statusId, 'Проверка завершена', 'done');
        finalizeOpCard(card, statusId, [
          { label: `OK: ${okCount}`, color: 'green' },
          { label: `Замечания: ${warnCount}`, color: 'amber' },
          { label: `Ошибки: ${errCount}`, color: 'red' },
        ], null, fileCards, { expandDetails: true, detailLabel: 'Отчёты по файлам' });
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

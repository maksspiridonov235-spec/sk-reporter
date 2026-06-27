function escHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

let lastDownloadUrl = '';

function getActivityScrollEl() {
  return document.getElementById('activityPanelInner')
    || document.getElementById('activityPanelScroll')
    || document.getElementById('activityPanel');
}

function refreshActivityEmptyHint() {
  const emptyHint = document.getElementById('emptyHint');
  const panel = getActivityScrollEl();
  if (!emptyHint || !panel) return;
  const hasCards = panel.querySelectorAll('.op-card').length > 0;
  emptyHint.style.display = hasCards ? 'none' : '';
}

function setProgress(active) {
  const bar = document.getElementById('progressBar');
  if (bar) bar.classList.toggle('active', !!active);
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
          <span class="live-status-text">Запуск…</span>
        </div>
      </div>
    </div>
  `;
  panel.insertBefore(card, document.getElementById('emptyHint'));
  card.querySelector('.op-card-collapse').addEventListener('click', () => {
    card.classList.toggle('is-collapsed');
  });
  refreshActivityEmptyHint();
  return { card, statusId };
}

function finalizeCard(card, statusId, logs, ok) {
  const liveEl = document.getElementById(statusId);
  if (liveEl) liveEl.remove();
  card.classList.remove('is-running');
  card.classList.add(ok ? 'is-done' : 'is-error');
  const body = card.querySelector('.op-card-body') || card;
  if (logs && logs.length) {
    const pre = document.createElement('pre');
    pre.className = 'deployment-log';
    pre.style.cssText = 'white-space:pre-wrap;font-size:11px;margin:8px 0;max-height:320px;overflow:auto';
    pre.textContent = logs.join('\n');
    body.appendChild(pre);
  }
}

async function apiFetch(url, options) {
  const resp = await fetch(url, options);
  if (!resp.ok) {
    let detail = resp.statusText;
    try {
      const j = await resp.json();
      detail = j.detail || detail;
    } catch (_) {}
    throw new Error(detail);
  }
  return resp.json();
}

async function refreshStatus() {
  try {
    const data = await apiFetch('/api/deployment/status');
    const el = document.getElementById('uploadStatus');
    const parts = [];
    parts.push(`<div class="result-file">Отчёты: ${data.reports.length ? escHtml(data.reports.join(', ')) : '—'}</div>`);
    parts.push(`<div class="result-file">Прил.7: ${data.has_pril7 ? '✓ загружено' : '—'}</div>`);
    el.innerHTML = parts.join('');

    const resEl = document.getElementById('resultFiles');
    if (data.results && data.results.length) {
      resEl.innerHTML = data.results.map(f =>
        `<div class="result-file"><a href="/download/deployment/${encodeURIComponent(f)}" download>${escHtml(f)}</a></div>`
      ).join('');
      lastDownloadUrl = '/download/deployment/' + encodeURIComponent(data.results[data.results.length - 1]);
      document.getElementById('btnDownload').disabled = false;
    } else {
      resEl.innerHTML = '<div class="no-results">Пока нет</div>';
    }
  } catch (e) {
    console.error(e);
  }
}

async function uploadFiles(url, input) {
  const files = input.files;
  if (!files || !files.length) return;
  const fd = new FormData();
  if (url.includes('reports')) {
    for (const f of files) fd.append('files', f);
  } else {
    fd.append('file', files[0]);
  }
  const { card, statusId } = createOpCard('Загрузка', url.includes('reports') ? 'Отчёты .docx' : 'Excel');
  try {
    await fetch(url, { method: 'POST', body: fd });
    setCardDone(statusId, 'Готово');
    finalizeCard(card, statusId, [], true);
    await refreshStatus();
  } catch (e) {
    setCardError(statusId, String(e.message || e));
    finalizeCard(card, statusId, [String(e.message || e)], false);
  }
  input.value = '';
}

function setCardDone(statusId, msg) {
  const el = document.getElementById(statusId);
  if (!el) return;
  el.classList.add('done');
  const sp = el.querySelector('.spinner');
  if (sp) sp.style.display = 'none';
  const txt = el.querySelector('.live-status-text');
  if (txt) txt.textContent = msg;
}

function setCardError(statusId, msg) {
  const el = document.getElementById(statusId);
  if (!el) return;
  el.classList.add('error');
  const sp = el.querySelector('.spinner');
  if (sp) sp.style.display = 'none';
  const txt = el.querySelector('.live-status-text');
  if (txt) txt.textContent = msg;
}

function uploadReports(input) { uploadFiles('/upload/deployment/reports', input); }
function uploadPril7(input) { uploadFiles('/upload/deployment/pril7', input); }

async function generateDeployment() {
  const dateEl = document.getElementById('reportDate');
  const reportDate = dateEl && dateEl.value;
  if (!reportDate) {
    alert('Выберите дату');
    return;
  }

  const btn = document.getElementById('btnGenerate');
  btn.disabled = true;
  setProgress(true);

  const { card, statusId } = createOpCard('Формирование расстановки', reportDate);
  const logs = [];

  try {
    const fd = new FormData();
    fd.append('report_date', reportDate);
    const resp = await fetch('/api/deployment/generate/stream', { method: 'POST', body: fd });
    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buf = '';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      const parts = buf.split('\n\n');
      buf = parts.pop() || '';
      for (const chunk of parts) {
        const line = chunk.split('\n').find(l => l.startsWith('data: '));
        if (!line) continue;
        const data = JSON.parse(line.slice(6));
        if (data.type === 'log' && data.msg) {
          logs.push(data.msg);
          const txt = document.querySelector(`#${statusId} .live-status-text`);
          if (txt) txt.textContent = data.msg;
        }
        if (data.type === 'done' && data.ok && data.download) {
          lastDownloadUrl = data.download;
          document.getElementById('btnDownload').disabled = false;
        }
        if (data.type === 'error' && data.msg) {
          logs.push('ОШИБКА: ' + data.msg);
        }
      }
    }

    const ok = logs.some(l => l.includes('Архив готов')) || !!lastDownloadUrl;
    setCardDone(statusId, ok ? 'Готово' : 'Ошибка');
    finalizeCard(card, statusId, logs, ok);
    await refreshStatus();
  } catch (e) {
    logs.push(String(e.message || e));
    setCardError(statusId, String(e.message || e));
    finalizeCard(card, statusId, logs, false);
  } finally {
    btn.disabled = false;
    setProgress(false);
  }
}

function downloadResult() {
  if (lastDownloadUrl) window.location.href = lastDownloadUrl;
}

async function clearReports() {
  if (!confirm('Удалить загруженные отчёты?')) return;
  await apiFetch('/clear/deployment/reports', { method: 'DELETE' });
  await refreshStatus();
}

async function resetAll() {
  if (!confirm('Сбросить все файлы расстановки?')) return;
  await apiFetch('/clear/deployment/all', { method: 'DELETE' });
  lastDownloadUrl = '';
  document.getElementById('btnDownload').disabled = true;
  await refreshStatus();
}

document.addEventListener('DOMContentLoaded', () => {
  const d = document.getElementById('reportDate');
  if (d) d.valueAsDate = new Date();
  refreshStatus();
});

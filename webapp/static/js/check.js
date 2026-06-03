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
        const fileCardEls = collectedFileCards.map(fc =>
          buildFileCardEl(fc.filename, fc.hasErrors, fc.reportText, downloadMap[fc.filename] || null)
        );
        setCardProgress(statusId, progressId, total || processed, total || processed, 'Проверка завершена', 'done');
        finalizeOpCard(card, statusId, [
          { label: `Файлов: ${total}`, color: 'blue' },
          { label: `Без замечаний: ${total - errors}`, color: 'green' },
          ...(errors > 0 ? [{ label: `С замечаниями: ${errors}`, color: 'amber' }] : []),
        ], null, fileCardEls, {
          expandDetails: true,
          detailLabel: 'Отчёты по файлам',
        });
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

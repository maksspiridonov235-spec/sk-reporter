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

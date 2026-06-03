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

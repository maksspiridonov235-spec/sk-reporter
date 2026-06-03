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

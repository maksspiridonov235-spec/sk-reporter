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

async function downloadAllFixed() {
  const { card, statusId } = createOpCard('Скачать исправленные ZIP');
  setCardStatus(statusId, 'Подготавливаю архив...');
  _triggerDownload('/download/fixed/all.zip', 'исправленные.zip');
  finalizeOpCard(card, statusId, [{ label: 'Загрузка началась', color: 'green' }], null, null);
}

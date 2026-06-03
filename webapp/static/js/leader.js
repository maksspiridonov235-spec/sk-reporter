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

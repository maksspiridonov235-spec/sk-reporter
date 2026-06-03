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
      { label: 'Имена на диске без изменений', color: 'blue' },
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

// ── Файлы ─────────────────────────────────────────────────────────────────

let fixedFiles = [];

function addFixed(filename, downloadUrl) {
  if (fixedFiles.includes(downloadUrl)) return;
  fixedFiles.push(downloadUrl);
  const container = document.getElementById('fixedFiles');
  const noRes = container.querySelector('.no-results');
  if (noRes) noRes.remove();
  const item = document.createElement('div');
  item.className = 'result-item';
  item.innerHTML = `<span title="${escHtml(filename)}">${escHtml(filename)}</span><a href="${downloadUrl}" download>Скачать</a>`;
  container.appendChild(item);
}

async function refreshReportsList() {
  const data = await apiFetch('/files/reports');
  const files = data.files || [];
  const el = document.getElementById('list-reports');
  el.innerHTML = files.length
    ? files.map(f => `<div>${escHtml(f)}</div>`).join('')
    : '<div style="color:#aaa;font-style:italic">—</div>';
}

async function refreshResults() {
  const data = await apiFetch('/results');
  const el = document.getElementById('results-list');
  el.innerHTML = data.files && data.files.length
    ? data.files.map(f =>
        `<div class="result-item">
           <span title="${escHtml(f)}">${escHtml(f)}</span>
           <a href="/download/${encodeURIComponent(f)}" download>Скачать</a>
         </div>`
      ).join('')
    : '<div class="no-results">Пока нет</div>';
}

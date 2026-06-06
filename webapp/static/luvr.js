(function () {
  const el = document.getElementById("luvrList");

  function esc(s) {
    return String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/"/g, "&quot;");
  }

  function fileTable(files, folder) {
    if (!files?.length) {
      return `<p class="hint-text">Нет файлов${folder ? ` в ${esc(folder)}` : ""}.</p>`;
    }
    const rows = files
      .map(
        (f) =>
          `<tr><td>${esc(f.name)}</td><td>${esc(f.rel || "—")}</td><td>${esc(f.suffix || "—")}</td><td>${f.size_kb ?? "—"}</td></tr>`
      )
      .join("");
    return `<table class="planning-table"><thead><tr><th>Файл</th><th>Путь</th><th>Тип</th><th>КБ</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  fetch("/api/luvr")
    .then((r) => {
      if (!r.ok) throw new Error(r.statusText);
      return r.json();
    })
    .then((data) => {
      el.innerHTML = `<p class="planning-meta">Папка: <code>${esc(data.folder)}</code></p>${fileTable(data.files, data.folder)}`;
    })
    .catch((e) => {
      el.innerHTML = `<p class="error-text">${esc(e.message)}</p>`;
    });
})();

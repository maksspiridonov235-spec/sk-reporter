(function () {
  const el = document.getElementById("luvrList");

  function esc(s) {
    return String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/"/g, "&quot;");
  }

  fetch("/api/luvr")
    .then((r) => {
      if (!r.ok) throw new Error(r.statusText);
      return r.json();
    })
    .then((data) => window.LuvrPanel.render(el, data))
    .catch((e) => {
      if (el) el.innerHTML = `<p class="error-text">${esc(e.message)}</p>`;
    });
})();

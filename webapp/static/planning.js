(function () {
  const tabs = document.querySelectorAll(".planning-tab");
  const panels = {
    projects: document.getElementById("panel-projects"),
    personnel: document.getElementById("panel-personnel"),
    otkk: document.getElementById("panel-otkk"),
  };
  const loaded = {};

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
          `<tr><td>${esc(f.name)}</td><td>${esc(f.suffix || "—")}</td><td>${f.size_kb ?? "—"}</td></tr>`
      )
      .join("");
    return `<table class="planning-table"><thead><tr><th>Файл</th><th>Тип</th><th>КБ</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  function renderProjects(data) {
    const el = document.getElementById("projectsList");
    const items = data.items || [];
    if (!items.length) {
      el.innerHTML = '<p class="hint-text">Проекты не найдены.</p>';
      return;
    }
    el.innerHTML = items
      .map((p) => {
        const badge = p.has_vor_cache ? '<span class="planning-badge">vor.json</span>' : "";
        return `<article class="planning-card">
          <h3>${esc(p.title)} ${badge}</h3>
          <p class="planning-meta"><code>${esc(p.path)}</code>${p.vor_docx ? ` · ВОР: ${esc(p.vor_docx)}` : ""}</p>
          ${fileTable(p.files)}
        </article>`;
      })
      .join("");
  }

  function renderSimpleList(elId, data) {
    const el = document.getElementById(elId);
    let extra = "";
    if (data.people_count != null) {
      extra = `<p class="planning-meta">Записей в personnel.yaml: <strong>${data.people_count}</strong></p>`;
    }
    if (data.count != null) {
      extra = `<p class="planning-meta">Карт в каталоге: <strong>${data.count}</strong></p>`;
    }
    el.innerHTML = extra + fileTable(data.files || data.cards?.map((c) => ({
      name: c.file,
      suffix: (c.file || "").split(".").pop(),
      size_kb: c.size_kb,
    })), data.folder);
  }

  function renderOtkk(data) {
    const el = document.getElementById("otkkList");
    const cards = data.cards || [];
    if (!cards.length) {
      el.innerHTML = '<p class="hint-text">ОТКК не найдены. Запустите build_engineer_data.py --tk</p>';
      return;
    }
    const rows = cards
      .map(
        (c) =>
          `<tr><td>${esc(c.id)}</td><td>${esc(c.file)}</td><td>${c.present ? c.size_kb : "—"}</td><td>${c.present ? "✓" : "—"}</td></tr>`
      )
      .join("");
    el.innerHTML = `<p class="planning-meta">Папка: <code>${esc(data.folder)}</code> · ${cards.length} карт</p>
      <table class="planning-table"><thead><tr><th>ID</th><th>Файл</th><th>КБ</th><th>На диске</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  async function loadTab(name) {
    if (loaded[name]) return;
    const res = await fetch(`/api/planning/${name}`);
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();
    if (name === "projects") renderProjects(data);
    else if (name === "otkk") renderOtkk(data);
    else if (name === "personnel") renderSimpleList("personnelList", data);
    loaded[name] = true;
  }

  function activateTab(name) {
    tabs.forEach((t) => {
      const on = t.dataset.tab === name;
      t.classList.toggle("is-active", on);
      t.setAttribute("aria-selected", on ? "true" : "false");
    });
    Object.entries(panels).forEach(([key, panel]) => {
      panel.hidden = key !== name;
    });
    loadTab(name).catch((e) => {
      const panel = panels[name];
      if (panel) panel.insertAdjacentHTML("beforeend", `<p class="error-text">${esc(e.message)}</p>`);
    });
  }

  tabs.forEach((tab) => {
    tab.addEventListener("click", () => activateTab(tab.dataset.tab));
  });

  const hash = (location.hash || "#projects").replace("#", "");
  activateTab(["projects", "personnel", "otkk"].includes(hash) ? hash : "projects");
})();

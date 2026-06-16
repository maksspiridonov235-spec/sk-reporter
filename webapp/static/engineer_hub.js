(function () {
  const listEl = document.getElementById("engineerHubList");
  const statusEl = document.getElementById("hubStatus");

  function esc(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/"/g, "&quot;");
  }

  function projectSummary(projects) {
    if (!projects?.length) return "Нет объектов";
    const names = projects.map((p) => p.title).slice(0, 3);
    const tail = projects.length > 3 ? ` и ещё ${projects.length - 3}` : "";
    return names.join(" · ") + tail;
  }

  function renderEngineers(engineers) {
    listEl.innerHTML = "";
    if (!engineers.length) {
      statusEl.textContent =
        "Пока нет инженеров с назначениями. Назначьте инженера на проект в Планировании → Проекты.";
      return;
    }
    statusEl.textContent = `Инженеров с объектами: ${engineers.length}`;
    engineers.forEach((eng) => {
      const li = document.createElement("li");
      const href = eng.href || "#";
      const meta = [
        eng.position,
        `${eng.projects_count} ${pluralProjects(eng.projects_count)}`,
      ]
        .filter(Boolean)
        .join(" · ");
      li.innerHTML = `
        <a class="home-card" href="${esc(href)}">
          <span class="home-card-title">${esc(eng.fio)}</span>
          <span class="home-card-desc">${esc(projectSummary(eng.projects))}</span>
          ${meta ? `<span class="home-card-meta">${esc(meta)}</span>` : ""}
        </a>
      `;
      listEl.appendChild(li);
    });
  }

  function pluralProjects(n) {
    const m10 = n % 10;
    const m100 = n % 100;
    if (m10 === 1 && m100 !== 11) return "объект";
    if (m10 >= 2 && m10 <= 4 && (m100 < 10 || m100 >= 20)) return "объекта";
    return "объектов";
  }

  async function loadHub() {
    statusEl.textContent = "Загрузка…";
    const res = await fetch("/api/engineer-hub");
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      statusEl.textContent = err.detail || res.statusText || "Ошибка загрузки";
      return;
    }
    const data = await res.json();
    renderEngineers(data.engineers || []);
  }

  loadHub();
})();

(function () {
  const sectionOnly = document.body.dataset.planningSection;
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
    const engineers = data.engineers_available || [];
    if (!items.length) {
      el.innerHTML = '<p class="hint-text">Проекты не найдены. Добавьте каталог в data/projects/</p>';
      return;
    }
    el.innerHTML = items
      .map((p) => {
        const v = p.vor || {};
        const vorLine = v.ready
          ? `ВОР: ${v.works} работ · ${v.objects} объектов · ${v.stages} этапов`
          : `<span class="warn-text">${esc(v.message || "ВОР не обработан")}</span>`;
        const vorSources = [p.vor_docx, ...(p.vor_doc || [])].filter(Boolean);
        const vorMeta = vorSources.length
          ? ` · ${vorSources.map((f) => esc(f)).join(", ")}`
          : "";
        const tkLine = p.tk_mappings
          ? `Сопоставлений work→ТК: ${p.tk_mappings}`
          : "Сопоставления work→ТК не заданы";
        const assigned = (p.engineers || [])
          .map((e) => `<span class="engineer-chip">${esc(e.fio)}</span>`)
          .join("") || '<span class="hint-text">Инженеры не назначены</span>';
        const assignedCount = (p.engineers || []).length;
        const checkboxes = engineers
          .map((e) => {
            const on = (p.engineer_ids || []).includes(e.id);
            return `<label class="engineer-check">
              <input type="checkbox" class="engineer-cb" value="${esc(e.id)}"${on ? " checked" : ""}/>
              <span>${esc(e.fio)}</span>
            </label>`;
          })
          .join("");
        return `<article class="planning-card" data-project-id="${esc(p.id)}">
          <h3>${esc(p.title)}</h3>
          <p class="planning-meta"><code>${esc(p.path)}</code>${vorMeta}</p>
          <dl class="project-stats">
            <div><dt>ВОР</dt><dd>${vorLine}</dd></div>
            <div><dt>ТК</dt><dd>${tkLine}</dd></div>
            <div><dt>Инженеры</dt><dd class="engineer-chips">${assignedCount ? `<span class="hint-text">${assignedCount}:</span> ` : ""}${assigned}</dd></div>
          </dl>
          <div class="project-assign">
            <label class="field-label">Инженеры на проекте <span class="hint-text">(отметьте несколько)</span></label>
            <div class="engineer-checklist">${checkboxes}</div>
            <button type="button" class="btn btn-primary btn-sm save-engineers">Сохранить</button>
            <span class="save-status hint-text"></span>
          </div>
        </article>`;
      })
      .join("");

    el.querySelectorAll(".save-engineers").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const card = btn.closest(".planning-card");
        const projectId = card.dataset.projectId;
        const status = card.querySelector(".save-status");
        const ids = Array.from(card.querySelectorAll(".engineer-cb:checked")).map((cb) => cb.value);
        btn.disabled = true;
        status.textContent = "Сохранение…";
        try {
          const res = await fetch(`/api/planning/projects/${encodeURIComponent(projectId)}/engineers`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ engineer_ids: ids }),
          });
          if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            throw new Error(err.detail || res.statusText);
          }
          loaded.projects = false;
          await loadTab("projects", true);
          status.textContent = "Сохранено";
        } catch (e) {
          status.textContent = e.message;
          status.classList.add("error-text");
        } finally {
          btn.disabled = false;
        }
      });
    });
  }

  function renderPersonnel(data) {
    const el = document.getElementById("personnelList");
    let people = data.people || [];
    if (!people.length && data.people_count > 0) {
      el.innerHTML =
        '<p class="warn-text">Справочник на диске есть (' +
        data.people_count +
        " записей), но сервер отдаёт старый API. Перезапустите сервер (Ctrl+C → снова uvicorn).</p>";
      return;
    }
    if (!people.length) {
      el.innerHTML =
        '<p class="hint-text">Справочник пуст. Запустите: <code>python scripts/build_engineer_data.py --personnel</code></p>';
      return;
    }

    el.innerHTML = `
      <div class="personnel-toolbar">
        <p class="planning-meta">${data.people_count} сотрудников · <strong>${data.engineers_count}</strong> инженеров СК</p>
        <label class="personnel-filter">
          <input type="checkbox" id="personnelEngineersOnly" checked/>
          Только инженеры СК
        </label>
        <input type="search" id="personnelSearch" class="field-input personnel-search" placeholder="Поиск по ФИО…"/>
      </div>
      <div class="personnel-table-wrap">
        <table class="planning-table personnel-table" id="personnelTable">
          <thead>
            <tr>
              <th>ФИО</th>
              <th>Должность</th>
              <th>Телефон</th>
              <th>Проекты</th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>
      </div>`;

    const tbody = el.querySelector("#personnelTable tbody");
    const engineersOnly = el.querySelector("#personnelEngineersOnly");
    const search = el.querySelector("#personnelSearch");

    function projectCell(projects) {
      if (!projects?.length) return '<span class="hint-text">—</span>';
      return projects.map((pr) => `<span class="engineer-chip">${esc(pr.title)}</span>`).join(" ");
    }

    function renderRows() {
      const q = (search.value || "").trim().toLowerCase();
      const onlyEng = engineersOnly.checked;
      const rows = people.filter((p) => {
        if (onlyEng && !p.is_engineer) return false;
        if (q && !p.fio.toLowerCase().includes(q)) return false;
        return true;
      });
      tbody.innerHTML = rows
        .map(
          (p) =>
            `<tr data-engineer="${p.is_engineer ? "1" : "0"}">
              <td>${esc(p.fio)}</td>
              <td>${esc(p.position || "—")}</td>
              <td>${esc(p.phone || "—")}</td>
              <td class="engineer-chips">${projectCell(p.projects)}</td>
            </tr>`
        )
        .join("");
      if (!rows.length) {
        tbody.innerHTML = '<tr><td colspan="4" class="hint-text">Никого не найдено</td></tr>';
      }
    }

    engineersOnly.addEventListener("change", renderRows);
    search.addEventListener("input", renderRows);
    renderRows();
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

  function renderLuvr(data) {
    const el = document.getElementById("luvrList");
    if (!el) return;

    if (data.cache_ready) {
      const months = data.months || [];
      const defaultSheet = data.default_month || months[months.length - 1]?.sheet || "";
      const sourceLine = data.source
        ? `<strong>${esc(data.source)}</strong>${data.source_kb ? ` · ${data.source_kb} КБ` : ""}`
        : "luvr.yaml";
      const cacheNote = data.cache_from_yaml && !data.xlsx_present
        ? `<p class="hint-text">Таблица из <code>luvr.yaml</code> в репозитории. Чтобы обновить — положите xlsx и запустите <code>python scripts/build_engineer_data.py --luvr</code>.</p>`
        : "";

      el.innerHTML = `
      <div class="luvr-header">
        <p class="planning-meta">Источник: ${sourceLine} · <code>${esc(data.folder)}/</code></p>
        ${data.contract ? `<p class="planning-meta">Договор: ${esc(data.contract)}</p>` : ""}
        ${cacheNote}
        <label class="field-label">Месяц</label>
        <select id="luvrMonth" class="field-input luvr-month-select"></select>
      </div>
      <div class="personnel-table-wrap">
        <table class="planning-table personnel-table" id="luvrTable">
          <thead>
            <tr>
              <th>№</th>
              <th>ФИО</th>
              <th>Должность</th>
              <th>НРС</th>
              <th>Специальность</th>
              <th>Дней на объекте</th>
              <th>Всего отметок</th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>
      </div>`;

      const sel = el.querySelector("#luvrMonth");
      const tbody = el.querySelector("#luvrTable tbody");
      months.forEach((m) => {
        const opt = document.createElement("option");
        opt.value = m.sheet;
        opt.textContent = `${m.sheet}${m.year ? ` ${m.year}` : ""} — ${m.people_count} чел., ${m.days_in_sheet} дн.`;
        if (m.sheet === defaultSheet) opt.selected = true;
        sel.appendChild(opt);
      });

      function renderMonth(sheet) {
        const month = months.find((m) => m.sheet === sheet);
        const people = month?.people || [];
        tbody.innerHTML = people
          .map(
            (p) =>
              `<tr>
              <td>${esc(p.num)}</td>
              <td>${esc(p.fio)}</td>
              <td>${esc(p.position || "—")}</td>
              <td>${esc(p.nrs || "—")}</td>
              <td>${esc(p.specialty || "—")}</td>
              <td>${p.days_present ?? 0}</td>
              <td>${p.days_marked ?? 0}</td>
            </tr>`
          )
          .join("");
        if (!people.length) {
          tbody.innerHTML = '<tr><td colspan="7" class="hint-text">Нет данных за месяц</td></tr>';
        }
      }

      sel.addEventListener("change", () => renderMonth(sel.value));
      renderMonth(sel.value);
      return;
    }

    if (data.xlsx_present) {
      el.innerHTML = `<p class="planning-meta">Файл: <strong>${esc(data.source)}</strong> (${data.source_kb} КБ)</p>
        <p class="warn-text">Кэш не собран. На сервере: <code>python scripts/build_engineer_data.py --luvr</code></p>`;
      return;
    }

    el.innerHTML =
      '<p class="hint-text">Нет данных ЛУВР. Выполните <code>git pull</code> (в репозитории есть <code>data/luvr/luvr.yaml</code>) или положите <code>ЛУВР.xlsx</code> в <code>data/luvr/</code> и запустите build.</p>';
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

  async function loadTab(name, force) {
    if (!force && loaded[name]) return;
    const res = await fetch(`/api/planning/${name}`);
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();
    if (name === "projects") renderProjects(data);
    else if (name === "otkk") renderOtkk(data);
    else if (name === "personnel") renderPersonnel(data);
    else if (name === "luvr") renderLuvr(data);
    loaded[name] = true;
  }

  function showLoadError(name, e) {
    const ids = {
      projects: "projectsList",
      personnel: "personnelList",
      otkk: "otkkList",
      luvr: "luvrList",
    };
    const el = document.getElementById(ids[name]);
    if (el) el.insertAdjacentHTML("beforeend", `<p class="error-text">${esc(e.message)}</p>`);
  }

  if (sectionOnly) {
    loadTab(sectionOnly).catch((e) => showLoadError(sectionOnly, e));
    return;
  }
})();

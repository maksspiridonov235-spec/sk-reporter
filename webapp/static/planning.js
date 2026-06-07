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
    const dep = data.deployment || {};
    const ast = dep.assignments || data.assignment_stats || {};
    const assignLine = ast.projects_total != null
      ? `${ast.assignments || 0} назначений · ${ast.projects_with_engineers || 0}/${ast.projects_total || 0} проектов с инженерами`
      : "";

    const depToolbar =
      dep.template_present
        ? `<div class="luvr-sync-bar projects-toolbar">
            <span class="luvr-link-summary">${assignLine || "Назначьте инженеров на проекты ниже"}</span>
            <div class="luvr-sync-actions">
              <button type="button" class="btn btn-green btn-sm" id="planningBuildDeployment">Сформировать расстановку</button>
            </div>
          </div>
          <p id="planningDeployStatus" class="luvr-save-status" aria-live="polite"></p>`
        : dep.template_error
          ? `<p class="warn-text">Расстановка: ${esc(dep.template_error)}. Положите xlsm в <code>data/luvr/</code>.</p>`
          : "";

    if (!items.length) {
      el.innerHTML =
        depToolbar +
        '<p class="hint-text">Проекты не найдены. Добавьте каталог в data/projects/</p>';
      bindDeploymentButton(el);
      return;
    }
    el.innerHTML =
      depToolbar +
      items
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
          <h3>${esc(p.object_name || p.title)}</h3>
          <p class="planning-meta"><code>${esc(p.id)}</code>${p.title_page ? ` · титул: <code>${esc(p.title_page)}</code>` : ""}</p>
          <dl class="project-stats">
            <div><dt>ВОР</dt><dd>${vorLine}${vorMeta}</dd></div>
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

    bindDeploymentButton(el);

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

  function bindDeploymentButton(root) {
    const btn = root.querySelector("#planningBuildDeployment");
    const status = root.querySelector("#planningDeployStatus");
    if (!btn) return;
    btn.onclick = async () => {
      btn.disabled = true;
      if (status) {
        status.textContent = "Формирование расстановки…";
        status.className = "luvr-save-status luvr-save-status--pending";
      }
      try {
        const res = await fetch("/api/planning/deployment/build", { method: "POST" });
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error(err.detail || res.statusText);
        }
        const result = await res.json();
        if (status) {
          status.textContent = `Готово: ${result.rows_written} строк, ${result.unique_people} чел., ${result.unique_objects} объектов`;
          status.className = "luvr-save-status luvr-save-status--ok";
        }
        if (result.download) window.location.href = result.download;
      } catch (e) {
        if (status) {
          status.textContent = e.message;
          status.className = "luvr-save-status luvr-save-status--error";
        }
      } finally {
        btn.disabled = false;
      }
    };
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
              <td>${esc(p.fio)}${p.id ? ` <span class="hint-text">(${esc(p.id)})</span>` : ""}</td>
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
    window.LuvrPanel.render(document.getElementById("luvrList"), data);
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

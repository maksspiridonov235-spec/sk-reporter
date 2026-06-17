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

  function renderPersonnel(data) {
    const el = document.getElementById("personnelList");
    const db = data.db || {};
    let people = data.people || [];

    const dbError = !db.ok ? (db.error || "PostgreSQL недоступна") : "";
    const storageBadge = `<span class="storage-badge storage-badge--db" title="${esc(dbError)}">PostgreSQL${db.count != null ? ` · ${db.count}` : ""}</span>`;

    const importToolbar = `<div class="personnel-import-bar">
        <label class="btn btn-secondary btn-sm personnel-upload-label">
          Загрузить Excel
          <input type="file" id="personnelUploadXlsx" accept=".xlsx,.xls" hidden/>
        </label>
        <span id="personnelImportStatus" class="hint-text" aria-live="polite"></span>
      </div>`;

    if (!people.length && data.people_count > 0) {
      el.innerHTML =
        '<p class="warn-text">Справочник на диске есть (' +
        data.people_count +
        " записей), но сервер отдаёт старый API. Перезапустите сервер.</p>";
      return;
    }

    if (!people.length) {
      const emptyHint = dbError
        ? `Справочник недоступен: ${esc(dbError)}`
        : "Справочник в базе пуст. Загрузите Excel со списком сотрудников.";
      el.innerHTML =
        `<div class="personnel-toolbar">${storageBadge}</div>` +
        (db.ok ? importToolbar : "") +
        `<p class="hint-text">${emptyHint}</p>`;
      bindPersonnelImport(el);
      return;
    }

    el.innerHTML = `
      <div class="personnel-toolbar">
        ${storageBadge}
        <p class="planning-meta">${data.people_count} сотрудников · <strong>${data.engineers_count}</strong> инженеров СК</p>
        <label class="personnel-filter">
          <input type="checkbox" id="personnelEngineersOnly" checked/>
          Только инженеры СК
        </label>
        <input type="search" id="personnelSearch" class="field-input personnel-search" placeholder="Поиск по ФИО…"/>
      </div>
      ${importToolbar}
      <div class="personnel-table-wrap">
        <table class="planning-table personnel-table" id="personnelTable">
          <thead>
            <tr>
              <th>ФИО</th>
              <th>Должность</th>
              <th>Телефон</th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>
      </div>`;

    bindPersonnelImport(el);

    const tbody = el.querySelector("#personnelTable tbody");
    const engineersOnly = el.querySelector("#personnelEngineersOnly");
    const search = el.querySelector("#personnelSearch");

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
            </tr>`
        )
        .join("");
      if (!rows.length) {
        tbody.innerHTML = '<tr><td colspan="3" class="hint-text">Никого не найдено</td></tr>';
      }
    }

    engineersOnly.addEventListener("change", renderRows);
    search.addEventListener("input", renderRows);
    renderRows();
  }

  function bindPersonnelImport(root) {
    const xlsxInput = root.querySelector("#personnelUploadXlsx");
    const status = root.querySelector("#personnelImportStatus");
    if (!xlsxInput) return;

    async function setStatus(msg, isError) {
      if (!status) return;
      status.textContent = msg;
      status.classList.toggle("error-text", !!isError);
    }

    xlsxInput.addEventListener("change", async () => {
        const file = xlsxInput.files?.[0];
        if (!file) return;
        xlsxInput.disabled = true;
        await setStatus("Загрузка Excel…");
        try {
          const fd = new FormData();
          fd.append("file", file);
          const res = await fetch("/api/planning/personnel/upload-xlsx", {
            method: "POST",
            body: fd,
          });
          const body = await res.json().catch(() => ({}));
          if (!res.ok) throw new Error(body.detail || res.statusText);
          loaded.personnel = false;
          await loadTab("personnel", true);
        } catch (e) {
          await setStatus(e.message, true);
        } finally {
          xlsxInput.value = "";
          xlsxInput.disabled = false;
        }
      });
  }

  function renderSimpleList(elId, data) {
    const el = document.getElementById(elId);
    let extra = "";
    if (data.people_count != null && elId === "personnelList") {
      extra = `<p class="planning-meta">Записей в справочнике: <strong>${data.people_count}</strong></p>`;
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
    const db = data.db || {};
    const cards = data.cards || [];

    const dbError = !db.ok ? (db.error || "PostgreSQL недоступна") : "";
    const storageBadge = `<span class="storage-badge storage-badge--db" title="${esc(dbError)}">PostgreSQL · ${cards.length}</span>`;

    if (!cards.length) {
      const emptyHint = dbError
        ? `База недоступна: ${esc(dbError)}`
        : "В базе пока нет карт ОТКК. Эталоны заливаются при старте сервера из репозитория.";
      el.innerHTML =
        `<div class="personnel-toolbar">${storageBadge}</div>` +
        `<p class="hint-text">${emptyHint}</p>`;
      return;
    }

    el.innerHTML = `
      <div class="personnel-toolbar">
        ${storageBadge}
        <p class="planning-meta">${cards.length} карт в базе</p>
        <input type="search" id="otkkSearch" class="field-input personnel-search" placeholder="Поиск по ID, коду или названию…"/>
      </div>
      <div class="personnel-table-wrap">
        <table class="planning-table personnel-table" id="otkkTable">
          <thead>
            <tr>
              <th>ID</th>
              <th>Код</th>
              <th>Название</th>
              <th></th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>
      </div>
      <dialog id="otkkDetailDialog" class="otkk-detail-dialog">
        <div class="otkk-detail-header">
          <h3 id="otkkDetailTitle"></h3>
          <button type="button" class="btn btn-secondary btn-sm" id="otkkDetailClose">Закрыть</button>
        </div>
        <div id="otkkDetailBody" class="otkk-detail-body"></div>
      </dialog>`;

    bindOtkkDetail(el);

    const tbody = el.querySelector("#otkkTable tbody");
    const search = el.querySelector("#otkkSearch");

    function renderRows() {
      const q = (search.value || "").trim().toLowerCase();
      const rows = cards.filter((c) => {
        if (!q) return true;
        const hay = [c.id, c.code, c.title, c.source_file].join(" ").toLowerCase();
        return hay.includes(q);
      });
      tbody.innerHTML = rows
        .map(
          (c) =>
            `<tr>
              <td><code>${esc(c.id)}</code></td>
              <td>${esc(c.code || "—")}</td>
              <td class="otkk-title-cell">${esc(c.title || "—")}</td>
              <td><button type="button" class="btn btn-link btn-sm otkk-open-btn" data-id="${esc(c.id)}">Открыть</button></td>
            </tr>`
        )
        .join("");
      if (!rows.length) {
        tbody.innerHTML = '<tr><td colspan="4" class="hint-text">Ничего не найдено</td></tr>';
      }
      tbody.querySelectorAll(".otkk-open-btn").forEach((btn) => {
        btn.addEventListener("click", () => openOtkkDetail(btn.dataset.id));
      });
    }

    search.addEventListener("input", renderRows);
    renderRows();
  }

  function bindOtkkDetail(root) {
    const dialog = root.querySelector("#otkkDetailDialog");
    const closeBtn = root.querySelector("#otkkDetailClose");
    if (!dialog || !closeBtn) return;
    closeBtn.addEventListener("click", () => dialog.close());
    dialog.addEventListener("click", (e) => {
      if (e.target === dialog) dialog.close();
    });
  }

  function renderOtkkSegments(segments) {
    if (!Array.isArray(segments) || !segments.length) return "";
    return segments
      .map((seg) => {
        const t = seg.type;
        if (t === "heading") {
          return `<h4 class="otkk-section-heading">${esc(seg.text || "")}</h4>`;
        }
        if (t === "paragraph") {
          return `<p class="otkk-paragraph">${esc(seg.text || "").replace(/\n/g, "<br>")}</p>`;
        }
        if (t === "bullets") {
          const items = (seg.items || []).map((x) => `<li>${esc(x)}</li>`).join("");
          return `<ul class="otkk-bullets">${items}</ul>`;
        }
        if (t === "subbullets") {
          const items = (seg.items || []).map((x) => `<li>${esc(x)}</li>`).join("");
          return `<ul class="otkk-bullets otkk-sub-bullets">${items}</ul>`;
        }
        if (t === "lines") {
          const items = (seg.items || []).map((x) => `<li>${esc(x)}</li>`).join("");
          return `<ul class="otkk-lines">${items}</ul>`;
        }
        if (t === "table") {
          const layout = seg.layout || "standard";
          const headerRows = seg.header_rows || [];
          const headers = seg.headers || [];
          let thead = "";
          if (headerRows.length) {
            thead = headerRows
              .map((row) => `<tr>${row.map((h) => `<th>${esc(h)}</th>`).join("")}</tr>`)
              .join("");
          } else if (headers.length) {
            thead = `<tr>${headers.map((h) => `<th>${esc(h)}</th>`).join("")}</tr>`;
          }
          const colCount = headerRows[0]?.length || headers.length || 3;
          const body = (seg.rows || [])
            .map((row) => {
              if (row.type === "section") {
                const text = esc((row.cells && row.cells[0]) || "");
                return `<tr class="otkk-section-row"><td colspan="${colCount}">${text}</td></tr>`;
              }
              const rowClass = row.type === "subrow" ? "otkk-subrow" : "";
              const cells = (row.cells || [])
                .map((c) => `<td>${esc(c || "").replace(/\n/g, "<br>")}</td>`)
                .join("");
              return `<tr class="${rowClass}">${cells}</tr>`;
            })
            .join("");
          return (
            `<div class="otkk-inner-table-wrap">` +
            `<div class="otkk-table-caption">${esc(seg.caption || "")}</div>` +
            `<table class="otkk-inner-table otkk-inner-table--${esc(layout)}"><thead>${thead}</thead><tbody>${body}</tbody></table>` +
            `</div>`
          );
        }
        return "";
      })
      .join("");
  }

  function renderOtkkRowBody(row) {
    const label = esc(row.label || "");
    let valueHtml;
    if (Array.isArray(row.segments) && row.segments.length) {
      valueHtml = `<div class="otkk-rich">${renderOtkkSegments(row.segments)}</div>`;
    } else {
      valueHtml = `<div class="otkk-value">${esc(row.value || "")}</div>`;
    }
    return `<tr><th scope="row">${label}</th><td>${valueHtml}</td></tr>`;
  }

  function renderContractors(data) {
    const el = document.getElementById("contractorsList");
    if (!el) return;
    const db = data.db || {};
    const contractors = data.contractors || [];
    const dbError = !db.ok ? (db.error || "PostgreSQL недоступна") : "";
    const storageBadge = `<span class="storage-badge storage-badge--db" title="${esc(dbError)}">PostgreSQL · ${contractors.length}</span>`;
    const seedBtn = `<button type="button" class="btn btn-secondary btn-sm" id="contractorsSeedBtn">Загрузить всех из болванок</button>`;

    if (!contractors.length) {
      el.innerHTML =
        `<div class="personnel-toolbar">${storageBadge} ${db.ok ? seedBtn : ""}</div>` +
        `<p class="hint-text">${dbError ? esc(dbError) : "Подрядчиков пока нет. Перезапустите сервер (Евракор) или нажмите кнопку загрузки."}</p>`;
      bindContractorsSeed(el);
      return;
    }

    el.innerHTML = `
      <div class="personnel-toolbar">
        ${storageBadge}
        ${seedBtn}
        <p class="planning-meta">${contractors.length} подрядчиков</p>
      </div>
      <div class="personnel-table-wrap">
        <table class="planning-table personnel-table">
          <thead>
            <tr>
              <th>Код</th>
              <th>Название</th>
              <th>Болванка</th>
              <th>Проектов</th>
            </tr>
          </thead>
          <tbody>
            ${contractors
              .map(
                (c) =>
                  `<tr>
                    <td><code>${esc(c.id)}</code></td>
                    <td>${esc(c.name)}</td>
                    <td>${esc(c.template_stem)}${c.template_exists ? "" : ' <span class="hint-text">(нет файла)</span>'}</td>
                    <td>${c.projects_count ?? 0}</td>
                  </tr>`
              )
              .join("")}
          </tbody>
        </table>
      </div>`;
    bindContractorsSeed(el);
  }

  function bindContractorsSeed(root) {
    const btn = root.querySelector("#contractorsSeedBtn");
    if (!btn) return;
    btn.addEventListener("click", async () => {
      btn.disabled = true;
      try {
        const res = await fetch("/api/planning/contractors/seed-from-templates", { method: "POST" });
        const body = await res.json().catch(() => ({}));
        if (!res.ok) throw new Error(body.detail || res.statusText);
        loaded.contractors = false;
        await loadTab("contractors", true);
      } catch (e) {
        alert(e.message);
      } finally {
        btn.disabled = false;
      }
    });
  }

  function renderProjects(data) {
    const el = document.getElementById("projectsList");
    if (!el) return;
    const db = data.db || {};
    const projects = data.projects || [];
    const contractors = data.contractors || [];
    const engineers = data.engineers || [];
    const dbError = !db.ok ? (db.error || "PostgreSQL недоступна") : "";
    const storageBadge = `<span class="storage-badge storage-badge--db" title="${esc(dbError)}">PostgreSQL · ${projects.length}</span>`;

    const contractorOptions = contractors
      .map((c) => `<option value="${esc(c.id)}">${esc(c.name)}</option>`)
      .join("");

    const filterOptions =
      `<option value="">Все подрядчики</option>` +
      contractors.map((c) => `<option value="${esc(c.id)}">${esc(c.name)}</option>`).join("");

    const createForm =
      contractors.length && db.ok
        ? `<form id="projectCreateForm" class="project-create-form">
            <h3 class="planning-subtitle">Новый проект</h3>
            <div class="project-create-grid">
              <label>Код проекта <input class="field-input" name="project_id" required placeholder="например gkm-2025"/></label>
              <label>Подрядчик
                <select class="field-input" name="contractor_id" required>${contractorOptions}</select>
              </label>
              <label>Название <input class="field-input" name="title" placeholder="кратко"/></label>
              <label>Объект <input class="field-input" name="object_name" placeholder="полное имя объекта"/></label>
            </div>
            <button type="submit" class="btn btn-primary btn-sm">Создать</button>
            <span id="projectCreateStatus" class="hint-text" aria-live="polite"></span>
          </form>`
        : `<p class="hint-text">Сначала заведите подрядчика (раздел «Подрядчики»).</p>`;

    el.innerHTML = `
      <div class="personnel-toolbar">
        ${storageBadge}
        <label class="personnel-filter">Подрядчик
          <select id="projectsContractorFilter" class="field-input">${filterOptions}</select>
        </label>
      </div>
      ${createForm}
      <div id="projectsCards" class="projects-cards"></div>`;

    const filter = el.querySelector("#projectsContractorFilter");
    if (filter && data.contractor_id) filter.value = data.contractor_id;

    async function saveProjectEngineers(projectId, root) {
      const status = root.querySelector(`.project-save-status[data-project="${projectId}"]`);
      const ids = [...root.querySelectorAll(`input[type=checkbox][data-project="${projectId}"]:checked`)].map(
        (x) => x.value
      );
      try {
        const res = await fetch(`/api/planning/projects/${encodeURIComponent(projectId)}/engineers`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ engineer_ids: ids }),
        });
        const body = await res.json().catch(() => ({}));
        if (!res.ok) throw new Error(body.detail || res.statusText);
        if (status) status.textContent = "Сохранено";
      } catch (e) {
        if (status) {
          status.textContent = e.message;
          status.classList.add("error-text");
        }
      }
    }

    function renderCards() {
      const cardsEl = el.querySelector("#projectsCards");
      if (!projects.length) {
        cardsEl.innerHTML = '<p class="hint-text">Проектов пока нет.</p>';
        return;
      }
      cardsEl.innerHTML = projects
        .map((p) => {
          const assigned = new Set(p.engineer_ids || []);
          const checks = engineers.length
            ? engineers
                .map(
                  (eng) =>
                    `<label class="project-engineer-check">
                      <input type="checkbox" data-project="${esc(p.id)}" value="${esc(eng.id)}" ${assigned.has(eng.id) ? "checked" : ""}/>
                      ${esc(eng.fio)}
                    </label>`
                )
                .join("")
            : '<p class="hint-text">Нет инженеров в справочнике сотрудников.</p>';
          return `<article class="project-card" data-project-id="${esc(p.id)}">
            <header class="project-card-header">
              <h3><code>${esc(p.id)}</code> — ${esc(p.object_name || p.title)}</h3>
              <p class="hint-text">Подрядчик: ${esc(p.contractor_name || p.contractor_id)}</p>
            </header>
            <div class="project-engineers">
              <p class="planning-meta">Инженеры на объекте</p>
              ${checks}
              <button type="button" class="btn btn-secondary btn-sm project-save-engineers" data-project="${esc(p.id)}">Сохранить назначения</button>
              <span class="hint-text project-save-status" data-project="${esc(p.id)}"></span>
            </div>
          </article>`;
        })
        .join("");

      cardsEl.querySelectorAll(".project-save-engineers").forEach((btn) => {
        btn.addEventListener("click", async () => saveProjectEngineers(btn.dataset.project, cardsEl));
      });
    }

    const form = el.querySelector("#projectCreateForm");
    if (form) {
      form.addEventListener("submit", async (ev) => {
        ev.preventDefault();
        const status = el.querySelector("#projectCreateStatus");
        const fd = new FormData(form);
        const payload = {
          project_id: String(fd.get("project_id") || "").trim(),
          contractor_id: String(fd.get("contractor_id") || "").trim(),
          title: String(fd.get("title") || "").trim(),
          object_name: String(fd.get("object_name") || "").trim(),
        };
        try {
          const res = await fetch("/api/planning/projects", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
          });
          const body = await res.json().catch(() => ({}));
          if (!res.ok) throw new Error(body.detail || res.statusText);
          form.reset();
          if (status) status.textContent = "Проект создан";
          loaded.projects = false;
          await loadTab("projects", true, filter?.value || "");
        } catch (e) {
          if (status) {
            status.textContent = e.message;
            status.classList.add("error-text");
          }
        }
      });
    }

    if (filter) {
      filter.addEventListener("change", async () => {
        loaded.projects = false;
        await loadTab("projects", true, filter.value || "");
      });
    }

    renderCards();
  }

  async function openOtkkDetail(cardId) {
    const dialog = document.getElementById("otkkDetailDialog");
    const titleEl = document.getElementById("otkkDetailTitle");
    const bodyEl = document.getElementById("otkkDetailBody");
    if (!dialog || !titleEl || !bodyEl) return;
    titleEl.textContent = "Загрузка…";
    bodyEl.innerHTML = "";
    dialog.showModal();
    try {
      const res = await fetch(`/api/planning/otkk/${encodeURIComponent(cardId)}`);
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data.detail || res.statusText);
      const content = data.content || {};
      titleEl.textContent = [content.code, content.title].filter(Boolean).join(" — ") || cardId;
      const rows = (content.rows || []).map(renderOtkkRowBody).join("");
      bodyEl.innerHTML = `<table class="planning-table otkk-structure-table"><tbody>${rows}</tbody></table>`;
    } catch (e) {
      bodyEl.innerHTML = `<p class="error-text">${esc(e.message)}</p>`;
      titleEl.textContent = cardId;
    }
  }

  async function loadTab(name, force, contractorId) {
    if (!force && loaded[name]) return;
    const qs = name === "projects" && contractorId ? `?contractor_id=${encodeURIComponent(contractorId)}` : "";
    const res = await fetch(`/api/planning/${name}${qs}`);
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();
    if (name === "otkk") renderOtkk(data);
    else if (name === "personnel") renderPersonnel(data);
    else if (name === "contractors") renderContractors(data);
    else if (name === "projects") renderProjects(data);
    loaded[name] = true;
  }

  function showLoadError(name, e) {
    const ids = {
      personnel: "personnelList",
      contractors: "contractorsList",
      projects: "projectsList",
      otkk: "otkkList",
    };
    const el = document.getElementById(ids[name]);
    if (el) el.insertAdjacentHTML("beforeend", `<p class="error-text">${esc(e.message)}</p>`);
  }

  if (sectionOnly) {
    loadTab(sectionOnly).catch((e) => showLoadError(sectionOnly, e));
    return;
  }
})();

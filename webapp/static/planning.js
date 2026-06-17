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
    const dbError = !db.ok ? (db.error || "PostgreSQL недоступна") : "";
    const storageBadge = `<span class="storage-badge storage-badge--db" title="${esc(dbError)}">PostgreSQL · ${projects.length}</span>`;

    if (!projects.length) {
      el.innerHTML =
        `<div class="personnel-toolbar">${storageBadge}</div>` +
        `<p class="hint-text">${dbError || "Нет карточек в PostgreSQL. Эталоны заливаются при старте сервера."}</p>` +
        `<dialog id="projectDetailDialog" class="otkk-detail-dialog">
          <div class="otkk-detail-header project-detail-header">
            <button type="button" class="btn btn-secondary btn-sm" id="projectDetailClose">Закрыть</button>
          </div>
          <div id="projectDetailBody" class="otkk-detail-body project-detail-body"></div>
        </dialog>`;
      bindProjectDetail(el);
      return;
    }

    el.innerHTML = `
      <div class="personnel-toolbar">
        ${storageBadge}
        <p class="planning-meta">${projects.length} проектов · ${db.with_content ?? projects.filter((p) => p.has_content).length} с данными</p>
        <input type="search" id="projectsSearch" class="field-input personnel-search" placeholder="Поиск по шифру или объекту…"/>
      </div>
      <div class="personnel-table-wrap projects-table-wrap">
        <table class="planning-table projects-table" id="projectsTable">
          <thead>
            <tr>
              <th>Объект</th>
              <th>Шифр проекта</th>
              <th></th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>
      </div>
      <dialog id="projectDetailDialog" class="otkk-detail-dialog">
        <div class="otkk-detail-header project-detail-header">
          <button type="button" class="btn btn-secondary btn-sm" id="projectDetailClose">Закрыть</button>
        </div>
        <div id="projectDetailBody" class="otkk-detail-body project-detail-body"></div>
      </dialog>`;

    bindProjectDetail(el);

    const tbody = el.querySelector("#projectsTable tbody");
    const search = el.querySelector("#projectsSearch");

    function renderRows() {
      const q = (search?.value || "").trim().toLowerCase();
      const rows = projects.filter((p) => {
        if (!q) return true;
        const hay = [p.id, p.cipher, p.title, p.object_name].join(" ").toLowerCase();
        return hay.includes(q);
      });
      tbody.innerHTML = rows
        .map(
          (p) =>
            `<tr>
              <td>${esc(p.object_name || p.title || "—")}</td>
              <td><code>${esc(p.cipher || p.id)}</code></td>
              <td><button type="button" class="btn btn-link btn-sm project-open-btn" data-id="${esc(p.id)}">Открыть</button></td>
            </tr>`
        )
        .join("");
      if (!rows.length) {
        tbody.innerHTML = '<tr><td colspan="3" class="hint-text">Ничего не найдено</td></tr>';
      }
      tbody.querySelectorAll(".project-open-btn").forEach((btn) => {
        btn.addEventListener("click", () => openProjectDetail(btn.dataset.id));
      });
    }

    search?.addEventListener("input", renderRows);
    renderRows();
  }

  function bindProjectDetail(root) {
    const dialog = root.querySelector("#projectDetailDialog") || document.getElementById("projectDetailDialog");
    const closeBtn = root.querySelector("#projectDetailClose") || document.getElementById("projectDetailClose");
    if (!dialog || !closeBtn) return;
    closeBtn.addEventListener("click", () => dialog.close());
    dialog.addEventListener("click", (e) => {
      if (e.target === dialog) dialog.close();
    });
  }

  function renderVorSection(vor) {
    if (!vor || !vor.stages?.length) {
      return '<p class="hint-text">ВОР не распознан.</p>';
    }
    return (vor.stages || [])
      .map((stage) => {
        const stageLabel = String(stage.title || "").trim();
        const showStageHeading = stageLabel && !/\.docx?$/i.test(stageLabel);
        let html = showStageHeading
          ? `<h4 class="otkk-section-heading">${esc(stageLabel || "Этап")}</h4>`
          : "";
        const objects = stage.objects || [];
        if (!objects.length && (stage.works || []).length) {
          html += renderVorWorksTable(stage.works);
          return html;
        }
        html += objects
          .map((obj) => {
            const works = obj.works || [];
            if (!works.length) {
              return `<p class="planning-meta">${esc(obj.title || "")}</p>`;
            }
            return `<p class="planning-meta">${esc(obj.title || "")}</p>${renderVorWorksTable(works)}`;
          })
          .join("");
        return html;
      })
      .join("");
  }

  function renderVorWorksTable(works) {
    if (!works?.length) return "";
    const rows = works
      .map(
        (w) =>
          `<tr>
            <td>${esc(w.name || "")}</td>
            <td>${esc(w.unit || "")}</td>
            <td>${esc(w.quantity || "")}</td>
            <td>${esc(w.note || "")}</td>
          </tr>`
      )
      .join("");
    return `<div class="otkk-inner-table-wrap"><table class="planning-table otkk-inner-table"><thead><tr><th>Работа</th><th>Ед.</th><th>Объём</th><th>Примечание</th></tr></thead><tbody>${rows}</tbody></table></div>`;
  }

  function renderTlSection(tl) {
    if (tl?.rows?.length) {
      const rows = tl.rows
        .map(
          (row) =>
            `<tr><th class="otkk-label-cell">${esc(row.label || "")}</th><td class="otkk-value-cell">${(row.value || "").split("\n").map(esc).join("<br>")}</td></tr>`
        )
        .join("");
      return `<table class="planning-table otkk-structure-table"><tbody>${rows}</tbody></table>`;
    }
    if (!tl?.tables?.length) {
      return '<p class="hint-text">ТЛ не заполнен.</p>';
    }
    return (tl.tables || [])
      .map((table, idx) => {
        const rows = (table.rows || [])
          .map((cells) => `<tr>${cells.map((c) => `<td>${esc(c)}</td>`).join("")}</tr>`)
          .join("");
        return `<h4 class="otkk-section-heading">Таблица ${idx + 1}</h4><table class="planning-table"><tbody>${rows}</tbody></table>`;
      })
      .join("");
  }

  async function openProjectDetail(projectId) {
    const dialog = document.getElementById("projectDetailDialog");
    const bodyEl = document.getElementById("projectDetailBody");
    if (!dialog || !bodyEl) return;
    bodyEl.innerHTML = '<p class="hint-text">Загрузка…</p>';
    dialog.showModal();
    try {
      const res = await fetch(`/api/planning/projects/${encodeURIComponent(projectId)}`);
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data.detail || res.statusText);
      const content = data.content || {};
      bodyEl.innerHTML = `
        ${renderTlSection(content.tl)}
        <h4 class="otkk-section-heading">Ведомость объёмов работ</h4>
        ${renderVorSection(content.vor)}`;
    } catch (e) {
      bodyEl.innerHTML = `<p class="error-text">${esc(e.message)}</p>`;
    }
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

  async function loadTab(name, force) {
    if (!force && loaded[name]) return;
    const res = await fetch(`/api/planning/${name}`);
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

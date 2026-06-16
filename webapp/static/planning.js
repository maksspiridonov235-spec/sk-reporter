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
              <th>Проекты</th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>
      </div>`;

    bindPersonnelImport(el);

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
    const storageBadge = `<span class="storage-badge storage-badge--db" title="${esc(dbError)}">PostgreSQL${db.count != null ? ` · ${db.count}` : ""}</span>`;

    const importToolbar = `<div class="personnel-import-bar">
        <label class="btn btn-primary btn-sm personnel-upload-label">
          Загрузить карту (.doc / .docx)
          <input type="file" id="otkkUploadDoc" accept=".doc,.docx" hidden/>
        </label>
        <details class="otkk-admin-import">
          <summary class="hint-text">Дополнительно</summary>
          <div class="personnel-import-bar">
            <label class="btn btn-secondary btn-sm personnel-upload-label">
              Импорт manifest.yaml
              <input type="file" id="otkkUploadManifest" accept=".yaml,.yml" hidden/>
            </label>
            <button type="button" class="btn btn-secondary btn-sm" id="otkkScanDisk">Сканировать data/tk/</button>
          </div>
        </details>
        <span id="otkkImportStatus" class="hint-text" aria-live="polite"></span>
      </div>`;

    if (!cards.length) {
      const emptyHint = dbError
        ? `Каталог недоступен: ${esc(dbError)}`
        : "В базе пока нет карт. Загрузите файл ОТКК (.doc) кнопкой выше — содержимое сохранится в PostgreSQL.";
      el.innerHTML =
        `<div class="personnel-toolbar">${storageBadge}</div>` +
        (db.ok ? importToolbar : "") +
        `<p class="hint-text">${emptyHint}</p>`;
      bindOtkkImport(el);
      return;
    }

    const missing = cards.filter((c) => !c.present).length;
    const noContent = cards.filter((c) => !c.has_content).length;
    const metaExtra = [
      missing ? `<span class="warn-text">${missing} без файла на диске</span>` : "",
      noContent ? `<span class="warn-text">${noContent} без текста в БД</span>` : "",
    ]
      .filter(Boolean)
      .join(" · ");

    el.innerHTML = `
      <div class="personnel-toolbar">
        ${storageBadge}
        <p class="planning-meta">${cards.length} карт · на диске: <strong>${data.present_count ?? 0}</strong> · в БД: <strong>${data.content_count ?? 0}</strong>${metaExtra ? ` · ${metaExtra}` : ""}</p>
        <input type="search" id="otkkSearch" class="field-input personnel-search" placeholder="Поиск по ID, коду или названию…"/>
      </div>
      ${importToolbar}
      <p class="hint-text">Повторная загрузка того же номера ОТКК полностью перезаписывает карту в базе.</p>
      <div class="personnel-table-wrap">
        <table class="planning-table personnel-table" id="otkkTable">
          <thead>
            <tr>
              <th>ID</th>
              <th>Код</th>
              <th>Название</th>
              <th>В БД</th>
              <th>На диске</th>
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

    bindOtkkImport(el);
    bindOtkkDetail(el);

    const tbody = el.querySelector("#otkkTable tbody");
    const search = el.querySelector("#otkkSearch");

    function renderRows() {
      const q = (search.value || "").trim().toLowerCase();
      const rows = cards.filter((c) => {
        if (!q) return true;
        const hay = [c.id, c.file, c.code, c.title].join(" ").toLowerCase();
        return hay.includes(q);
      });
      tbody.innerHTML = rows
        .map(
          (c) =>
            `<tr>
              <td><code>${esc(c.id)}</code></td>
              <td>${esc(c.code || "—")}</td>
              <td class="otkk-title-cell">${esc(c.title || c.file || "—")}</td>
              <td>${c.has_content ? "✓" : '<span class="warn-text">нет</span>'}</td>
              <td>${c.present ? "✓" : '<span class="warn-text">нет</span>'}</td>
              <td>${c.has_content ? `<button type="button" class="btn btn-link btn-sm otkk-open-btn" data-id="${esc(c.id)}">Открыть</button>` : ""}</td>
            </tr>`
        )
        .join("");
      if (!rows.length) {
        tbody.innerHTML = '<tr><td colspan="6" class="hint-text">Ничего не найдено</td></tr>';
      }
      tbody.querySelectorAll(".otkk-open-btn").forEach((btn) => {
        btn.addEventListener("click", () => openOtkkDetail(btn.dataset.id));
      });
    }

    search.addEventListener("input", renderRows);
    renderRows();
  }

  function bindOtkkImport(root) {
    const docInput = root.querySelector("#otkkUploadDoc");
    const manifestInput = root.querySelector("#otkkUploadManifest");
    const scanBtn = root.querySelector("#otkkScanDisk");
    const status = root.querySelector("#otkkImportStatus");
    if (!docInput && !manifestInput && !scanBtn) return;

    async function setStatus(msg, isError) {
      if (!status) return;
      status.textContent = msg;
      status.classList.toggle("error-text", !!isError);
    }

    if (docInput) {
      docInput.addEventListener("change", async () => {
        const file = docInput.files?.[0];
        if (!file) return;
        docInput.value = "";
        await setStatus(`Загрузка ${file.name}…`);
        try {
          const fd = new FormData();
          fd.append("file", file);
          const res = await fetch("/api/planning/otkk/upload-doc", {
            method: "POST",
            body: fd,
          });
          const body = await res.json().catch(() => ({}));
          if (!res.ok) throw new Error(body.detail || res.statusText);
          await setStatus(`Сохранено: ${body.id} (${body.rows ?? "?"} строк)`);
          await loadTab("otkk", true);
        } catch (e) {
          await setStatus(e.message, true);
        }
      });
    }

    if (manifestInput) {
      manifestInput.addEventListener("change", async () => {
        const file = manifestInput.files?.[0];
        if (!file) return;
        manifestInput.value = "";
        await setStatus("Импорт manifest…");
        try {
          const fd = new FormData();
          fd.append("file", file);
          const res = await fetch("/api/planning/otkk/import-manifest", {
            method: "POST",
            body: fd,
          });
          const body = await res.json().catch(() => ({}));
          if (!res.ok) throw new Error(body.detail || res.statusText);
          await setStatus(`Импортировано: ${body.upserted ?? "?"}`);
          await loadTab("otkk", true);
        } catch (e) {
          await setStatus(e.message, true);
        }
      });
    }

    if (scanBtn) {
      scanBtn.addEventListener("click", async () => {
        scanBtn.disabled = true;
        await setStatus("Сканирование data/tk/…");
        try {
          const res = await fetch("/api/planning/otkk/scan-disk", { method: "POST" });
          const body = await res.json().catch(() => ({}));
          if (!res.ok) throw new Error(body.detail || res.statusText);
          await setStatus(`В каталоге: ${body.upserted ?? "?"}`);
          await loadTab("otkk", true);
        } catch (e) {
          await setStatus(e.message, true);
        } finally {
          scanBtn.disabled = false;
        }
      });
    }
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

  function renderOtkkRowBody(row) {
    const label = esc(row.label || "");
    let valueHtml = esc(row.value || "");
    const body = row.body;
    if (body) {
      const parts = [];
      for (const p of body.paragraphs || []) {
        parts.push(`<p>${esc(p)}</p>`);
      }
      if (body.bullets?.length) {
        parts.push(
          `<ul class="otkk-bullets">${body.bullets.map((b) => `<li>${esc(b)}</li>`).join("")}</ul>`
        );
      }
      if (parts.length) valueHtml = parts.join("");
    }
    if (row.codes?.length) {
      valueHtml += `<p class="hint-text">Коды: ${row.codes.map((c) => `<code>${esc(c)}</code>`).join(", ")}</p>`;
    }
    return `<tr><th scope="row">${label}</th><td>${valueHtml}</td></tr>`;
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
      const sig = content.signature;
      const sigHtml = sig?.text
        ? `<p class="otkk-signature"><strong>${esc(sig.label || "Подпись")}:</strong> ${esc(sig.text)}</p>`
        : "";
      bodyEl.innerHTML = `<table class="planning-table otkk-structure-table"><tbody>${rows}</tbody></table>${sigHtml}`;
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
    if (name === "projects") renderProjects(data);
    else if (name === "otkk") renderOtkk(data);
    else if (name === "personnel") renderPersonnel(data);
    loaded[name] = true;
  }

  function showLoadError(name, e) {
    const ids = {
      projects: "projectsList",
      personnel: "personnelList",
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

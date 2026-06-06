(function () {
  const projectSelect = document.getElementById("projectSelect");
  const reportDate = document.getElementById("reportDate");
  const buildBtn = document.getElementById("buildBtn");
  const worksBody = document.getElementById("worksBody");
  const workFilter = document.getElementById("workFilter");
  const selectAll = document.getElementById("selectAll");
  const statusMsg = document.getElementById("statusMsg");
  const templateWarn = document.getElementById("templateWarn");
  const profileName = document.getElementById("profileName");
  const workStats = document.getElementById("workStats");

  let config = null;
  let currentWorks = [];

  function todayIso() {
    const d = new Date();
    return d.toISOString().slice(0, 10);
  }

  function setStatus(text, isError) {
    statusMsg.textContent = text || "";
    statusMsg.classList.toggle("error-text", !!isError);
  }

  function projectById(id) {
    return (config?.projects || []).find((p) => p.id === id);
  }

  function renderWorks(projectId) {
    const proj = projectById(projectId);
    currentWorks = proj?.works || [];
    const q = (workFilter.value || "").trim().toLowerCase();
    worksBody.innerHTML = "";
    let shown = 0;
    currentWorks.forEach((w) => {
      const hay = `${w.object} ${w.name} ${w.stage}`.toLowerCase();
      if (q && !hay.includes(q)) return;
      shown += 1;
      const tr = document.createElement("tr");
      tr.dataset.key = w.key;
      tr.innerHTML = `
        <td><input type="checkbox" class="row-check"/></td>
        <td class="cell-object" title="${esc(w.object)}">${esc(w.object || "—")}</td>
        <td class="cell-name" title="${esc(w.name)}">${esc(w.name)}</td>
        <td>${esc(w.unit)}</td>
        <td>${esc(w.quantity)}</td>
        <td><input type="text" class="cell-daily field-input field-input--sm" placeholder="0"/></td>
        <td><input type="text" class="cell-cum field-input field-input--sm" placeholder=""/></td>
        <td><input type="text" class="cell-loc field-input field-input--sm" placeholder=""/></td>
        <td><input type="text" class="cell-ref field-input field-input--sm" placeholder=""/></td>
        <td class="cell-tk">${w.tk_id ? esc(w.tk_id) : "—"}</td>
      `;
      worksBody.appendChild(tr);
    });
    workStats.textContent = shown
      ? `Показано ${shown} из ${currentWorks.length}`
      : "Нет работ (обновите vor.json)";
    selectAll.checked = false;
  }

  function esc(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/"/g, "&quot;");
  }

  function collectEntries() {
    const rows = worksBody.querySelectorAll("tr");
    const entries = [];
    rows.forEach((tr) => {
      const cb = tr.querySelector(".row-check");
      if (!cb?.checked) return;
      const key = tr.dataset.key;
      const w = currentWorks.find((x) => x.key === key);
      if (!w) return;
      entries.push({
        key,
        name: w.name,
        unit: w.unit,
        project_qty: w.quantity,
        daily_qty: tr.querySelector(".cell-daily")?.value?.trim() || "",
        cumulative_qty: tr.querySelector(".cell-cum")?.value?.trim() || "",
        location: tr.querySelector(".cell-loc")?.value?.trim() || "",
        reference: tr.querySelector(".cell-ref")?.value?.trim() || "",
        stage: w.stage,
        object: w.object,
      });
    });
    return entries;
  }

  async function loadConfig() {
    setStatus("Загрузка…");
    const res = await fetch("/api/engineer/config");
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail || res.statusText);
    }
    config = await res.json();
    profileName.textContent = config.profile?.name || config.profile?.id || "";
    templateWarn.hidden = config.template_ok;
    templateWarn.textContent = config.template_ok
      ? ""
      : "Шаблон отчёта не найден — укажите report_template в профиле инженера.";

    projectSelect.innerHTML = "";
    (config.projects || []).forEach((p) => {
      const opt = document.createElement("option");
      opt.value = p.id;
      opt.textContent = `${p.title} (${p.works_count})`;
      projectSelect.appendChild(opt);
    });
    if (projectSelect.options.length) {
      renderWorks(projectSelect.value);
      buildBtn.disabled = !config.template_ok;
      setStatus("");
    } else {
      setStatus("Нет проектов в профиле", true);
    }
  }

  async function buildReport() {
    const entries = collectEntries();
    if (!entries.length) {
      setStatus("Отметьте работы и укажите объём за сутки", true);
      return;
    }
    buildBtn.disabled = true;
    setStatus("Формирование…");
    try {
      const res = await fetch("/api/engineer/build", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          project_id: projectSelect.value,
          report_date: reportDate.value || todayIso(),
          entries,
        }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.detail || res.statusText);
      }
      const blob = await res.blob();
      const cd = res.headers.get("Content-Disposition") || "";
      const m = cd.match(/filename=\"?([^\";]+)/);
      const filename = m ? m[1] : "report.docx";
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
      setStatus("Готово — файл скачан");
    } catch (e) {
      setStatus(e.message || "Ошибка", true);
    } finally {
      buildBtn.disabled = !config?.template_ok;
    }
  }

  projectSelect.addEventListener("change", () => renderWorks(projectSelect.value));
  workFilter.addEventListener("input", () => renderWorks(projectSelect.value));
  selectAll.addEventListener("change", () => {
    worksBody.querySelectorAll(".row-check").forEach((cb) => {
      cb.checked = selectAll.checked;
    });
  });
  buildBtn.addEventListener("click", buildReport);

  reportDate.value = todayIso();
  loadConfig().catch((e) => setStatus(e.message, true));
})();

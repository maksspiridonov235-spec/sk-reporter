(function () {
  const el = document.getElementById("luvrList");

  function esc(s) {
    return String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/"/g, "&quot;");
  }

  function renderLuvr(data) {
    if (!data.xlsx_present) {
      el.innerHTML = '<p class="hint-text">Положите <code>ЛУВР.xlsx</code> в data/luvr/</p>';
      return;
    }
    if (!data.cache_ready) {
      el.innerHTML = `<p class="planning-meta">Файл: <strong>${esc(data.source)}</strong> (${data.source_kb} КБ)</p>
        <p class="warn-text">Кэш не собран. Запустите: <code>python scripts/build_engineer_data.py --luvr</code></p>`;
      return;
    }

    const months = data.months || [];
    const defaultSheet = data.default_month || months[months.length - 1]?.sheet || "";

    el.innerHTML = `
      <div class="luvr-header">
        <p class="planning-meta">Файл: <strong>${esc(data.source)}</strong> · ${data.source_kb} КБ · <code>${esc(data.folder)}/${esc(data.source)}</code></p>
        ${data.contract ? `<p class="planning-meta">Договор: ${esc(data.contract)}</p>` : ""}
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
  }

  fetch("/api/luvr")
    .then((r) => {
      if (!r.ok) throw new Error(r.statusText);
      return r.json();
    })
    .then(renderLuvr)
    .catch((e) => {
      el.innerHTML = `<p class="error-text">${esc(e.message)}</p>`;
    });
})();

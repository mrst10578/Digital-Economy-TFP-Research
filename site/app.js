const $ = (id) => document.getElementById(id);

const fmt = (value) => {
  if (value === null || value === undefined || value === "") return "—";
  if (typeof value === "number") return new Intl.NumberFormat("fa-IR").format(value);
  return String(value);
};

const setText = (id, value) => {
  const el = $(id);
  if (el) el.textContent = fmt(value);
};

const sumObjectValues = (obj) => {
  if (!obj || typeof obj !== "object") return null;
  return Object.values(obj).reduce((sum, value) => sum + (Number(value) || 0), 0);
};

const renderVariables = (columns = []) => {
  const box = $("expectedColumns");
  if (!box) return;
  box.innerHTML = "";
  if (!columns.length) {
    box.innerHTML = "<span>در انتظار تعریف نهایی متغیرها</span>";
    return;
  }
  columns.forEach((name) => {
    const chip = document.createElement("span");
    chip.textContent = name;
    box.appendChild(chip);
  });
};

const renderCoefficients = (rows) => {
  if (!Array.isArray(rows) || !rows.length) return;
  $("coefficientsState")?.classList.add("hidden");
  const wrap = $("coefficientsTable");
  wrap?.classList.remove("hidden");

  const head = `
    <table>
      <thead>
        <tr>
          <th>Variable</th>
          <th>Coefficient</th>
          <th>Std. Error</th>
          <th>p-value</th>
          <th>95% CI</th>
        </tr>
      </thead>
      <tbody>
  `;

  const body = rows.map((row) => `
    <tr>
      <td>${row.variable ?? "—"}</td>
      <td>${row.coefficient ?? "—"}</td>
      <td>${row.std_error ?? "—"}</td>
      <td>${row.p_value ?? "—"}</td>
      <td>${row.ci95 ?? "—"}</td>
    </tr>
  `).join("");

  wrap.innerHTML = head + body + "</tbody></table>";
};

const renderDiagnostics = (items) => {
  if (!Array.isArray(items) || !items.length) return;
  $("diagnosticsState")?.classList.add("hidden");
  const box = $("diagnosticsList");
  box?.classList.remove("hidden");
  box.innerHTML = items.map((item) => `
    <div class="diagnostic-row">
      <span>${item.name ?? "Test"}</span>
      <b>${item.result ?? "—"}</b>
    </div>
  `).join("");
};

fetch("./data/project.json", { cache: "no-store" })
  .then((response) => {
    if (!response.ok) throw new Error("dashboard data unavailable");
    return response.json();
  })
  .then((data) => {
    setText("headerStage", data.stage?.short ?? "پروژه پژوهشی");
    setText("stageTitle", data.stage?.title ?? "در حال توسعه");
    setText("stageDescription", data.stage?.description ?? "");
    setText("lastUpdated", data.generated_at_local ?? data.generated_at ?? "—");
    setText("buildStamp", "آخرین build: " + (data.generated_at_local ?? data.generated_at ?? "—"));

    const audit = data.audit ?? {};
    setText("metricRows", audit.n_rows);
    setText("metricEntities", audit.n_entities);
    setText("metricPeriod", data.period_label);
    setText("metricMissing", sumObjectValues(audit.missing_values));

    setText("auditColumns", audit.n_columns);
    setText("auditDuplicates", audit.duplicate_id_time);
    setText("auditMinPeriods", audit.min_periods_per_entity);
    setText("auditMaxPeriods", audit.max_periods_per_entity);

    renderVariables(data.expected_columns ?? []);
    renderCoefficients(data.results?.coefficients ?? []);
    renderDiagnostics(data.results?.diagnostics ?? []);

    if (audit.n_rows) {
      const note = $("dataNote");
      if (note) {
        note.textContent = "این بخش از آخرین خروجی واقعی pipeline خوانده شده است.";
      }
    }

    if (data.model?.specified) {
      setText("modelTitle", data.model.name ?? "مدل مشخص شده است");
      setText("modelMessage", data.model.description ?? "");
      const badge = $("modelBadge");
      if (badge) {
        badge.textContent = "SPECIFIED";
        badge.classList.remove("waiting");
        badge.classList.add("ready");
      }
    }

    if (Array.isArray(data.results?.coefficients) && data.results.coefficients.length) {
      const badge = $("resultsBadge");
      if (badge) {
        badge.textContent = "ESTIMATED";
        badge.classList.remove("neutral");
        badge.classList.add("ready");
      }
    }
  })
  .catch(() => {
    setText("headerStage", "Research Portal");
    setText("lastUpdated", "نامشخص");
    setText("buildStamp", "Dashboard data unavailable");
    renderVariables([]);
  });

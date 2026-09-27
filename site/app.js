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

  const body = rows.map((row) => `
    <tr>
      <td>${row.variable ?? "—"}</td>
      <td>${row.coefficient ?? "—"}</td>
      <td>${row.std_error ?? "—"}</td>
      <td>${row.p_value ?? "—"}</td>
      <td>${row.ci95 ?? "—"}</td>
    </tr>
  `).join("");

  wrap.innerHTML = `
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
      <tbody>${body}</tbody>
    </table>
  `;
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
      if (note) note.textContent = "این بخش از آخرین خروجی واقعی Python خوانده شده است.";
    }

    if (data.model?.specified) {
      setText("modelTitle", data.model.name ?? "مدل مشخص شده");
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
    setText("headerStage", "در انتظار مدل نهایی");
    setText("lastUpdated", "نامشخص");
    setText("buildStamp", "وضعیت build در دسترس نیست");
    renderVariables([]);
  });

const parseCSV = (text) => {
  const lines = text.trim().split(/\r?\n/);
  const headers = lines.shift().split(",");
  return lines.map((line) => {
    const values = line.split(",");
    return Object.fromEntries(headers.map((h, i) => [h, values[i]]));
  });
};

const renderDemoTable = (rows) => {
  const tbody = document.querySelector("#demoTable tbody");
  if (!tbody) return;
  tbody.innerHTML = rows.slice(0, 10).map((row) => `
    <tr>
      <td>${row.country}</td>
      <td>${row.year}</td>
      <td>${row.digital_economy}</td>
      <td>${row.ai}</td>
      <td>${row.human_capital}</td>
      <td>${row.trade_openness}</td>
      <td>${row.tfp}</td>
    </tr>
  `).join("");
};

const renderDemoChart = (rows, country) => {
  const svg = $("demoChart");
  if (!svg) return;

  const series = rows
    .filter((row) => row.country === country)
    .map((row) => ({ year: Number(row.year), tfp: Number(row.tfp) }))
    .sort((a, b) => a.year - b.year);

  if (!series.length) return;

  const w = 760, h = 300;
  const m = { l: 55, r: 22, t: 24, b: 42 };
  const innerW = w - m.l - m.r;
  const innerH = h - m.t - m.b;

  const minY = Math.min(...series.map(d => d.tfp));
  const maxY = Math.max(...series.map(d => d.tfp));
  const pad = Math.max((maxY - minY) * 0.25, 0.015);
  const y0 = minY - pad;
  const y1 = maxY + pad;

  const x = (i) => m.l + (i / Math.max(series.length - 1, 1)) * innerW;
  const y = (v) => m.t + (1 - ((v - y0) / (y1 - y0))) * innerH;

  const grid = [0, .25, .5, .75, 1].map((t) => {
    const yy = m.t + t * innerH;
    const val = (y1 - t * (y1 - y0)).toFixed(3);
    return `
      <line class="grid-line" x1="${m.l}" y1="${yy}" x2="${w - m.r}" y2="${yy}"></line>
      <text x="${m.l - 10}" y="${yy + 4}" text-anchor="end">${val}</text>
    `;
  }).join("");

  const points = series.map((d, i) => `${x(i)},${y(d.tfp)}`).join(" ");
  const dots = series.map((d, i) => `
    <circle class="series-dot" cx="${x(i)}" cy="${y(d.tfp)}" r="4"></circle>
    <text x="${x(i)}" y="${h - 15}" text-anchor="middle">${d.year}</text>
  `).join("");

  svg.innerHTML = `
    ${grid}
    <line class="axis-line" x1="${m.l}" y1="${m.t}" x2="${m.l}" y2="${h - m.b}"></line>
    <line class="axis-line" x1="${m.l}" y1="${h - m.b}" x2="${w - m.r}" y2="${h - m.b}"></line>
    <polyline class="series-line" points="${points}"></polyline>
    ${dots}
  `;
};

fetch("./data/demo_synthetic.csv", { cache: "no-store" })
  .then((response) => {
    if (!response.ok) throw new Error("demo data unavailable");
    return response.text();
  })
  .then((text) => {
    const rows = parseCSV(text);
    renderDemoTable(rows);

    const countries = [...new Set(rows.map((row) => row.country))];
    const select = $("demoCountry");
    if (!select) return;

    select.innerHTML = countries.map((country) => `<option value="${country}">${country}</option>`).join("");
    const first = countries[0];
    renderDemoChart(rows, first);
    select.addEventListener("change", () => renderDemoChart(rows, select.value));
  })
  .catch(() => {
    const select = $("demoCountry");
    if (select) select.innerHTML = "<option>Demo unavailable</option>";
  });

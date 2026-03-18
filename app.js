
(() => {
  "use strict";

  const state = {
    workbookName: "",
    sheetName: "",
    rawRows: [],
    rows: [],
    grouped: [],
    messages: [],
    controlName: "",
    primaryTreatment: "",
    secondaryTreatment: "",
    analysis: null,
    sensitivity: []
  };

  const aliases = {
    treatment_name: ["amendment name", "treatment name", "treatment_name", "amendment_name"],
    practice_label: ["practice change label", "practice_change_label", "practice label"],
    treatment_id: ["treatment id", "treatment_id"],
    replicate_id: ["replicate id", "replicate_id"],
    yield_t_ha: ["crop_yield_t_ha", "crop yield t ha", "crop_yield_t_ha ", "yield_t_ha", "yield"],
    total_cost_per_ha: ["total farm costs_per_ha", "total farm costs per ha", "total farm costs", "total farm costs/ha", "total_cost_per_ha", "total farm costs ha"]
  };

  const $ = (id) => document.getElementById(id);
  const tabs = Array.from(document.querySelectorAll(".tab-btn"));

  function init() {
    bindTabs();
    bindEvents();
    updateDiscountExplanation();
    loadDefaultWorkbook(true);
  }

  function bindTabs() {
    tabs.forEach(btn => btn.addEventListener("click", () => activateTab(btn.dataset.tab)));
  }

  function activateTab(name) {
    tabs.forEach(btn => btn.classList.toggle("active", btn.dataset.tab === name));
    document.querySelectorAll(".tab-panel").forEach(panel => panel.classList.toggle("active", panel.id === `tab-${name}`));
  }

  function bindEvents() {
    $("fileInput").addEventListener("change", onFileSelected);
    $("useDefaultBtn").addEventListener("click", () => { loadDefaultWorkbook(false); activateTab("data"); });
    $("loadDefaultDataBtn").addEventListener("click", () => loadDefaultWorkbook(false));
    $("goToDataBtn").addEventListener("click", () => activateTab("data"));
    $("goToResultsBtn").addEventListener("click", () => activateTab("results"));
    $("clearDataBtn").addEventListener("click", clearData);
    $("validateDataBtn").addEventListener("click", validateOnly);
    $("runAnalysisBtn").addEventListener("click", () => { runAnalysis(); activateTab("results"); });
    $("runAnalysisBtnTop").addEventListener("click", () => { runAnalysis(); activateTab("results"); });
    $("runAnalysisBtnAssumptions").addEventListener("click", () => { runAnalysis(); activateTab("results"); });
    $("runSensitivityBtn").addEventListener("click", () => { if (!state.analysis) runAnalysis(); runSensitivity(); });
    $("downloadReportBtn").addEventListener("click", downloadReport);
    $("printReportBtn").addEventListener("click", printReport);
    $("resetAssumptionsBtn").addEventListener("click", resetAssumptions);
    $("discountMode").addEventListener("change", updateDiscountExplanation);
    $("discountInitial").addEventListener("input", updateDiscountExplanation);
    $("discountLater").addEventListener("input", updateDiscountExplanation);
    $("discountSwitch").addEventListener("input", updateDiscountExplanation);
    $("comparisonMode").addEventListener("change", renderResults);
    $("controlSelect").addEventListener("change", () => { state.controlName = $("controlSelect").value; if (state.analysis) runAnalysis(); });
    $("primaryTreatmentSelect").addEventListener("change", () => { state.primaryTreatment = $("primaryTreatmentSelect").value; if (state.analysis) renderResults(); });
    $("secondaryTreatmentSelect").addEventListener("change", () => { state.secondaryTreatment = $("secondaryTreatmentSelect").value; if (state.analysis) renderResults(); });
  }

  function resetAssumptions() {
    $("grainPriceInput").value = 500;
    $("yearsInput").value = 10;
    $("discountMode").value = "constant";
    $("discountInitial").value = 5;
    $("discountLater").value = 3;
    $("discountSwitch").value = 5;
    updateDiscountExplanation();
  }

  function updateStatus(title, text, type = "neutral") {
    $("statusTitle").textContent = title;
    $("statusText").textContent = text;
    $("statusBanner").className = `status-banner ${type}`;
    $("summaryFile").textContent = state.workbookName || "No workbook loaded";
    $("summaryRows").textContent = `Rows: ${state.rows.length || 0}`;
    $("summaryTreatments").textContent = `Treatments: ${state.grouped.length || 0}`;
  }

  function normaliseHeader(value) {
    return String(value || "").trim().toLowerCase().replace(/[%()]/g, "").replace(/[^a-z0-9]+/g, " ").trim();
  }

  function mapHeaders(headers) {
    const mapped = {};
    headers.forEach((header, idx) => {
      const norm = normaliseHeader(header);
      Object.entries(aliases).forEach(([key, options]) => {
        if (!(key in mapped) && options.includes(norm)) mapped[key] = idx;
      });
    });
    return mapped;
  }

  function toNumber(value) {
    if (value === null || value === undefined || value === "") return NaN;
    const cleaned = String(value).replace(/[$,%\s,]/g, "");
    const num = Number(cleaned);
    return Number.isFinite(num) ? num : NaN;
  }

  function detectControlName(rows) {
    const hit = rows.find(row => {
      const a = String(row.treatment_name || "").toLowerCase();
      const b = String(row.practice_label || "").toLowerCase();
      const c = String(row.treatment_id || "").toLowerCase();
      return a === "control" || a.includes("control") || b.includes("no change") || c === "t00";
    });
    return hit ? hit.treatment_name : "";
  }

  function cleanRows(sourceRows) {
    const headers = sourceRows[0] || [];
    const records = sourceRows.slice(1);
    const mapped = mapHeaders(headers);
    const required = ["treatment_name", "yield_t_ha", "total_cost_per_ha"];
    const missing = required.filter(k => !(k in mapped));
    const messages = [];

    if (missing.length) {
      return { rows: [], messages: [{ type: "error", text: `Required columns are missing: ${missing.join(", ")}.` }] };
    }

    const cleaned = [];
    let skipped = 0;
    records.forEach((record, idx) => {
      const row = {
        treatment_id: mapped.treatment_id !== undefined ? record[mapped.treatment_id] : "",
        replicate_id: mapped.replicate_id !== undefined ? record[mapped.replicate_id] : "",
        treatment_name: String(record[mapped.treatment_name] || "").trim(),
        practice_label: mapped.practice_label !== undefined ? String(record[mapped.practice_label] || "").trim() : "",
        yield_t_ha: toNumber(record[mapped.yield_t_ha]),
        total_cost_per_ha: toNumber(record[mapped.total_cost_per_ha]),
        source_index: idx + 2
      };
      if (row.treatment_name && Number.isFinite(row.yield_t_ha) && Number.isFinite(row.total_cost_per_ha)) cleaned.push(row);
      else skipped += 1;
    });

    if (cleaned.length) {
      messages.push({ type: "success", text: `Workbook loaded successfully. ${cleaned.length} valid rows were found.` });
      if (skipped) messages.push({ type: "warning", text: `${skipped} rows were skipped because a treatment name, yield value, or total cost was missing.` });
    } else {
      messages.push({ type: "error", text: "The workbook was opened, but no valid data rows were found." });
    }

    if (!detectControlName(cleaned)) messages.push({ type: "warning", text: "A control treatment was not detected automatically. Please choose the control manually." });

    return { rows: cleaned, messages };
  }

  function groupRows(rows) {
    const map = new Map();
    rows.forEach(row => {
      const key = row.treatment_name;
      if (!map.has(key)) map.set(key, { treatment_name: key, practice_label: row.practice_label || "", replicates: 0, yields: [], costs: [] });
      const item = map.get(key);
      item.replicates += 1;
      item.yields.push(row.yield_t_ha);
      item.costs.push(row.total_cost_per_ha);
    });
    return Array.from(map.values()).map(item => ({
      treatment_name: item.treatment_name,
      practice_label: item.practice_label,
      replicates: item.replicates,
      mean_yield_t_ha: average(item.yields),
      mean_cost_per_ha: average(item.costs)
    })).sort((a,b) => a.treatment_name.localeCompare(b.treatment_name));
  }

  function average(values) {
    return values.length ? values.reduce((a,b) => a + b, 0) / values.length : NaN;
  }

  function renderValidationMessages(messages) {
    const box = $("validationMessages");
    box.innerHTML = "";
    if (!messages.length) {
      box.innerHTML = `<div class="message">No validation messages yet.</div>`;
      return;
    }
    messages.forEach(item => {
      const div = document.createElement("div");
      div.className = `message ${item.type || ""}`;
      div.textContent = item.text;
      box.appendChild(div);
    });
  }

  function renderSummary() {
    $("fileNameValue").textContent = state.workbookName || "-";
    $("sheetNameValue").textContent = state.sheetName || "-";
    $("rowsReadValue").textContent = state.rawRows.length ? String(state.rawRows.length - 1) : "-";
    $("validRowsValue").textContent = state.rows.length ? String(state.rows.length) : "-";
    $("treatmentsValue").textContent = state.grouped.length ? String(state.grouped.length) : "-";
    $("controlValue").textContent = state.controlName || "Choose manually if needed";
  }

  function renderPreview() {
    const table = $("previewTable");
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");
    thead.innerHTML = "";
    tbody.innerHTML = "";
    if (!state.rawRows.length) {
      thead.innerHTML = "<tr><th>No preview available</th></tr>";
      return;
    }
    const preview = state.rawRows.slice(0, 7);
    const headerTr = document.createElement("tr");
    preview[0].forEach(cell => {
      const th = document.createElement("th");
      th.textContent = String(cell ?? "");
      headerTr.appendChild(th);
    });
    thead.appendChild(headerTr);
    preview.slice(1).forEach(row => {
      const tr = document.createElement("tr");
      preview[0].forEach((_, idx) => {
        const td = document.createElement("td");
        td.textContent = String(row[idx] ?? "");
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
  }

  function renderSelectors() {
    const names = state.grouped.map(x => x.treatment_name);
    fillSelect($("controlSelect"), names, state.controlName);
    fillSelect($("primaryTreatmentSelect"), names, state.primaryTreatment);
    fillSelect($("secondaryTreatmentSelect"), names, state.secondaryTreatment);
  }

  function fillSelect(select, names, selected) {
    select.innerHTML = "";
    if (!names.length) {
      select.innerHTML = `<option value="">No treatments available</option>`;
      return;
    }
    names.forEach(name => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      select.appendChild(option);
    });
    if (selected && names.includes(selected)) select.value = selected;
  }

  function loadDefaultWorkbook(initialLoad) {
    const workbook = window.DEFAULT_BCA_WORKBOOK;
    if (!workbook || !Array.isArray(workbook.data)) {
      updateStatus("Default workbook missing", "The bundled default data could not be found.", "error");
      renderValidationMessages([{ type: "error", text: "Default workbook missing." }]);
      return;
    }
    applyParsedData(workbook.data, "provided_trial_workbook.xlsx", workbook.sheetName || "Sheet1");
    if (initialLoad) runAnalysis();
  }

  function clearData() {
    state.workbookName = "";
    state.sheetName = "";
    state.rawRows = [];
    state.rows = [];
    state.grouped = [];
    state.messages = [];
    state.controlName = "";
    state.primaryTreatment = "";
    state.secondaryTreatment = "";
    state.analysis = null;
    state.sensitivity = [];
    renderValidationMessages([]);
    renderSummary();
    renderPreview();
    renderSelectors();
    $("kpiGrid").innerHTML = "";
    $("comparisonSummary").innerHTML = "";
    $("rankingTable").querySelector("thead").innerHTML = "";
    $("rankingTable").querySelector("tbody").innerHTML = "";
    $("sensitivityTable").querySelector("thead").innerHTML = "";
    $("sensitivityTable").querySelector("tbody").innerHTML = "";
    updateStatus("Data cleared", "The current workbook has been removed. Reload the default data or upload a filled workbook.", "warning");
  }

  function validateOnly() {
    if (!state.rows.length) {
      updateStatus("No workbook loaded", "Load the default workbook or upload a filled workbook before validating.", "warning");
      return;
    }
    renderValidationMessages(state.messages);
    updateStatus("Validation complete", "Review the validation messages and workbook summary.", "success");
  }

  async function onFileSelected(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    try {
      const ext = file.name.toLowerCase().split(".").pop();
      let rows, sheetName = "Uploaded data";
      if (ext === "csv" || ext === "tsv") {
        const text = await file.text();
        rows = parseDelimited(text, ext === "tsv" ? "\t" : ",");
      } else {
        if (!window.XLSX) throw new Error("Excel reader library not available. Check your internet connection and refresh the page.");
        const buffer = await file.arrayBuffer();
        const workbook = window.XLSX.read(buffer, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        sheetName = firstSheetName;
        rows = window.XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { header: 1, raw: true, defval: "" });
      }
      applyParsedData(rows, file.name, sheetName);
    } catch (error) {
      const message = error && error.message ? error.message : "Could not read the uploaded file.";
      renderValidationMessages([{ type: "error", text: message }]);
      updateStatus("Workbook could not be read", message, "error");
    }
  }

  function parseDelimited(text, delimiter) {
    return text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n").filter(Boolean).map(line => line.split(delimiter));
  }

  function applyParsedData(rows, fileName, sheetName) {
    if (!Array.isArray(rows) || !rows.length || !Array.isArray(rows[0])) {
      renderValidationMessages([{ type: "error", text: "The workbook was opened, but no data rows were found." }]);
      updateStatus("Workbook could not be read", "The workbook was opened, but no data rows were found.", "error");
      return;
    }
    state.workbookName = fileName;
    state.sheetName = sheetName || "Sheet1";
    state.rawRows = rows;
    const cleaned = cleanRows(rows);
    state.rows = cleaned.rows;
    state.messages = cleaned.messages;
    state.grouped = groupRows(state.rows);
    state.controlName = detectControlName(state.rows) || state.grouped[0]?.treatment_name || "";
    state.primaryTreatment = state.grouped.find(x => x.treatment_name !== state.controlName)?.treatment_name || state.grouped[0]?.treatment_name || "";
    state.secondaryTreatment = state.grouped.find(x => x.treatment_name !== state.primaryTreatment)?.treatment_name || state.grouped[0]?.treatment_name || "";
    state.analysis = null;
    state.sensitivity = [];
    renderValidationMessages(state.messages);
    renderSummary();
    renderPreview();
    renderSelectors();
    if (state.rows.length) updateStatus("Workbook loaded", "The workbook was read successfully. You can now run the analysis.", "success");
    else updateStatus("Check workbook", "The workbook was opened, but valid data rows were not found.", "error");
  }

  function getAssumptions() {
    return {
      grainPrice: Number($("grainPriceInput").value || 0),
      years: Math.max(1, Number($("yearsInput").value || 1)),
      mode: $("discountMode").value,
      initialRate: Math.max(0, Number($("discountInitial").value || 0)) / 100,
      laterRate: Math.max(0, Number($("discountLater").value || 0)) / 100,
      switchYear: Math.max(1, Number($("discountSwitch").value || 1))
    };
  }

  function updateDiscountExplanation() {
    const a = getAssumptions();
    let text;
    if (a.mode === "constant") text = `A constant annual discount rate of ${(a.initialRate * 100).toFixed(1)}% is applied across the ${a.years}-year analysis period.`;
    else if (a.mode === "declining") text = `A discount rate of ${(a.initialRate * 100).toFixed(1)}% is used through year ${a.switchYear}, then ${(a.laterRate * 100).toFixed(1)}% is used from the following year onward.`;
    else text = `A discount rate of ${(a.initialRate * 100).toFixed(1)}% is used through year ${a.switchYear}, then ${(a.laterRate * 100).toFixed(1)}% is used from the following year onward.`;
    $("discountExplanation").textContent = text;
  }

  function discountFactor(a) {
    let factor = 0;
    let cumulative = 1;
    for (let year = 1; year <= a.years; year += 1) {
      let rate = a.initialRate;
      if (a.mode !== "constant" && year > a.switchYear) rate = a.laterRate;
      cumulative *= (1 + rate);
      factor += 1 / cumulative;
    }
    return factor;
  }

  function analyseGroup(group, assumptions, tweak = {}) {
    const grainPrice = tweak.grainPrice !== undefined ? tweak.grainPrice : assumptions.grainPrice;
    const benefitAdj = 1 + ((tweak.benefitPct || 0) / 100);
    const costAdj = 1 + ((tweak.costPct || 0) / 100);
    const df = discountFactor(assumptions);
    const pvBenefits = group.mean_yield_t_ha * grainPrice * benefitAdj * df;
    const pvCosts = group.mean_cost_per_ha * costAdj * df;
    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts !== 0 ? pvBenefits / pvCosts : NaN;
    const grossProfitMargin = pvBenefits !== 0 ? ((pvBenefits - pvCosts) / pvBenefits) * 100 : NaN;
    return { ...group, pvBenefits, pvCosts, npv, bcr, grossProfitMargin, discount_factor: df };
  }

  function runAnalysis() {
    if (!state.grouped.length) {
      updateStatus("No data loaded", "Load the default workbook or upload a filled workbook first.", "warning");
      activateTab("data");
      return;
    }
    const a = getAssumptions();
    if (!Number.isFinite(a.grainPrice) || a.grainPrice <= 0) {
      updateStatus("Check assumptions", "Grain price must be a positive number.", "warning");
      activateTab("assumptions");
      return;
    }
    const treatments = state.grouped.map(group => analyseGroup(group, a)).sort((x, y) => y.npv - x.npv);
    state.analysis = { assumptions: a, treatments, byName: Object.fromEntries(treatments.map(x => [x.treatment_name, x])) };
    if (!state.analysis.byName[state.controlName]) state.controlName = treatments[0]?.treatment_name || "";
    if (!state.analysis.byName[state.primaryTreatment] || state.primaryTreatment === state.controlName) state.primaryTreatment = treatments.find(x => x.treatment_name !== state.controlName)?.treatment_name || treatments[0]?.treatment_name || "";
    if (!state.analysis.byName[state.secondaryTreatment]) state.secondaryTreatment = treatments.find(x => x.treatment_name !== state.primaryTreatment)?.treatment_name || treatments[0]?.treatment_name || "";
    renderSelectors();
    renderResults();
    updateStatus("Analysis complete", "Results are ready. Review the comparison, ranking table, and report options.", "success");
  }

  function renderResults() {
    const warningBox = $("comparisonWarning");
    warningBox.classList.add("hidden");
    warningBox.textContent = "";
    if (!state.analysis) return;
    const mode = $("comparisonMode").value;
    const primary = state.analysis.byName[state.primaryTreatment];
    const control = state.analysis.byName[state.controlName];
    const secondary = state.analysis.byName[state.secondaryTreatment];
    if (!primary) return;

    let reference = control;
    let title = `${primary.treatment_name} compared with ${control ? control.treatment_name : "control"}`;
    if (mode === "pair") {
      reference = secondary;
      title = `${primary.treatment_name} compared with ${secondary ? secondary.treatment_name : "selected treatment"}`;
      if (primary.treatment_name === state.secondaryTreatment) {
        warningBox.textContent = "Please choose two different treatments for side-by-side comparison.";
        warningBox.classList.remove("hidden");
      }
    } else if (primary.treatment_name === state.controlName) {
      warningBox.textContent = "The selected treatment is the same as the control. Choose a different treatment to compare against the control.";
      warningBox.classList.remove("hidden");
    }

    const cards = [
      kpiCard("Present value of benefits", money(primary.pvBenefits), title),
      kpiCard("Present value of costs", money(primary.pvCosts), `${primary.replicates} replicates`),
      kpiCard("Net present value", money(primary.npv), reference ? `${signedMoney(primary.npv - reference.npv)} versus comparison` : "No comparison"),
      kpiCard("Benefit-cost ratio", num(primary.bcr, 2), reference ? `${signedNumber(primary.bcr - reference.bcr, 2)} versus comparison` : "No comparison"),
      kpiCard("Gross profit margin", pct(primary.grossProfitMargin), reference ? `${signedNumber(primary.grossProfitMargin - reference.grossProfitMargin, 1)} percentage points` : "No comparison")
    ];
    $("kpiGrid").innerHTML = cards.join("");

    $("comparisonSummary").innerHTML = reference ? `
      <h3>${escapeHtml(title)}</h3>
      <div class="compare-grid">
        ${comparePanel(primary)}
        ${comparePanel(reference)}
      </div>
      <div class="note-box">
        <strong>Difference in net present value:</strong> ${signedMoney(primary.npv - reference.npv)}<br>
        <strong>Difference in present value of benefits:</strong> ${signedMoney(primary.pvBenefits - reference.pvBenefits)}<br>
        <strong>Difference in present value of costs:</strong> ${signedMoney(primary.pvCosts - reference.pvCosts)}
      </div>` : "<p>No comparison selected.</p>";
    renderRankingTable();
  }

  function kpiCard(title, value, sub) {
    return `<div class="kpi-card"><div class="kpi-title">${escapeHtml(title)}</div><div class="kpi-value">${escapeHtml(value)}</div><div class="kpi-sub">${escapeHtml(sub)}</div></div>`;
  }

  function comparePanel(item) {
    return `<div class="compare-panel">
      <h4>${escapeHtml(item.treatment_name)}</h4>
      <div class="metric-line"><span>Replicates</span><strong>${item.replicates}</strong></div>
      <div class="metric-line"><span>Mean yield (t/ha)</span><strong>${num(item.mean_yield_t_ha, 2)}</strong></div>
      <div class="metric-line"><span>Mean cost ($/ha)</span><strong>${money(item.mean_cost_per_ha)}</strong></div>
      <div class="metric-line"><span>PV benefits</span><strong>${money(item.pvBenefits)}</strong></div>
      <div class="metric-line"><span>PV costs</span><strong>${money(item.pvCosts)}</strong></div>
      <div class="metric-line"><span>NPV</span><strong>${money(item.npv)}</strong></div>
      <div class="metric-line"><span>BCR</span><strong>${num(item.bcr, 2)}</strong></div>
      <div class="metric-line"><span>Gross profit margin</span><strong>${pct(item.grossProfitMargin)}</strong></div>
    </div>`;
  }

  function renderRankingTable() {
    const thead = $("rankingTable").querySelector("thead");
    const tbody = $("rankingTable").querySelector("tbody");
    thead.innerHTML = "<tr><th>Rank</th><th>Treatment</th><th>Replicates</th><th>Mean yield (t/ha)</th><th>Mean cost ($/ha)</th><th>NPV</th><th>BCR</th><th>Gross profit margin</th></tr>";
    tbody.innerHTML = state.analysis.treatments.map((item, i) => `<tr>
      <td>${i + 1}</td><td>${escapeHtml(item.treatment_name)}</td><td>${item.replicates}</td>
      <td>${num(item.mean_yield_t_ha, 2)}</td><td>${money(item.mean_cost_per_ha)}</td>
      <td>${money(item.npv)}</td><td>${num(item.bcr, 2)}</td><td>${pct(item.grossProfitMargin)}</td></tr>`).join("");
  }

  function runSensitivity() {
    if (!state.analysis) return;
    const base = state.grouped.find(x => x.treatment_name === state.primaryTreatment) || state.grouped[0];
    const a = state.analysis.assumptions;
    const altPrice = Number($("sensPrice").value || a.grainPrice);
    const benefitPct = Number($("sensBenefitPct").value || 0);
    const costPct = Number($("sensCostPct").value || 0);
    const low = Math.max(0, Number($("sensDiscountLow").value || 0)) / 100;
    const high = Math.max(0, Number($("sensDiscountHigh").value || 0)) / 100;
    state.sensitivity = [
      ["Current assumptions", analyseGroup(base, a)],
      [`Alternative grain price (${money(altPrice)}/t)`, analyseGroup(base, a, { grainPrice: altPrice })],
      [`Benefit adjustment (${benefitPct >= 0 ? "+" : ""}${benefitPct}%)`, analyseGroup(base, a, { benefitPct })],
      [`Cost adjustment (${costPct >= 0 ? "+" : ""}${costPct}%)`, analyseGroup(base, a, { costPct })],
      [`Alternative discount rate A (${(low * 100).toFixed(1)}%)`, analyseGroup(base, { ...a, mode: "constant", initialRate: low, laterRate: low })],
      [`Alternative discount rate B (${(high * 100).toFixed(1)}%)`, analyseGroup(base, { ...a, mode: "constant", initialRate: high, laterRate: high })]
    ];
    const thead = $("sensitivityTable").querySelector("thead");
    const tbody = $("sensitivityTable").querySelector("tbody");
    thead.innerHTML = "<tr><th>Scenario</th><th>PV benefits</th><th>PV costs</th><th>NPV</th><th>BCR</th><th>Gross profit margin</th></tr>";
    tbody.innerHTML = state.sensitivity.map(([label, item]) => `<tr>
      <td>${escapeHtml(label)}</td><td>${money(item.pvBenefits)}</td><td>${money(item.pvCosts)}</td>
      <td>${money(item.npv)}</td><td>${num(item.bcr, 2)}</td><td>${pct(item.grossProfitMargin)}</td></tr>`).join("");
    updateStatus("Sensitivity complete", "Alternative numeric assumptions have been tested for the selected treatment.", "success");
  }

  function createReportHtml() {
    if (!state.analysis) return "";
    const primary = state.analysis.byName[state.primaryTreatment];
    const control = state.analysis.byName[state.controlName];
    const a = state.analysis.assumptions;
    const rows = state.analysis.treatments.map((item, i) => `<tr><td>${i + 1}</td><td>${escapeHtml(item.treatment_name)}</td><td>${item.replicates}</td><td>${num(item.mean_yield_t_ha,2)}</td><td>${money(item.mean_cost_per_ha)}</td><td>${money(item.npv)}</td><td>${num(item.bcr,2)}</td><td>${pct(item.grossProfitMargin)}</td></tr>`).join("");
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>SOIL CRC BCA Basic Report</title><style>
      body{font-family:Arial,sans-serif;color:#1f2a23;padding:28px;line-height:1.45} h1,h2{color:#1f5a3d}
      table{border-collapse:collapse;width:100%} th,td{border:1px solid #d5e2d9;padding:8px;text-align:left} th{background:#edf7f0}
      .grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.card{border:1px solid #d5e2d9;border-radius:12px;padding:10px;background:#f7fbf8}
      </style></head><body>
      <h1>SOIL CRC Benefit-Cost Analysis Report</h1>
      <p><strong>Selected treatment:</strong> ${escapeHtml(primary.treatment_name)}<br><strong>Reference control:</strong> ${escapeHtml(control ? control.treatment_name : "-")}</p>
      <h2>Assumptions used</h2>
      <ul><li>Grain price: ${money(a.grainPrice)} per tonne</li><li>Analysis period: ${a.years} years</li><li>Discounting method: ${escapeHtml(a.mode)}</li><li>Initial discount rate: ${(a.initialRate*100).toFixed(1)}%</li><li>Later discount rate: ${(a.laterRate*100).toFixed(1)}%</li><li>Switch year: ${a.switchYear}</li></ul>
      <h2>Core results</h2><div class="grid">
      <div class="card"><strong>PV benefits</strong><br>${money(primary.pvBenefits)}</div>
      <div class="card"><strong>PV costs</strong><br>${money(primary.pvCosts)}</div>
      <div class="card"><strong>NPV</strong><br>${money(primary.npv)}</div>
      <div class="card"><strong>BCR</strong><br>${num(primary.bcr,2)}</div>
      <div class="card"><strong>Gross profit margin</strong><br>${pct(primary.grossProfitMargin)}</div>
      <div class="card"><strong>Mean yield</strong><br>${num(primary.mean_yield_t_ha,2)} t/ha</div>
      <div class="card"><strong>Mean cost</strong><br>${money(primary.mean_cost_per_ha)} /ha</div>
      <div class="card"><strong>Replicates</strong><br>${primary.replicates}</div></div>
      <h2>Ranking table</h2><table><thead><tr><th>Rank</th><th>Treatment</th><th>Replicates</th><th>Mean yield (t/ha)</th><th>Mean cost ($/ha)</th><th>NPV</th><th>BCR</th><th>Gross profit margin</th></tr></thead><tbody>${rows}</tbody></table>
      <p>Generated from ${escapeHtml(state.workbookName)}.</p></body></html>`;
  }

  function downloadReport() {
    if (!state.analysis) runAnalysis();
    if (!state.analysis) return;
    const blob = new Blob([createReportHtml()], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "soil_crc_bca_basic_report.html";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function printReport() {
    if (!state.analysis) runAnalysis();
    if (!state.analysis) return;
    const win = window.open("", "_blank", "noopener,noreferrer,width=1100,height=800");
    if (!win) return;
    win.document.open();
    win.document.write(createReportHtml());
    win.document.close();
    win.focus();
    setTimeout(() => win.print(), 250);
  }

  function money(value) {
    return Number.isFinite(value) ? `$${Number(value).toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}` : "-";
  }
  function num(value, digits = 2) {
    return Number.isFinite(value) ? Number(value).toLocaleString(undefined, { minimumFractionDigits: digits, maximumFractionDigits: digits }) : "-";
  }
  function pct(value) {
    return Number.isFinite(value) ? `${Number(value).toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%` : "-";
  }
  function signedMoney(value) {
    if (!Number.isFinite(value)) return "-";
    return `${value >= 0 ? "+" : "-"}$${Math.abs(value).toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`;
  }
  function signedNumber(value, digits = 2) {
    if (!Number.isFinite(value)) return "-";
    return `${value >= 0 ? "+" : ""}${Number(value).toLocaleString(undefined, { minimumFractionDigits: digits, maximumFractionDigits: digits })}`;
  }
  function escapeHtml(value) {
    return String(value).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;");
  }

  init();
})();

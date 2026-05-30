"use strict";

const form = document.getElementById("compare-form");
const submitBtn = document.getElementById("submit-btn");
const resetBtn = document.getElementById("reset-btn");
const statusEl = document.getElementById("status");
const resultsEl = document.getElementById("results");
const resultsBody = document.getElementById("results-body");

form.addEventListener("submit", async (evt) => {
  evt.preventDefault();
  await runCompare();
});

resetBtn.addEventListener("click", () => {
  form.reset();
  hide(statusEl);
  hide(resultsEl);
  resultsBody.innerHTML = "";
});

async function runCompare() {
  const data = new FormData();
  const golden = document.getElementById("golden").files[0];
  const others = document.getElementById("others").files;
  const floatTol = document.getElementById("float_tol").value || "0";
  const categories = Array.from(
    document.querySelectorAll('input[name="categories"]:checked'),
  ).map((el) => el.value);

  if (!golden) {
    showStatus("Please choose a golden file.", "error");
    return;
  }
  if (!others.length) {
    showStatus("Please choose at least one candidate file.", "error");
    return;
  }

  data.append("golden", golden);
  for (const f of others) data.append("others", f);
  data.append("float_tol", floatTol);
  for (const c of categories) data.append("categories", c);

  setBusy(true);
  showStatus("Comparing…", "loading");
  hide(resultsEl);
  resultsBody.innerHTML = "";

  try {
    const resp = await fetch("/api/compare", { method: "POST", body: data });
    const payload = await resp.json().catch(() => ({}));

    if (!resp.ok) {
      const detail = payload && payload.detail ? payload.detail : resp.statusText;
      showStatus(`Error: ${detail}`, "error");
      return;
    }

    renderResults(payload);
    hide(statusEl);
    show(resultsEl);
  } catch (err) {
    showStatus(`Network/server error: ${err.message || err}`, "error");
  } finally {
    setBusy(false);
  }
}

function renderResults(payload) {
  resultsBody.innerHTML = "";
  const { golden_filename, results } = payload;

  const intro = document.createElement("p");
  intro.innerHTML = `Compared against golden <code>${escapeHtml(
    golden_filename || "(unknown)",
  )}</code> · <strong>${results.length}</strong> candidate file(s).`;
  resultsBody.appendChild(intro);

  for (const r of results) {
    resultsBody.appendChild(renderFileResult(r));
  }
}

function renderFileResult(r) {
  const details = document.createElement("details");
  details.className = "file-result";
  details.open = true;

  const summary = document.createElement("summary");
  const total = (r.summary && r.summary.TOTAL) || 0;

  const nameEl = document.createElement("span");
  nameEl.className = "filename";
  nameEl.textContent = r.filename;
  summary.appendChild(nameEl);

  if (r.error) {
    summary.appendChild(badge("ERROR", "bad"));
  } else if (total === 0) {
    summary.appendChild(badge("No differences", "ok"));
  } else {
    summary.appendChild(badge(`${total} diff${total === 1 ? "" : "s"}`, "bad"));
    for (const [cat, count] of Object.entries(r.summary || {})) {
      if (cat === "TOTAL" || !count) continue;
      summary.appendChild(badge(`${cat}: ${count}`, `cat-${cat}`));
    }
  }
  details.appendChild(summary);

  const body = document.createElement("div");
  body.className = "file-result-body";

  if (r.error) {
    const err = document.createElement("div");
    err.className = "status error";
    err.textContent = r.error;
    body.appendChild(err);
  } else if (total === 0) {
    const empty = document.createElement("div");
    empty.className = "empty";
    empty.textContent = "✅ No differences found.";
    body.appendChild(empty);
  } else {
    body.appendChild(renderDiffGroups(r.diffs));
  }

  if (!r.error) {
    const downloads = document.createElement("div");
    downloads.className = "download-row";

    const jsonBtn = document.createElement("button");
    jsonBtn.type = "button";
    jsonBtn.textContent = "⬇ Download JSON";
    jsonBtn.addEventListener("click", () =>
      downloadBlob(
        JSON.stringify({ filename: r.filename, summary: r.summary, diffs: r.diffs }, null, 2),
        `${r.filename}.diff.json`,
        "application/json",
      ),
    );
    downloads.appendChild(jsonBtn);

    const txtBtn = document.createElement("button");
    txtBtn.type = "button";
    txtBtn.textContent = "⬇ Download text report";
    txtBtn.addEventListener("click", () =>
      downloadBlob(r.text_report || "No differences.", `${r.filename}.diff.txt`, "text/plain"),
    );
    downloads.appendChild(txtBtn);

    body.appendChild(downloads);
  }

  details.appendChild(body);
  return details;
}

function renderDiffGroups(diffs) {
  const container = document.createElement("div");

  const workbookLevel = [];
  const bySheet = new Map();
  for (const d of diffs) {
    if (d.category === "SHEET_MISSING" || d.category === "SHEET_EXTRA") {
      workbookLevel.push(d);
    } else {
      if (!bySheet.has(d.sheet)) bySheet.set(d.sheet, []);
      bySheet.get(d.sheet).push(d);
    }
  }

  if (workbookLevel.length) {
    const h = document.createElement("h4");
    h.textContent = "Workbook-level";
    container.appendChild(h);
    container.appendChild(buildTable(workbookLevel, /*includeSheet*/ true));
  }

  const sortedSheets = Array.from(bySheet.keys()).sort();
  for (const sheet of sortedSheets) {
    const h = document.createElement("h4");
    h.textContent = `Sheet "${sheet}"`;
    container.appendChild(h);
    container.appendChild(buildTable(bySheet.get(sheet), /*includeSheet*/ false));
  }

  return container;
}

function buildTable(diffs, includeSheet) {
  const table = document.createElement("table");
  table.className = "diff-table";

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  const cols = ["Category"];
  if (includeSheet) cols.push("Sheet");
  cols.push("Cell", "Attribute", "Golden", "Other");
  for (const c of cols) {
    const th = document.createElement("th");
    th.textContent = c;
    headRow.appendChild(th);
  }
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (const d of diffs) {
    const tr = document.createElement("tr");
    tr.appendChild(td(badge(d.category, `cat-${d.category}`)));
    if (includeSheet) tr.appendChild(td(d.sheet || ""));
    tr.appendChild(td(d.cell || ""));
    tr.appendChild(td(d.attribute || ""));
    tr.appendChild(td(formatVal(d.golden), "val"));
    tr.appendChild(td(formatVal(d.other), "val"));
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  return table;
}

function td(content, cls) {
  const el = document.createElement("td");
  if (cls) el.className = cls;
  if (content instanceof Node) el.appendChild(content);
  else el.textContent = content == null ? "" : String(content);
  return el;
}

function badge(text, cls) {
  const span = document.createElement("span");
  span.className = `badge ${cls || ""}`.trim();
  span.textContent = text;
  return span;
}

function formatVal(v) {
  if (v === null || v === undefined) return "∅";
  if (v === "") return '""';
  return String(v);
}

function downloadBlob(text, filename, mime) {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 500);
}

function setBusy(busy) {
  submitBtn.disabled = busy;
  submitBtn.textContent = busy ? "Comparing…" : "Compare";
}

function showStatus(text, kind) {
  statusEl.textContent = text;
  statusEl.className = `status ${kind || ""}`.trim();
  statusEl.hidden = false;
}

function show(el) {
  el.hidden = false;
}
function hide(el) {
  el.hidden = true;
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

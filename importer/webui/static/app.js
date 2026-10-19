const $ = (sel) => document.querySelector(sel);

async function api(path, opts = {}) {
  const res = await fetch(path, opts);
  let body = {};
  try { body = await res.json(); } catch (e) { /* not JSON: the status says it */ }
  if (!res.ok) throw new Error(body.error || `HTTP ${res.status}`);
  return body;
}

function esc(s) {
  return (s ?? "").toString().replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// ---- step navigation ----

function goToStep(n) {
  document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
  $("#step-" + n).classList.add("active");
  document.querySelectorAll(".step").forEach((el) => {
    const s = Number(el.dataset.step);
    el.classList.toggle("active", s === n);
    el.classList.toggle("done", s < n);
  });
}

// ---- modals: open/close with matching motion, Esc + backdrop dismiss, focus restore ----

let lastFocus = null;

function openModal(id) {
  const m = $(id);
  lastFocus = document.activeElement;
  m.classList.remove("closing", "hidden");
  const first = m.querySelector("input, .primary");
  if (first) first.focus({ preventScroll: true });
}

function closeModal(id) {
  const m = $(id);
  if (m.classList.contains("hidden") || m.classList.contains("closing")) return;
  m.classList.add("closing");
  const finish = () => {
    m.classList.add("hidden");
    m.classList.remove("closing");
    if (lastFocus && lastFocus.focus) lastFocus.focus({ preventScroll: true });
  };
  m.addEventListener("animationend", (e) => { if (e.target === m) finish(); }, { once: true });
}

document.querySelectorAll(".modal").forEach((m) =>
  m.addEventListener("pointerdown", (e) => { if (e.target === m) closeModal("#" + m.id); })
);
document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;
  document.querySelectorAll(".modal:not(.hidden)").forEach((m) => closeModal("#" + m.id));
});

// ---- connection (tucked behind the settings modal; connects silently) ----

async function loadConfigIntoModal() {
  const cfg = await api("/api/config");
  $("#s-base_url").value = cfg.base_url;
  $("#s-email").value = cfg.email;
}

async function testConnection() {
  const dot = $("#status-dot");
  dot.className = "status-dot";
  dot.title = "Checking connection…";
  try {
    const r = await api("/api/login_test", { method: "POST" });
    if (r.ok) {
      dot.classList.add("ok");
      dot.title = `Connected — ${r.workspace_name} (${r.user})`;
      $("#settings-status").textContent = `Connected — ${r.workspace_name} (${r.user})`;
    } else {
      dot.classList.add("bad");
      dot.title = r.error;
      $("#settings-status").textContent = r.error;
    }
  } catch (e) {
    dot.classList.add("bad");
    dot.title = e.message;
    $("#settings-status").textContent = e.message;
  }
}

$("#btn-settings").addEventListener("click", async () => {
  await loadConfigIntoModal();
  $("#settings-status").textContent = "";
  openModal("#settings-modal");
});
$("#settings-cancel").addEventListener("click", () => closeModal("#settings-modal"));
$("#settings-save").addEventListener("click", async () => {
  await api("/api/config", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      base_url: $("#s-base_url").value,
      email: $("#s-email").value,
      password: $("#s-password").value,
    }),
  });
  $("#s-password").value = "";
  await testConnection();
});
$("#status-dot").addEventListener("click", () => $("#btn-settings").click());

// ---- danger zone: drop + recreate the dev database, reseed the demo baseline ----

$("#btn-reset-db").addEventListener("click", async () => {
  if (!confirm("This stops the dev stack, drops the margince database, and reseeds the demo baseline. " +
    "Everything currently in the CRM — including anything imported from here — is gone for good. Continue?")) {
    return;
  }
  const btn = $("#btn-reset-db");
  const status = $("#reset-db-status");
  btn.disabled = true;
  status.textContent = "Starting…";
  try {
    const { job_id } = await api("/api/reset_db", { method: "POST" });
    while (true) {
      const job = await api(`/api/job/${job_id}`);
      if (job.status === "running") {
        const secs = Math.round(Date.now() / 1000 - job.started_at);
        status.textContent = `${job.step}… (${secs}s)`;
        await new Promise((r) => setTimeout(r, 1000));
        continue;
      }
      if (job.status === "error") {
        status.textContent = "Failed: " + job.error;
      } else {
        status.textContent = "Done — the CRM is back to the seeded demo baseline. Upload your files again to import.";
      }
      break;
    }
  } catch (e) {
    status.textContent = "Failed: " + e.message;
  } finally {
    btn.disabled = false;
    testConnection();
  }
});

// ---- step 1: the front door. Every file is recognised by the server from what is inside it ----

const dropzone = $("#dropzone");
const fileInput = $("#file-input");
const folderInput = $("#folder-input");

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileInput.click(); }
});
$("#btn-pick-files").addEventListener("click", () => fileInput.click());
$("#btn-pick-folder").addEventListener("click", () => folderInput.click());
["dragover", "dragleave", "drop"].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.toggle("dragover", evt === "dragover");
  })
);
dropzone.addEventListener("drop", async (e) => {
  // The entries must be taken before the first await: the browser empties the DataTransfer afterwards.
  const entries = Array.from(e.dataTransfer.items || []).map((i) => (i.webkitGetAsEntry ? i.webkitGetAsEntry() : null)).filter(Boolean);
  const plain = Array.from(e.dataTransfer.files || []).map((f) => ({ file: f, path: f.name }));
  uploadFiles(entries.length ? await filesOfEntries(entries) : plain);
});
fileInput.addEventListener("change", () => uploadFiles(Array.from(fileInput.files).map((f) => ({ file: f, path: f.name }))));
folderInput.addEventListener("change", () =>
  uploadFiles(Array.from(folderInput.files).map((f) => ({ file: f, path: f.webkitRelativePath || f.name }))));

async function filesOfEntries(entries) {
  const out = [];
  const walk = async (entry) => {
    if (entry.isFile) {
      const file = await new Promise((ok, no) => entry.file(ok, no));
      out.push({ file, path: entry.fullPath.replace(/^\//, "") });
      return;
    }
    const reader = entry.createReader();
    for (;;) {  // readEntries returns the folder in batches
      const batch = await new Promise((ok, no) => reader.readEntries(ok, no));
      if (!batch.length) break;
      for (const child of batch) await walk(child);
    }
  };
  for (const entry of entries) await walk(entry);
  return out;
}

const isExcel = (name) => /\.(xlsx|xlsm|xls)$/i.test(name);
const isJunk = (path) => /(^|\/)(\.|~\$)[^/]*$/.test(path);

async function uploadFiles(items) {
  const errBox = $("#upload-error");
  errBox.classList.add("hidden");
  const inFolder = items.some((i) => i.path.includes("/"));
  const candidates = items.filter((i) => !isJunk(i.path));
  const files = candidates.filter((i) => isExcel(i.path));
  const notExcel = candidates.filter((i) => !isExcel(i.path));
  // A folder holds all sorts of files and only the workbooks matter. Files chosen one by one must all be Excel:
  // never drop one silently.
  if (!files.length || (!inFolder && notExcel.length)) {
    const bad = notExcel.map((i) => i.path).join(", ");
    errBox.textContent = `Please choose Excel files (.xlsx, .xlsm or .xls) only. Not an Excel file: ${bad || "(nothing chosen)"}`;
    errBox.classList.remove("hidden");
    fileInput.value = folderInput.value = "";
    return;
  }
  dropzone.classList.add("dragover");
  dropzone.querySelector(".dz-title").textContent = `Reading ${files.length} file${files.length === 1 ? "" : "s"}…`;

  const form = new FormData();
  files.forEach((i) => form.append("files", i.file, i.path));  // the path keeps the folder: a rep name, or the 01-09 folder
  if ($("#rep-select").value) form.append("rep", $("#rep-select").value);
  try {
    const d = await api("/api/upload", { method: "POST", body: form });
    renderBatch(d);
    goToStep(2);
    if (notExcel.length) toast(`${notExcel.length} file${notExcel.length === 1 ? "" : "s"} in the folder ${notExcel.length === 1 ? "is" : "are"} not Excel and ${notExcel.length === 1 ? "was" : "were"} left out.`);
  } catch (e) {
    errBox.textContent = e.message;
    errBox.classList.remove("hidden");
  } finally {
    dropzone.classList.remove("dragover");
    dropzone.querySelector(".dz-title").textContent = "Drop your Excel lists here";
    fileInput.value = folderInput.value = "";
  }
}

api("/api/reps").then((d) => {
  $("#rep-select").insertAdjacentHTML("beforeend", d.reps.map((r) => `<option>${esc(r)}</option>`).join(""));
});

function toast(message, kind = "") {
  const el = document.createElement("div");
  el.className = `toast ${kind}`.trim();
  el.textContent = message;
  $("#toasts").appendChild(el);
  setTimeout(() => el.remove(), 6000);
}

// ---- step 2: review. The table of files first, then one panel for each kind of list ----

let lastBatch = null;
let activeFam = null;

const FILE_STATUS = {
  ready: ["ok", "Ready"], empty: ["neutral", "Nothing in it"], duplicate: ["neutral", "Identical copy"], superseded: ["neutral", "Older copy"],
  kept_out: ["info", "Kept out on purpose"], unknown: ["warn", "Not recognised"], unreadable: ["bad", "Unreadable"],
  not_read: ["bad", "Not read"],
};

function renderBatch(d) {
  lastBatch = d;
  const b = d.batch;
  renderEntry(b.entry);

  const fams = b.families;
  const keys = fams.map((f) => f.key);
  if (!keys.includes(activeFam)) activeFam = (fams.find((f) => f.ready) || fams[0] || {}).key || null;

  const tabs = $("#family-tabs");
  tabs.classList.toggle("hidden", !fams.length);
  tabs.innerHTML = fams.map((f) => {
    const waiting = f.waiting ? `<span class="badge warn" title="waiting for an answer">${f.waiting} waiting</span>` : "";
    return `<button class="ftab${f.key === activeFam ? " active" : ""}" role="tab" aria-selected="${f.key === activeFam}" data-fam="${f.key}">`
      + `${esc(f.label)}<span class="badge info">${f.files.length} file${f.files.length === 1 ? "" : "s"}</span>${waiting}</button>`;
  }).join("");

  $("#fam-visits").classList.toggle("hidden", activeFam !== "visits");
  if (keys.includes("visits")) renderVisits(d);
  const generic = fams.filter((f) => f.key !== "visits");
  $("#fam-generic").innerHTML = generic.map((f) => familyPanel(f, f.key === activeFam)).join("");
  generic.forEach((f) => renderCompanyQuestions(f.questions, $(`#fam-${f.key} .cq`)));

  const banner = $("#nothing-banner");
  const counts = {};
  b.entry.forEach((e) => { counts[e.status] = (counts[e.status] || 0) + 1; });
  const nothing = !b.ready_families.length;
  banner.classList.toggle("hidden", !nothing);
  if (nothing) {
    const parts = [];
    if (counts.kept_out) parts.push(`${counts.kept_out} recognised and kept out of the CRM on purpose`);
    if (counts.unknown) parts.push(`${counts.unknown} not recognised`);
    if (counts.empty) parts.push(`${counts.empty} with nothing in ${counts.empty === 1 ? "it" : "them"}`);
    if (counts.unreadable || counts.not_read) parts.push(`${(counts.unreadable || 0) + (counts.not_read || 0)} that could not be read`);
    const waiting = fams.some((f) => f.waiting);
    banner.textContent = (waiting ? "Nothing is ready to import yet: answer the company questions below first. " : "Nothing in this upload can be imported. ")
      + (parts.length ? `Of the files: ${parts.join(", ")}.` : "");
  }
  $("#btn-go-import").disabled = nothing;
  $("#btn-dry-run").disabled = nothing;
}

$("#family-tabs").addEventListener("click", (e) => {
  const tab = e.target.closest(".ftab");
  if (!tab) return;
  activeFam = tab.dataset.fam;
  renderBatch(lastBatch);
});

function renderEntry(rows) {
  const n = rows.length;
  const ready = rows.filter((r) => r.status === "ready").length;
  $("#entry-count").textContent = `${n} file${n === 1 ? "" : "s"}, ${ready} ready to import`;
  $("#this-is").innerHTML = rows.map((r) => `<li><b>${esc(r.display_name)}</b>: this is ${esc(r.what)}${r.family ? ` <span class="hint">(${esc(r.family)})</span>` : ""}</li>`).join("");
  $("#entry-body").innerHTML = rows.map((r) => {
    const [cls, label] = FILE_STATUS[r.status] || ["neutral", r.status];
    const folder = r.folder ? `<div class="folder">Folder ${esc(r.folder)}</div>` : "";
    const rowsCell = r.records ? `${r.records.toLocaleString("en")} ${esc(r.unit)}` : "&mdash;";
    const rowClass = r.status === "kept_out" ? "kept" : r.status === "duplicate" || r.status === "superseded" ? "dup" : "";
    return `<tr class="${rowClass}" data-status="${esc(r.status)}">
      <td class="fname" data-label="File">${esc(r.display_name)}${folder}</td>
      <td data-label="What it is">${esc(r.what)}</td>
      <td data-label="Goes into the CRM as">${esc(r.target)}</td>
      <td class="n" data-label="Rows">${rowsCell}</td>
      <td data-label="What happens"><span class="badge ${cls}">${label}</span><div class="outcome">${esc(r.outcome)}</div></td>
    </tr>`;
  }).join("");
}

function familyPanel(f, active) {
  const stats = f.counts.map((c) => `<div class="stat"><div class="n">${esc(c.value)}</div><div class="label">${esc(c.label)}</div></div>`).join("");
  const flags = f.flags.length
    ? `<p class="hint">Flagged for your review: these are imported with a visible tag in the CRM, unless you leave them out here.</p><div class="flag-list">`
      + f.flags.map((fl) => `<label class="flag${fl.included ? "" : " off"}">
          <input type="checkbox" data-family="${f.key}" data-flag="${esc(fl.id)}"${fl.included ? " checked" : ""}>
          <span><span class="flag-name">${esc(fl.label)}</span> <span class="chip">${fl.count}</span>
            <span class="flag-help">${esc(fl.help)}</span></span>
          <span class="badge ${fl.included ? "info" : "neutral"} flag-state">${fl.included ? "Included, tagged" : "Left out"}</span></label>`).join("")
      + `</div>`
    : "";
  const waiting = f.waiting
    ? `<div class="warning">${f.waiting} record${f.waiting === 1 ? " waits" : "s wait"} for ${f.questions.length ? "an answer to the company question" + (f.questions.length === 1 ? "" : "s") + " below" : "a company answer"} and ${f.waiting === 1 ? "is" : "are"} not imported until then.</div>`
    : "";
  const notes = f.notes.length
    ? `<div class="banner"><b>Notes on ${f.files.length === 1 ? "this list" : "these lists"}</b><ul class="note-list">${f.notes.map((n) => `<li>${esc(n)}</li>`).join("")}</ul></div>`
    : "";
  const error = f.error ? `<div class="inline-error">These files could not be read: ${esc(f.error)}</div>` : "";
  return `<div class="fam-panel${active ? "" : " hidden"}" id="fam-${f.key}" data-fam="${f.key}">
    <p class="hint fam-target"><b>Goes into the CRM as:</b> ${esc(f.target)}. ${f.files.map(esc).join(", ")}</p>
    ${error}<div class="summary">${stats}</div>${flags}${waiting}<div class="cq hidden"></div>${notes}
  </div>`;
}

$("#fam-generic").addEventListener("change", async (e) => {
  const box = e.target.closest("input[data-flag]");
  if (!box) return;
  box.disabled = true;
  try {
    renderBatch(await api("/api/flag", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ family: box.dataset.family, flag: box.dataset.flag, include: box.checked }),
    }));
  } catch (err) {
    box.checked = !box.checked;
    box.disabled = false;
    toast(err.message, "bad");
  }
});

function renderVisits(d) {
  $("#summary-grid").innerHTML = [
    ["visits found", d.found],
    ["visits to import", d.visits],
    ["unsure (not imported)", d.unsure],
    ["companies", d.companies],
    ["contacts", d.contacts],
    ["tasks", d.tasks],
    ["weekly reports", d.weeks],
    ["no date", d.no_date],
  ].map(([label, n]) => `<div class="stat"><div class="n">${n}</div><div class="label">${label}</div></div>`).join("");

  const warn = $("#no-date-warning");
  if (d.no_date > 0) {
    warn.classList.remove("hidden");
    warn.textContent = `${d.no_date} visit row(s) have a date cell that doesn't match a known style and will be skipped. First few: ${d.no_date_samples.join(", ")}`;
  } else {
    warn.classList.add("hidden");
  }

  const unsure = $("#unsure-warning");
  if (d.unsure > 0 || (d.skipped_files && d.skipped_files.length)) {
    unsure.classList.remove("hidden");
    const reasons = Object.entries(d.unsure_reasons || {}).map(([why, n]) => `${n}× ${why}`).join("; ");
    const rows = (d.unsure_rows || []).slice(0, 20)
      .map((r) => `${r.file} › ${r.sheet} › row ${r.row}: ${r.date || "no date"} · ${r.customer || "no customer"} — ${r.why}`).join("\n");
    const skipped = (d.skipped_files || []).map((f) => `${f.file}: ${f.why}`).join("\n");
    unsure.textContent = (d.unsure ? `${d.unsure} visit(s) found but will NOT be imported: ${reasons}\n`
        + (d.needs_excel_fix ? `${d.needs_excel_fix} of them can only be cleared by correcting the cell in the Excel file and uploading it again; the rest import once the company questions below are answered.\n` : "")
        + rows : "")
      + (skipped ? `\nFiles and sheets not used:\n${skipped}` : "");
  } else {
    unsure.classList.add("hidden");
  }

  renderCompanyQuestions(d.company_questions || [], $("#company-questions"));

  const sep = $("#separate-warning");
  if (d.decisions && d.decisions.length) {
    sep.classList.remove("hidden");
    sep.innerHTML = `${d.decisions.length} short company name(s) could not be tied to one full name and will be imported as their own company — check these after the import and use "Merge company" where they are the same: ` +
      d.decisions.map((x) => `<b>${esc(x.written)}</b> (${esc(x.city) || "no city"})`).join(", ");
  } else {
    sep.classList.add("hidden");
  }

  $("#companies-body").innerHTML = d.company_rows
    .map((c) => `<tr><td>${esc(c.name)}</td><td>${esc(c.city) || "—"}</td><td>${esc(c.potential) || "—"}</td>
      <td>${c.type ? `<span class="tag">${esc(c.type)}</span>` : "—"}</td><td>${esc(c.reps)}</td>
      <td>${c.visits}</td><td>${esc(c.last_visit) || "—"}</td></tr>`)
    .join("");
  $("#contacts-body").innerHTML = d.contact_rows
    .map((c) => `<tr><td>${esc(c.full_name)}</td><td>${esc(c.company)}</td><td>${esc(c.phones) || "—"}</td><td>${esc(c.emails) || "—"}</td></tr>`)
    .join("");
}

document.querySelectorAll(".tab").forEach((tab) =>
  tab.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
    document.querySelectorAll(".tab-panel").forEach((p) => p.classList.add("hidden"));
    tab.classList.add("active");
    $("#panel-" + tab.dataset.tab).classList.remove("hidden");
  })
);

async function startOver() {
  await api("/api/clear", { method: "POST" });
  lastBatch = null;
  activeFam = null;
  goToStep(1);
}
$("#btn-start-over").addEventListener("click", startOver);

// ---- dry run and import: the same look at what will be written; only the import writes ----

// short label/value pairs in a grid; a long text value (a list of field names) goes on a line of its own
function rowsHtml(rows, extra = "") {
  const long = (r) => typeof r.value === "string" && r.value.length > 24;
  const pairs = rows.filter((r) => !long(r)).map((r) => `<dt>${esc(r.label)}</dt><dd>${esc(r.value)}</dd>`).join("");
  const texts = rows.filter(long).map((r) => `<p class="long"><b>${esc(r.label)}:</b> ${esc(r.value)}</p>`).join("");
  return `<dl>${pairs}${extra}</dl>${texts}`;
}

function dryCard(f) {
  const notes = f.notes.length ? `<ul class="note-list">${f.notes.map((n) => `<li>${esc(n)}</li>`).join("")}</ul>` : "";
  const error = f.error ? `<p class="cq-error">Could not be checked: ${esc(f.error)}</p>` : "";
  return `<div class="report-card" data-fam="${f.key}"><h4>${esc(f.label)}</h4>${rowsHtml(f.rows)}${notes}${error}</div>`;
}

async function showPlan(forImport) {
  const btn = forImport ? $("#btn-go-import") : $("#btn-dry-run");
  const errBox = $("#import-review-error");
  errBox.classList.add("hidden");
  const label = btn.textContent;
  $("#btn-go-import").disabled = $("#btn-dry-run").disabled = true;
  btn.textContent = "Checking…";
  try {
    const r = await api("/api/dry_run", { method: "POST" });
    $("#confirm-report").innerHTML = r.families.map(dryCard).join("");
    $("#confirm-title").textContent = forImport ? "Ready to import" : "Test import (dry run)";
    $("#confirm-hint").textContent = forImport
      ? "Here's what will land in the CRM. This step writes data. Running it again never duplicates."
      : "Nothing was written. This is exactly what an import would do right now.";
    $("#confirm-ok").classList.toggle("hidden", !forImport);
    $("#confirm-cancel").textContent = forImport ? "Cancel" : "Close";
    openModal("#confirm-modal");
  } catch (e) {
    errBox.textContent = "Could not check the import: " + e.message;
    errBox.classList.remove("hidden");
  } finally {
    $("#btn-go-import").disabled = $("#btn-dry-run").disabled = false;
    btn.textContent = label;
  }
}
$("#btn-go-import").addEventListener("click", () => showPlan(true));
$("#btn-dry-run").addEventListener("click", () => showPlan(false));

$("#confirm-cancel").addEventListener("click", () => closeModal("#confirm-modal"));
$("#confirm-ok").addEventListener("click", async () => {
  closeModal("#confirm-modal");
  goToStep(3);
  await runCommit();
});

const STEP_LABEL = {
  starting: "Starting…",
  companies_import_start: "Importing companies…",
  companies_import_done: "Companies imported",
  vocabulary: "Preparing fields…",
  contacts_start: "Writing contacts…",
  activities_start: "Logging visits and tasks…",
  weeks_start: "Writing weekly reports…",
};
const STATE_CLASS = { waiting: "skip", running: "info", done: "ok", nothing_new: "ok", partial: "warn", failed: "bad" };
const STATE_LABEL = { waiting: "Waiting", running: "Writing…", done: "Done", nothing_new: "Nothing new", partial: "Partly done", failed: "Failed" };

function setProgress(pct, label) {
  $("#progress-fill").style.width = pct + "%";
  $("#progress-label").textContent = label;
}

function stateLine(s) {
  if (s.status === "running") {
    if (s.total) return `${s.done ?? 0} of ${s.total}`;
    return STEP_LABEL[s.step] || (s.step ? s.step.replace(/_/g, " ") : "Writing…");
  }
  if (s.status === "waiting") return "";
  return `${s.written ?? 0} written, ${s.present ?? 0} already in the CRM` + (s.failed ? `, ${s.failed} failed` : "");
}

function renderProgress(job) {
  const order = job.order || Object.keys(job.families);
  $("#family-progress").innerHTML = order.map((k) => {
    const s = job.families[k];
    const line = stateLine(s);
    return `<li class="${STATE_CLASS[s.status] || "skip"}"><div>${esc(s.label)}${line ? `<span class="sub">${esc(line)}</span>` : ""}</div>`
      + `<span class="badge ${s.status === "running" ? "info" : s.status === "waiting" ? "neutral" : STATE_CLASS[s.status] || "neutral"}">${STATE_LABEL[s.status] || s.status}</span></li>`;
  }).join("");
  const parts = order.map((k) => {
    const s = job.families[k];
    if (s.status === "running") return s.total ? Math.min(0.95, (s.done ?? 0) / s.total) : 0.15;
    return s.status === "waiting" ? 0 : 1;
  });
  const pct = 4 + 96 * (parts.reduce((a, b) => a + b, 0) / (order.length || 1));
  const cur = job.families[job.family];
  setProgress(pct, cur ? `${cur.label}: ${stateLine(cur) || "starting"}` : "Starting…");
}

async function runCommit() {
  $("#import-working").classList.remove("hidden");
  $("#import-done").classList.add("hidden");
  $("#import-error").classList.add("hidden");
  $("#family-progress").innerHTML = "";
  setProgress(4, "Starting…");

  try {
    const { job_id } = await api("/api/commit", { method: "POST" });
    await pollJob(job_id);
  } catch (e) {
    showError(e.message);
  }
}

async function pollJob(jobId) {
  while (true) {
    const job = await api(`/api/commit/${jobId}`);
    if (job.status === "running") {
      renderProgress(job);
      await new Promise((r) => setTimeout(r, 400));
      continue;
    }
    if (job.status === "error") {
      showError(job.error);
      return;
    }
    showDone(job.result);
    return;
  }
}

const RESULT_BADGE = { done: "ok", nothing_new: "info", partial: "warn", failed: "bad" };

function resultCard(f) {
  const lead = f.status === "nothing_new" ? `<p class="lead">Nothing new: ${f.present} already in the CRM.</p>`
    : f.status === "failed" ? `<p class="lead">Nothing was written.</p>` : "";
  const skipped = f.skipped ? `<dt>Skipped</dt><dd>${f.skipped}</dd>` : "";
  const notes = f.notes.length ? `<ul class="note-list">${f.notes.map((n) => `<li>${esc(n)}</li>`).join("")}</ul>` : "";
  const errors = f.errors.length ? `<ul class="note-list errors">${f.errors.map((n) => `<li>${esc(n)}</li>`).join("")}</ul>` : "";
  return `<div class="report-card fam-result" data-fam="${f.key}" data-status="${f.status}">
    <h4>${esc(f.label)} <span class="badge ${RESULT_BADGE[f.status] || "neutral"}">${STATE_LABEL[f.status] || f.status}</span></h4>
    ${lead}${rowsHtml(f.rows, skipped)}${notes}${errors}</div>`;
}

function showDone(result) {
  $("#import-working").classList.add("hidden");
  $("#import-done").classList.remove("hidden");
  const fams = result.families;
  const written = fams.reduce((a, f) => a + f.written, 0);
  const present = fams.reduce((a, f) => a + f.present, 0);
  const bad = fams.filter((f) => f.status === "failed" || f.status === "partial");
  const hero = $("#result-hero");
  hero.classList.toggle("result-hero-error", bad.length > 0);
  if (bad.length) {
    $("#result-title").textContent = "Import finished with problems";
    $("#result-sub").textContent = `${bad.map((f) => f.label).join(", ")} could not be written completely; the other lists were not affected. `
      + `${written} records written, ${present} already in the CRM.`;
  } else if (!written) {
    $("#result-title").textContent = "Nothing new";
    $("#result-sub").textContent = `Everything in these files is already in the CRM (${present} records). Nothing was written.`;
  } else {
    $("#result-title").textContent = "Import complete";
    $("#result-sub").textContent = `${written} records written, ${present} were already in the CRM.`;
  }
  $("#result-report").innerHTML = fams.map(resultCard).join("");
}

function showError(message) {
  $("#import-working").classList.add("hidden");
  $("#import-error").classList.remove("hidden");
  $("#error-text").textContent = message;
}

$("#btn-import-another").addEventListener("click", startOver);
$("#btn-back-to-review").addEventListener("click", () => goToStep(2));

// ---- stale-server guard: warn when the files on disk are newer than the code this server runs ----

async function checkVersion() {
  try {
    const v = await api("/api/version");
    $("#stale-banner").classList.toggle("hidden", !v.stale);
  } catch (e) { /* server gone: the other calls will say so */ }
}
setInterval(checkVersion, 10000);

// ---- init ----

(async function init() {
  await loadConfigIntoModal();
  testConnection(); // fire and forget - never blocks the UI
  checkVersion();
  const d = await api("/api/batch");
  if (d.batch.entry.length) {
    renderBatch(d);
    goToStep(2);
  }
})();


// ---- company questions: "New company" or "Existing company" ----

function renderCompanyQuestions(questions, box) {
  if (!box) return;
  if (!questions.length) {
    box.classList.add("hidden");
    box.innerHTML = "";
    return;
  }
  box.classList.remove("hidden");
  const suggested = questions.filter((q) => q.suggest);
  box.innerHTML = `<h3 class="cq-title">${questions.length} ${questions.length === 1 ? "question needs" : "questions need"} your answer</h3>`
    + (suggested.length > 1 ? `<div class="cq-bulk"><button class="cq-btn cq-all">Confirm all ${suggested.length} suggestions</button>`
      + `<span class="cq-sub">Links each name to the closest company in your list (85% similar or more). The others stay open.</span><div class="cq-error hidden"></div></div>` : "")
    + questions.map((q, i) => q.kind === "pair" ? pairCard(q, i) : `
    <div class="cq-card" data-i="${i}">
      <div class="cq-name">${esc(q.name)}</div>
      <div class="cq-sub">as written in the Excel &middot; ${waitingText(q)} waiting</div>
      <div class="cq-question">Is that the same company?</div>
      <div class="cq-choices">
        ${q.candidates.slice(0, 3).map((c) => `<button type="button" class="cq-btn cq-cand" data-code="${esc(c.code)}">
          <span class="cq-cand-name">${esc(c.name)}</span>
          <span class="cq-cand-note">${esc([c.en && c.en !== c.name ? c.en : "", c.uscc ? "USCC " + c.uscc : "", c.legal ? "legal rep " + c.legal : "", Math.round(c.score * 100) + "% similar"].filter(Boolean).join(" · "))}</span></button>`).join("")}
        <button type="button" class="cq-btn cq-new" data-act="new">New company</button>
      </div>
      <button type="button" class="cq-link" data-act="existing">Not one of these? Search the customer list</button>
      <div class="cq-pick hidden">
        <div class="cq-label">Search by name or code</div>
        <input class="cq-search" type="search" placeholder="Type a name or customer code" aria-label="Search the customer list">
        <div class="cq-options cq-results"></div>
        <div class="cq-confirm hidden"><button class="cq-btn cq-go" disabled>Link to this company</button></div>
      </div>
      <div class="cq-error hidden"></div>
    </div>`).join("");
  box.querySelectorAll(".cq-card").forEach((card) => {
    const q = questions[Number(card.dataset.i)];
    if (q.kind === "pair") wirePair(card, q); else wireQuestion(card, q);
  });
  const all = box.querySelector(".cq-all");
  if (all) all.addEventListener("click", async () => {
    const err = box.querySelector(".cq-bulk .cq-error");
    err.classList.add("hidden");
    box.querySelectorAll("button, input").forEach((el) => { el.disabled = true; });
    try {
      renderBatch(await api("/api/company_answers", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ answers: suggested.map((q) => ({ name: q.name, code: q.suggest.code })) }),
      }));
    } catch (e) {
      err.textContent = e.message;
      err.classList.remove("hidden");
      box.querySelectorAll("button, input").forEach((el) => { el.disabled = false; });
    }
  });
}

function waitingText(q) {
  return q.unit ? `${q.count} ${q.count === 1 ? q.unit.replace(/s$/, "") : q.unit}` : `${q.visits} visit${q.visits === 1 ? "" : "s"}`;
}

function pairCard(q, i) {
  return `
    <div class="cq-card" data-i="${i}">
      <div class="cq-sub">Are these two names the same company?</div>
      <div class="cq-names">
        <div class="cq-name">${esc(q.a)}</div>
        <div class="cq-name">${esc(q.b)}</div>
      </div>
      <div class="cq-sub">${waitingText(q)} waiting. If they are the same, the fuller name is used.</div>
      <div class="cq-actions">
        <button class="cq-btn" data-act="same">Same company</button>
        <button class="cq-btn" data-act="different">Different companies</button>
      </div>
      <div class="cq-error hidden"></div>
    </div>`;
}

function wirePair(card, q) {
  const err = card.querySelector(".cq-error");
  const send = async (same) => {
    err.classList.add("hidden");
    card.querySelectorAll("button").forEach((b) => { b.disabled = true; });
    try {
      const d = await api("/api/company_pair_answer", {
        method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ a: q.a, b: q.b, same }),
      });
      renderBatch(d);
    } catch (e) {
      err.textContent = e.message;
      err.classList.remove("hidden");
      card.querySelectorAll("button").forEach((b) => { b.disabled = false; });
    }
  };
  card.querySelector('[data-act="same"]').addEventListener("click", () => send(true));
  card.querySelector('[data-act="different"]').addEventListener("click", () => send(false));
}

function cqOption(code, name, note) {
  return `<button type="button" class="cq-opt" data-code="${esc(code)}"><span class="cq-opt-name">${esc(name)}</span>`
    + `<span class="cq-opt-note">${esc(note || code)}</span></button>`;
}

function wireQuestion(card, q) {
  const pick = card.querySelector(".cq-pick");
  const go = card.querySelector(".cq-go");
  const err = card.querySelector(".cq-error");
  let chosen = null;

  const setBusy = (busy) => card.querySelectorAll("button, input").forEach((el) => { el.disabled = busy || (el === go && !chosen); });
  const send = async (code) => {
    err.classList.add("hidden");
    setBusy(true);
    try {
      const d = await api("/api/company_answer", {
        method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: q.name, code }),
      });
      renderBatch(d);
    } catch (e) {
      err.textContent = e.message;
      err.classList.remove("hidden");
      setBusy(false);
    }
  };
  const choose = (opt) => {
    card.querySelectorAll(".cq-opt").forEach((o) => o.classList.toggle("selected", o === opt));
    chosen = opt.dataset.code;
    go.disabled = false;
    card.querySelector(".cq-confirm").classList.remove("hidden");
    go.textContent = `Link to ${opt.querySelector(".cq-opt-name").textContent}`;
  };

  card.querySelector('[data-act="new"]').addEventListener("click", () => send(null));
  card.querySelectorAll(".cq-cand").forEach((b) => b.addEventListener("click", () => send(b.dataset.code)));
  card.querySelector('[data-act="existing"]').addEventListener("click", () => {
    pick.classList.toggle("hidden");
    card.querySelector('[data-act="existing"]').classList.toggle("on", !pick.classList.contains("hidden"));
    if (!pick.classList.contains("hidden")) card.querySelector(".cq-search").focus();
  });
  card.addEventListener("click", (ev) => {
    const opt = ev.target.closest(".cq-opt");
    if (opt) choose(opt);
  });
  go.addEventListener("click", () => chosen && send(chosen));

  let timer = null;
  card.querySelector(".cq-search").addEventListener("input", (ev) => {
    clearTimeout(timer);
    const text = ev.target.value.trim();
    const results = card.querySelector(".cq-results");
    if (!text) { results.innerHTML = ""; return; }
    timer = setTimeout(async () => {
      try {
        const d = await api("/api/company_search?q=" + encodeURIComponent(text));
        results.innerHTML = d.results.length
          ? d.results.map((c) => cqOption(c.code, c.cn || c.en || c.short, [c.en && c.cn ? c.en : "", c.code, c.uscc ? "USCC " + c.uscc : ""].filter(Boolean).join(" · "))).join("")
          : `<div class="cq-none">No company found. If it is not in your list, choose "New company".</div>`;
      } catch (e) {
        results.innerHTML = `<div class="cq-none">${esc(e.message)}</div>`;
      }
    }, 200);
  });
}

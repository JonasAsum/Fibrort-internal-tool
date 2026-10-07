const $ = (sel) => document.querySelector(sel);

async function api(path, opts = {}) {
  const res = await fetch(path, opts);
  const body = await res.json();
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

// ---- step 1: upload ----

const dropzone = $("#dropzone");
const fileInput = $("#file-input");

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileInput.click(); }
});
["dragover", "dragleave", "drop"].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.toggle("dragover", evt === "dragover");
  })
);
dropzone.addEventListener("drop", (e) => uploadFiles(e.dataTransfer.files));
fileInput.addEventListener("change", () => uploadFiles(fileInput.files));

async function uploadFiles(fileList) {
  const files = Array.from(fileList).filter((f) => f.name.toLowerCase().endsWith(".xlsx"));
  const errBox = $("#upload-error");
  errBox.classList.add("hidden");
  if (!files.length) {
    errBox.textContent = "Please choose one or more .xlsx files.";
    errBox.classList.remove("hidden");
    return;
  }
  dropzone.classList.add("dragover");
  dropzone.querySelector(".dz-title").textContent = `Reading ${files.length} file(s)…`;

  const form = new FormData();
  files.forEach((f) => form.append("files", f));
  try {
    const summary = await api("/api/upload", { method: "POST", body: form });
    renderReview(summary);
    goToStep(2);
  } catch (e) {
    errBox.textContent = e.message;
    errBox.classList.remove("hidden");
  } finally {
    dropzone.classList.remove("dragover");
    dropzone.querySelector(".dz-title").textContent = "Drop your visit-list workbook(s) here";
    fileInput.value = "";
  }
}

// ---- step 2: review ----

function renderReview(d) {
  $("#file-chips").innerHTML = d.files.map((f) => `<span class="file-chip">${esc(f)}</span>`).join("");

  $("#summary-grid").innerHTML = [
    ["visits", d.visits],
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

$("#btn-start-over").addEventListener("click", async () => {
  await api("/api/clear", { method: "POST" });
  goToStep(1);
});

// ---- import: preview as a confirm step, then commit with progress ----

function dispositionCard(title, report) {
  if (!report) return "";
  const d = report.disposition;
  const rows = Object.entries(d).map(([k, v]) => `<dt>${k}</dt><dd>${v}</dd>`).join("");
  const issues = (report.issues || []).slice(0, 6)
    .map((i) => `<li><b>line ${i.line}${i.column ? " / " + esc(i.column) : ""}:</b> ${esc(i.reason)}</li>`).join("");
  return `
    <div class="report-card">
      <h4>${title} — ${report.rows_read} rows</h4>
      <dl>${rows}</dl>
      ${issues ? `<ul class="issues">${issues}</ul>` : ""}
    </div>`;
}

$("#btn-go-import").addEventListener("click", async () => {
  const btn = $("#btn-go-import");
  const errBox = $("#review-error");
  errBox.classList.add("hidden");
  btn.disabled = true;
  btn.textContent = "Checking…";
  try {
    const r = await api("/api/preview", { method: "POST" });
    const p = r.plan;
    $("#confirm-report").innerHTML = dispositionCard("Companies", r.company) +
      `<div class="report-card"><h4>Also written</h4><dl>
        <dt>contacts</dt><dd>${p.contacts} (${p.contacts_with_phone} with phone, ${p.contacts_with_email} with email)</dd>
        <dt>visit meetings</dt><dd>${p.meetings}</dd>
        <dt>visit notes</dt><dd>${p.notes}</dd>
        <dt>tasks</dt><dd>${p.tasks}${p.task_cutoff ? ` (due before ${esc(p.task_cutoff)} marked done)` : ""}</dd>
        <dt>weekly reports</dt><dd>${p.weeks}</dd>
      </dl></div>`;
    openModal("#confirm-modal");
  } catch (e) {
    errBox.textContent = "Could not check the import: " + e.message;
    errBox.classList.remove("hidden");
  } finally {
    btn.disabled = false;
    btn.textContent = "Import to CRM →";
  }
});

$("#confirm-cancel").addEventListener("click", () => closeModal("#confirm-modal"));
$("#confirm-ok").addEventListener("click", async () => {
  closeModal("#confirm-modal");
  goToStep(3);
  await runCommit();
});

function setProgress(pct, label) {
  $("#progress-fill").style.width = pct + "%";
  $("#progress-label").textContent = label;
}

const STEP_LABEL = {
  starting: "Starting…",
  companies_import_start: "Importing companies…",
  companies_import_done: "Companies imported",
  vocabulary: "Preparing fields…",
  contacts_start: "Writing contacts…",
  activities_start: "Logging visits and tasks…",
  weeks_start: "Writing weekly reports…",
};

async function runCommit() {
  $("#import-working").classList.remove("hidden");
  $("#import-done").classList.add("hidden");
  $("#import-error").classList.add("hidden");
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
      const ranges = { companies_enrich: [20, 30, "Filling in company fields"], contacts_progress: [30, 45, "Writing contacts"],
        activity_progress: [45, 92, "Logging visits and tasks"], weeks_progress: [92, 100, "Writing weekly reports"] };
      if (ranges[job.step]) {
        const [from, to, label] = ranges[job.step];
        setProgress(from + (to - from) * (job.done / job.total), `${label}… ${job.done}/${job.total}`);
      } else {
        const base = { vocabulary: 3, companies_import_start: 6, companies_import_done: 20, contacts_start: 30,
          activities_start: 45, weeks_start: 92 }[job.step] ?? 4;
        setProgress(base, STEP_LABEL[job.step] || job.step);
      }
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

function showDone(r) {
  $("#import-working").classList.add("hidden");
  $("#import-done").classList.remove("hidden");
  $("#result-sub").textContent =
    `${r.activities_created} visits logged, ${r.activities_idempotent} were already there.`;
  const reps = Object.keys(r.reps_with_seat || {});
  $("#result-report").innerHTML =
    dispositionCard("Companies", r.companies_report) +
    `<div class="report-card"><h4>Records</h4><dl>
      <dt>companies enriched</dt><dd>${r.companies_updated}</dd>
      <dt>contacts created</dt><dd>${r.contacts_created}</dd>
      <dt>contacts given phone/email</dt><dd>${r.contacts_updated}</dd>
      <dt>custom fields added</dt><dd>${esc((r.fields_created || []).join(", ")) || "none"}</dd>
    </dl></div>` +
    `<div class="report-card"><h4>Timeline</h4><dl>
      <dt>visits created</dt><dd>${r.activities_created}</dd>
      <dt>visits already present</dt><dd>${r.activities_idempotent}</dd>
      <dt>skipped (no date)</dt><dd>${r.activities_skipped_no_date}</dd>
      <dt>tasks created</dt><dd>${r.tasks_created} (${r.tasks_closed} already done)</dd>
      <dt>weekly reports</dt><dd>${r.weeks_created} new, ${r.weeks_idempotent} present</dd>
    </dl></div>` +
    `<div class="report-card"><h4>Sales reps</h4><dl>
      <dt>matched to a seat</dt><dd>${esc(reps.join(", ")) || "none"}</dd>
      <dt>named as author only</dt><dd>${esc((r.reps_without_seat || []).join(", ")) || "none"}</dd>
    </dl></div>`;
}

function showError(message) {
  $("#import-working").classList.add("hidden");
  $("#import-error").classList.remove("hidden");
  $("#error-text").textContent = message;
}

$("#btn-import-another").addEventListener("click", async () => {
  await api("/api/clear", { method: "POST" });
  goToStep(1);
});
$("#btn-back-to-review").addEventListener("click", () => goToStep(2));

// ---- init ----

(async function init() {
  await loadConfigIntoModal();
  testConnection(); // fire and forget - never blocks the UI
  const d = await api("/api/dataset");
  if (d.loaded) {
    renderReview(d);
    goToStep(2);
  }
})();

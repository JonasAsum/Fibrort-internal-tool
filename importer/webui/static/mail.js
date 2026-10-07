// Email archive tab: extracted.json (from mail_import) -> Margince CRM.
// Shares $, api, esc, openModal and closeModal with app.js (loaded first).

let mailStep = 1;

// ---- source switch ----

function showMode(mode) {
  document.querySelectorAll(".mode-tab").forEach((t) => {
    const on = t.dataset.mode === mode;
    t.classList.toggle("active", on);
    t.setAttribute("aria-selected", on);
  });
  $("#view-visits").classList.toggle("hidden", mode !== "visits");
  $("#view-mail").classList.toggle("hidden", mode !== "mail");
  $("#view-review").classList.toggle("hidden", mode !== "review");
  if (mode === "review" && typeof reviewLoadCases === "function") reviewLoadCases(); // review.js loads after this file
  try { sessionStorage.setItem("importMode", mode); } catch (e) { /* private window: just don't remember */ }
  if (mode === "mail") mailGoToStep(mailStep); // app.js's goToStep clears every .step while the other view is open
}
document.querySelectorAll(".mode-tab").forEach((t) => t.addEventListener("click", () => showMode(t.dataset.mode)));

function mailGoToStep(n) {
  mailStep = n;
  document.querySelectorAll(".mpanel").forEach((p) => p.classList.remove("active"));
  $("#mstep-" + n).classList.add("active");
  document.querySelectorAll("[data-mstep]").forEach((el) => {
    const s = Number(el.dataset.mstep);
    el.classList.toggle("active", s === n);
    el.classList.toggle("done", s < n);
  });
}

// ---- step 1: upload ----

const mdrop = $("#mdropzone");
const mfile = $("#mfile-input");
mdrop.addEventListener("click", () => mfile.click());
mdrop.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); mfile.click(); }
});
["dragover", "dragleave", "drop"].forEach((evt) =>
  mdrop.addEventListener(evt, (e) => {
    e.preventDefault();
    mdrop.classList.toggle("dragover", evt === "dragover");
  })
);
mdrop.addEventListener("drop", (e) => mailUpload(e.dataTransfer.files));
mfile.addEventListener("change", () => mailUpload(mfile.files));

async function mailUpload(fileList) {
  const files = Array.from(fileList).filter((f) => f.name.toLowerCase().endsWith(".json"));
  const errBox = $("#mupload-error");
  errBox.classList.add("hidden");
  if (!files.length) {
    errBox.textContent = "Please choose extracted.json (and optionally enrichment.json).";
    errBox.classList.remove("hidden");
    return;
  }
  mdrop.classList.add("dragover");
  mdrop.querySelector(".dz-title").textContent = "Reading…";
  const form = new FormData();
  files.forEach((f) => form.append("files", f));
  try {
    renderMailReview(await api("/api/mail/upload", { method: "POST", body: form }));
    mailGoToStep(2);
  } catch (e) {
    errBox.textContent = e.message;
    errBox.classList.remove("hidden");
  } finally {
    mdrop.classList.remove("dragover");
    mdrop.querySelector(".dz-title").textContent = "Drop extracted.json here";
    mfile.value = "";
  }
}

$("#mbtn-use-current").addEventListener("click", async () => {
  const btn = $("#mbtn-use-current");
  const errBox = $("#mupload-error");
  errBox.classList.add("hidden");
  btn.disabled = true;
  btn.textContent = "Reading mail/staging…";
  try {
    renderMailReview(await api("/api/mail/use_current", { method: "POST" }));
    mailGoToStep(2);
  } catch (e) {
    errBox.textContent = e.message;
    errBox.classList.remove("hidden");
  } finally {
    btn.disabled = false;
    btn.textContent = "Use the current extraction";
  }
});

// ---- step 2: review ----

function renderMailReview(d) {
  const c = d.counts;
  $("#mfile-chips").innerHTML = (d.files || ["extracted.json"]).map((f) => `<span class="file-chip">${esc(f)}</span>`).join("");
  $("#msummary-grid").innerHTML = [
    ["emails", c.emails], ["contacts", c.contacts], ["with phone", c.contacts_with_phone],
    ["companies", c.companies], ["deals", c.deals], ["attachments", c.attachments],
  ].map(([label, n]) => `<div class="stat"><div class="n">${n ?? 0}</div><div class="label">${label}</div></div>`).join("");

  const notes = [];
  (d.warnings || []).forEach((w) => notes.push(esc(w)));
  if (d.documents_unreadable) notes.push(`${d.documents_unreadable} PDFs are scans without text; their facts need the reading pass.`);
  const r = d.reading;
  if (r && r.waiting_for_review !== undefined) {
    notes.unshift(`Read from the mails: ${r.auto} agreed, ${r.accepted} checked by you go in` +
      (r.waiting_for_review ? ` — <b>${r.waiting_for_review} still wait for your review and stay out</b> (Reading review tab)` : "") +
      (r.rejected_by_review ? `; ${r.rejected_by_review} you rejected` : "") + ".");
  }
  $("#mwarnings").innerHTML = notes.length
    ? `<div class="warning"><ul>${notes.map((n) => `<li>${n}</li>`).join("")}</ul></div>` : "";

  $("#mcompanies-body").innerHTML = d.company_rows.map((r) =>
    `<tr><td>${esc(r.name)}</td><td>${r.type ? `<span class="tag">${esc(r.type)}</span>` : "—"}</td><td>${esc(r.line) || "—"}</td>
     <td>${esc(r.domains) || "—"}</td><td>${r.messages}</td><td>${esc(r.origin) || "—"}</td></tr>`).join("");
  $("#mcontacts-body").innerHTML = d.contact_rows.map((r) =>
    `<tr><td>${esc(r.full_name)}</td><td>${esc(r.company) || "—"}</td><td>${esc(r.title) || "—"}</td>
     <td>${esc(r.phones) || "—"}</td><td>${esc(r.emails)}</td><td>${r.messages}</td></tr>`).join("");

  // Attachment files live next to the extractor's output, not in the JSON: offer them only when they are there.
  const box = $("#m-attachments");
  const have = d.attachments_on_disk > 0;
  box.disabled = !have;
  box.checked = have && !d.attachments_missing;
  $("#m-attachments-label").textContent = have
    ? `Upload attachment files (${d.attachments_on_disk} found${d.attachments_missing ? `, ${d.attachments_missing} missing` : ""})`
    : "Attachment files not found next to the extractor — only the mail itself is imported";
  box.closest("label").classList.toggle("off", !have);
}

document.querySelectorAll(".mtab").forEach((tab) =>
  tab.addEventListener("click", () => {
    document.querySelectorAll(".mtab").forEach((t) => t.classList.remove("active"));
    document.querySelectorAll(".mtab-panel").forEach((p) => p.classList.add("hidden"));
    tab.classList.add("active");
    $("#mpanel-" + tab.dataset.mtab).classList.remove("hidden");
  })
);

$("#mbtn-start-over").addEventListener("click", async () => {
  await api("/api/mail/clear", { method: "POST" });
  mailGoToStep(1);
});

// ---- preview as a confirm step ----

function kv(rows) { return rows.map(([k, v]) => `<dt>${k}</dt><dd>${v}</dd>`).join(""); }

$("#mbtn-go-import").addEventListener("click", async () => {
  const btn = $("#mbtn-go-import");
  const errBox = $("#mreview-error");
  errBox.classList.add("hidden");
  btn.disabled = true;
  btn.textContent = "Checking…";
  try {
    const r = await api("/api/mail/preview", { method: "POST" });
    const p = r.plan, crm = r.crm;
    const limit = Number($("#m-limit").value) || 0;
    $("#mconfirm-report").innerHTML =
      `<div class="report-card"><h4>New in the CRM</h4><dl>${kv([
        ["companies", `${crm.companies_new} of ${p.companies}`],
        ["contacts", `${crm.contacts_new} of ${p.contacts}`],
        ["deals", `${crm.deals_new} of ${p.deals}`],
        ["products", `${crm.products_new} of ${p.products}`],
      ])}</dl></div>` +
      `<div class="report-card"><h4>Also written</h4><dl>${kv([
        ["emails", `${limit ? Math.min(limit, p.emails) : p.emails} (already-imported ones are skipped)`],
        ["tasks / notes", `${p.tasks} / ${p.notes}`],
        ["read deals / moved mails", `${p.read_deals ?? 0} / ${p.routed_emails ?? 0} (as AI agent)`],
        ["attachments", $("#m-attachments").checked ? `${p.attachments} files` : "not uploaded"],
      ])}</dl></div>`;
    openModal("#mconfirm-modal");
  } catch (e) {
    errBox.textContent = "Could not check the import: " + e.message;
    errBox.classList.remove("hidden");
  } finally {
    btn.disabled = false;
    btn.textContent = "Import to CRM →";
  }
});

$("#mconfirm-cancel").addEventListener("click", () => closeModal("#mconfirm-modal"));
$("#mconfirm-ok").addEventListener("click", async () => {
  closeModal("#mconfirm-modal");
  mailGoToStep(3);
  await mailCommit();
});

// ---- step 3: commit with progress ----

function mailProgress(pct, label) {
  $("#mprogress-fill").style.width = pct + "%";
  $("#mprogress-label").textContent = label;
}

async function mailCommit() {
  $("#mimport-working").classList.remove("hidden");
  $("#mimport-done").classList.add("hidden");
  $("#mimport-error").classList.add("hidden");
  mailProgress(3, "Starting…");
  try {
    const { job_id } = await api("/api/mail/commit", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ attachments: $("#m-attachments").checked, limit_emails: $("#m-limit").value }),
    });
    await mailPoll(job_id);
  } catch (e) {
    mailShowError(e.message);
  }
}

const MAIL_STEPS = {
  starting: [3, "Starting…"], vocabulary: [5, "Preparing fields…"], companies: [8, "Writing companies…"],
  contacts: [15, "Writing contacts…"], products: [25, "Writing products…"], deals: [28, "Writing deals and offers…"],
  emails: [32, "Logging emails…"], routing: [88, "Moving mails to the company/deal a reading found…"],
  titles: [89, "Filling in titles…"], "tasks and notes": [90, "Writing tasks and notes…"], attachments: [93, "Uploading attachments…"],
};

async function mailPoll(jobId) {
  while (true) {
    const job = await api(`/api/commit/${jobId}`);
    if (job.status === "running") {
      const [pct, label] = MAIL_STEPS[job.step] || [4, job.step];
      if (job.step === "emails" && job.total) {
        mailProgress(32 + 58 * (job.done / job.total), `${label} ${job.done}/${job.total}`);
      } else {
        mailProgress(pct, label);
      }
      await new Promise((r) => setTimeout(r, 500));
      continue;
    }
    if (job.status === "error") { mailShowError(job.error); return; }
    mailShowDone(job.result);
    return;
  }
}

function mailShowDone(r) {
  $("#mimport-working").classList.add("hidden");
  $("#mimport-done").classList.remove("hidden");
  const made = Object.values(r.created).reduce((a, b) => a + b, 0);
  const had = Object.values(r.existing).reduce((a, b) => a + b, 0);
  $("#mresult-title").textContent = r.failed_count ? "Import finished with problems" : "Import complete";
  $("#mresult-sub").textContent =
    `${made} records created, ${had} were already there${r.failed_count ? `, ${r.failed_count} failed` : ""}.`;
  const rows = (o) => kv(Object.entries(o).map(([k, v]) => [esc(k), v]));
  const fails = r.failed.map((f) => `<li><b>${esc(f.kind)} ${esc(f.ref)}:</b> ${esc(f.error.slice(0, 160))}</li>`).join("");
  $("#mresult-report").innerHTML =
    `<div class="report-card"><h4>Created</h4><dl>${rows(r.created) || "<dt>nothing new</dt><dd>0</dd>"}</dl></div>` +
    `<div class="report-card"><h4>Already present</h4><dl>${rows(r.existing)}</dl></div>` +
    (Object.keys(r.approvals || {}).length ? `<div class="report-card"><h4>CRM approval inbox</h4><dl>${rows(r.approvals)}</dl></div>` : "") +
    (r.failed_count ? `<div class="report-card"><h4>Failed — ${r.failed_count}</h4><ul class="fail-list">${fails}</ul></div>` : "");
}

function mailShowError(message) {
  $("#mimport-working").classList.add("hidden");
  $("#mimport-error").classList.remove("hidden");
  $("#merror-text").textContent = message;
}

$("#mbtn-import-another").addEventListener("click", async () => {
  await api("/api/mail/clear", { method: "POST" });
  mailGoToStep(1);
});
$("#mbtn-back-to-review").addEventListener("click", () => mailGoToStep(2));

// ---- fact check: mail/verify.py against the archive and the CRM ----

const FAILING = ["NOT FOUND", "MISSING", "CHANGED", "EXTRA", "DUPLICATE"];

function verifyCard(title, fields) {
  const rows = Object.entries(fields).map(([field, counts]) => {
    const bad = Object.keys(counts).some((s) => FAILING.includes(s));
    const text = Object.entries(counts).sort(([a], [b]) => (a === "VERIFIED" ? -1 : b === "VERIFIED" ? 1 : 0))
      .map(([s, n]) => `${n} ${s === "VERIFIED" ? "ok" : s.toLowerCase()}`).join(", ");
    return `<dt class="${bad ? "bad" : ""}">${esc(field)}</dt><dd class="${bad ? "bad" : ""}">${esc(text)}</dd>`;
  }).join("");
  return `<div class="report-card"><h4>${title}</h4><dl>${rows}</dl></div>`;
}

async function mailVerify() {
  $("#mverify-report").innerHTML = "";
  $("#mverify-problems").innerHTML = "";
  $("#mverify-status").textContent = "Starting…";
  openModal("#mverify-modal");
  try {
    const { job_id } = await api("/api/mail/verify", { method: "POST" });
    while (true) {
      const job = await api(`/api/commit/${job_id}`);
      if (job.status === "running") {
        const secs = Math.round(Date.now() / 1000 - job.started_at);
        $("#mverify-status").textContent = `${job.step}… (${secs}s)`;
        await new Promise((r) => setTimeout(r, 700));
        continue;
      }
      if (job.status === "error") { $("#mverify-status").textContent = "Failed: " + job.error; return; }
      const r = job.result;
      $("#mverify-status").textContent = r.problem_count
        ? `${r.problem_count} problem(s) found. Weak and manual items are listed in the CSV, not here.`
        : "Everything checked out.";
      $("#mverify-report").innerHTML =
        (r.fields.source ? verifyCard("Original archive → import", r.fields.source) : "") +
        (r.fields.crm ? verifyCard("Import → CRM", r.fields.crm) : "");
      const items = r.problems.map((f) => `<li><b>${esc(f.status)}</b> ${esc(f.record)} · ${esc(f.field)}: ` +
        `${esc(f.value.slice(0, 60))}${f.detail ? " — " + esc(f.detail.slice(0, 140)) : ""}</li>`).join("");
      $("#mverify-problems").innerHTML =
        `<p class="verify-legend">ok = found where it should be · weak = found, but somewhere else · manual = read by ` +
        `hand from a scan · ${r.has_enrichment ? "" : "<b>no enrichment.json was uploaded, so hand-read deals, tasks " +
        "and notes are not part of this check</b> · "}full list: ${esc(r.csv)}</p>` +
        (items ? `<ul class="fail-list">${items}</ul>` +
          (r.problem_count > r.problems.length ? `<p class="hint">… and ${r.problem_count - r.problems.length} more in the CSV.</p>` : "") : "");
      return;
    }
  } catch (e) {
    $("#mverify-status").textContent = "Failed: " + e.message;
  }
}

document.querySelectorAll(".mbtn-verify").forEach((b) => b.addEventListener("click", mailVerify));
$("#mverify-close").addEventListener("click", () => closeModal("#mverify-modal"));

// ---- init: restore a loaded dataset and the last-used source ----

(async function mailInit() {
  try {
    const d = await api("/api/mail/dataset");
    if (d.loaded) { renderMailReview(d); mailStep = 2; }
  } catch (e) { /* the tab simply starts empty */ }
  let mode = "visits";
  try { mode = sessionStorage.getItem("importMode") || "visits"; } catch (e) { /* default */ }
  showMode(mode);
})();

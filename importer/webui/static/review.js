// Reading review tab: decide what two readings of the mails found (mail/reading.py, READING.md).
// Shares $, api, esc with app.js. Decisions are written into mail/staging/enrichment/<case>.json.

let reviewSlug = null;
let reviewData = null;

const KIND_LABEL = { deal: "Deal", task: "Task", note: "Note", routing: "Mail moved", contact: "Title", document: "Scan" };

function money(minor, cur) {
  return minor === null || minor === undefined ? "—" : `${cur || ""} ${(minor / 100).toLocaleString("en", { minimumFractionDigits: 2 })}`.trim();
}

function fieldsOf(kind, v) {
  if (!v) return [];
  switch (kind) {
    case "deal": return [["number", v.key], ["status", v.status], ["amount", money(v.amount_minor, v.currency)],
      ["delivery / close", v.expected_close || "—"], ["lines", (v.lines || []).map((l) => `${l.qty} × ${l.model || l.text || ""} @ ${money(l.unit_price_minor, v.currency)}`).join("; ") || "—"],
      ["summary", v.summary || "—"]];
    case "task": return [["task", v.subject], ["due", v.due || "—"], ["done", v.done ? "yes" : "no"], ["details", v.body || "—"]];
    case "note": return [["type", v.type], ["subject", v.subject], ["note", v.body]];
    case "routing": return [["belongs to", v.company || "—"], ["deal", v.deal || "—"]];
    case "contact": return [["contact", v.email], ["title", v.title]];
    case "document": return [["document", `${v.type || ""} ${v.number || ""}`], ["date", v.date || "—"],
      ["total", money(v.total_minor, v.currency)], ["file", v.filename], ["says", v.summary || "—"]];
    default: return [];
  }
}

function sideTable(kind, v, other) {
  if (!v) return `<div class="review-side empty">not found by this reading</div>`;
  const theirs = Object.fromEntries(fieldsOf(kind, other));
  return `<dl class="review-side">${fieldsOf(kind, v).map(([k, val]) => {
    const differs = other && theirs[k] !== undefined && theirs[k] !== val;
    return `<dt>${esc(k)}</dt><dd class="${differs ? "differs" : ""}">${esc(val)}</dd>`;
  }).join("")}</dl>`;
}

function evidenceHtml(item) {
  return (item.evidence_shown || []).map((e) => {
    const open = e.doc ? ` <a href="/api/review/case/${reviewSlug}/file/${encodeURIComponent(e.doc)}" target="_blank" rel="noopener">open the scan</a>` : "";
    const where = e.scan ? `<span class="tag">scan — not checkable by code</span>` : (e.verified === false ? `<span class="tag bad">quote not found</span>` : `<span class="tag ok">found in the mail</span>`);
    return `<div class="evidence"><div class="ev-source">${esc(e.field || "")} · ${esc(e.source || "")} ${where}${open}</div>
      <div class="ev-text">${esc(e.before || "")}<mark>${esc(e.quote || "")}</mark>${esc(e.after || "")}</div>
      ${e.note ? `<div class="hint">${esc(e.note)}</div>` : ""}</div>`;
  }).join("");
}

function editForm(item, base) {
  const v = base || {};
  const input = (name, label, value, type = "text") =>
    `<label>${label} <input name="${name}" type="${type}" value="${esc(value ?? "")}"></label>`;
  let html = "";
  switch (item.kind) {
    case "deal":
      html = `<label>status <select name="status">${["open", "won", "lost"].map((s) => `<option${s === v.status ? " selected" : ""}>${s}</option>`).join("")}</select></label>` +
        input("amount", "amount", v.amount_minor != null ? v.amount_minor / 100 : "", "number") + input("currency", "currency", v.currency || "CNY") +
        input("expected_close", "delivery / close (YYYY-MM-DD)", v.expected_close);
      break;
    case "task":
      html = input("subject", "task", v.subject) + input("due", "due (YYYY-MM-DD)", v.due) +
        `<label class="check"><input name="done" type="checkbox"${v.done ? " checked" : ""}> done</label>`;
      break;
    case "note": html = input("subject", "subject", v.subject) + `<label>note <textarea name="body" rows="3">${esc(v.body || "")}</textarea></label>`; break;
    case "routing":
      html = `<label>belongs to <select name="company">${reviewData.companies.map((c) => `<option${c === v.company ? " selected" : ""}>${esc(c)}</option>`).join("")}</select></label>` +
        input("deal", "deal number (optional)", v.deal);
      break;
    case "contact": html = input("title", "title", v.title); break;
    case "document":
      html = input("number", "number", v.number) + input("date", "date (YYYY-MM-DD)", v.date) +
        input("total", "total", v.total_minor != null ? v.total_minor / 100 : "", "number") + input("currency", "currency", v.currency || "CNY");
      break;
  }
  return `<form class="review-edit form-grid" data-id="${esc(item.id)}">${html}
    <div class="modal-actions"><button type="button" class="secondary review-edit-cancel">Cancel</button>
    <button type="submit" class="primary">Save my version</button></div></form>`;
}

function readForm(form, item, base) {
  const v = JSON.parse(JSON.stringify(base || {}));
  const f = Object.fromEntries(new FormData(form).entries());
  const minor = (x) => (x === "" || x === undefined ? null : Math.round(Number(x) * 100));
  switch (item.kind) {
    case "deal": Object.assign(v, { status: f.status, amount_minor: minor(f.amount), currency: f.currency || null, expected_close: f.expected_close || null }); break;
    case "task": Object.assign(v, { subject: f.subject, due: f.due || null, done: form.querySelector("[name=done]").checked }); break;
    case "note": Object.assign(v, { subject: f.subject, body: f.body }); break;
    case "routing": Object.assign(v, { company: f.company, deal: f.deal || null }); break;
    case "contact": Object.assign(v, { title: f.title }); break;
    case "document": Object.assign(v, { number: f.number || null, date: f.date || null, total_minor: minor(f.total), currency: f.currency || null }); break;
  }
  return v;
}

function itemHtml(item) {
  const decided = item.review;
  const state = decided ? `decided: ${decided.choice === "a" ? "reading A" : decided.choice === "b" ? "reading B" : decided.choice === "edit" ? "your version" : "rejected"} — ${esc(decided.by)}, ${esc((decided.at || "").slice(0, 16).replace("T", " "))}`
    : item.state === "needs_review" ? "needs your decision" : item.state === "auto" ? "both readings agree — goes in" : "rejected by the check";
  const problems = [...item.a_problems.map((p) => "A: " + p), ...item.b_problems.map((p) => "B: " + p)];
  const canDecide = item.state !== "rejected";
  return `<article class="review-item ${item.state}${decided ? " decided" : ""}" id="item-${esc(item.id)}">
    <header><span class="tag">${KIND_LABEL[item.kind] || item.kind}</span> <b>${esc((item.a || item.b || {}).subject || (item.a || item.b || {}).key || (item.a || item.b || {}).title || (item.a || item.b || {}).company || item.id)}</b>
      <span class="review-state">${state}</span></header>
    ${item.reasons.length ? `<p class="reasons">${item.reasons.map((r) => `<span class="chip">${esc(r)}</span>`).join(" ")}</p>` : ""}
    ${problems.length ? `<ul class="fail-list">${problems.map((p) => `<li>${esc(p)}</li>`).join("")}</ul>` : ""}
    <div class="review-sides"><div><h5>Reading A</h5>${sideTable(item.kind, item.a, item.b)}</div>
      <div><h5>Reading B</h5>${sideTable(item.kind, item.b, item.a)}</div></div>
    ${evidenceHtml(item)}
    ${canDecide ? `<div class="actions review-buttons" data-id="${esc(item.id)}">
      ${item.a && !item.a_problems.length ? `<button class="secondary" data-choice="a">Accept A</button>` : ""}
      ${item.b && !item.b_problems.length ? `<button class="secondary" data-choice="b">Accept B</button>` : ""}
      <button class="secondary" data-choice="edit">Edit</button>
      <button class="danger" data-choice="reject">Reject</button></div><div class="review-edit-slot"></div>` : ""}
  </article>`;
}

async function reviewLoadCases() {
  $("#review-case").classList.add("hidden");
  $("#review-cases").classList.remove("hidden");
  try { $("#review-by").value = localStorage.getItem("reviewBy") || $("#review-by").value; } catch (e) { /* no storage */ }
  try {
    const { cases } = await api("/api/review/cases");
    const read = cases.filter((c) => c.merged);
    const todo = read.reduce((n, c) => n + c.needs_review, 0);
    $("#review-cases").innerHTML = `<p class="hint">${read.length} of ${cases.length} companies read · ${todo} item(s) wait for you.</p>
      <table class="review-table"><thead><tr><th>Company</th><th>Mails</th><th>Readings</th><th>Need you</th><th>Agreed</th><th>Decided</th><th>Rejected</th></tr></thead>
      <tbody>${cases.map((c) => `<tr class="${c.merged ? "clickable" : "muted"}${c.needs_review ? " todo" : ""}" data-slug="${esc(c.slug)}">
        <td>${esc(c.case)}${c.stale ? ' <span class="tag bad">changed since merge</span>' : ""}</td><td>${c.mails}</td><td>${esc(c.readings) || "—"}</td>
        <td>${c.merged ? c.needs_review : "—"}</td><td>${c.merged ? c.auto : "—"}</td><td>${c.merged ? c.reviewed : "—"}</td><td>${c.merged ? c.rejected : "—"}</td></tr>`).join("")}</tbody></table>`;
    document.querySelectorAll("#review-cases tr.clickable").forEach((tr) => tr.addEventListener("click", () => reviewLoadCase(tr.dataset.slug)));
  } catch (e) {
    $("#review-cases").innerHTML = `<div class="inline-error">${esc(e.message)}</div>`;
  }
}

async function reviewLoadCase(slug, keepScroll) {
  reviewSlug = slug;
  const y = window.scrollY;
  try {
    reviewData = await api(`/api/review/case/${encodeURIComponent(slug)}`);
  } catch (e) {
    $("#review-error").textContent = e.message;
    $("#review-error").classList.remove("hidden");
    return;
  }
  $("#review-error").classList.add("hidden");
  $("#review-cases").classList.add("hidden");
  $("#review-case").classList.remove("hidden");
  const c = reviewData.counts;
  $("#review-title").textContent = reviewData.case;
  $("#review-summary").innerHTML = `${c.needs_review} need you · ${c.auto} agreed · ${c.reviewed} decided · ${c.rejected} rejected by the check` +
    (reviewData.stale ? ` · <b>the mails changed since this was merged: run python -m mail.reading --company "${esc(reviewData.case)}"</b>` : "");
  const all = $("#review-show-all").checked;
  const items = reviewData.items.filter((it) => all || it.state === "needs_review" || it.review);
  $("#review-items").innerHTML = items.length ? items.map(itemHtml).join("") : `<p class="hint">Nothing waits for you here.</p>`;
  if (keepScroll) window.scrollTo(0, y);
}

async function reviewDecide(id, choice, value) {
  const by = $("#review-by").value.trim();
  try { localStorage.setItem("reviewBy", by); } catch (e) { /* no storage */ }
  try {
    await api(`/api/review/case/${encodeURIComponent(reviewSlug)}/decide`, {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id, choice, value, by }),
    });
    await reviewLoadCase(reviewSlug, true);
  } catch (e) {
    $("#review-error").textContent = e.message;
    $("#review-error").classList.remove("hidden");
    window.scrollTo(0, 0);
  }
}

$("#review-items").addEventListener("click", (e) => {
  const btn = e.target.closest(".review-buttons button");
  if (btn) {
    const id = btn.parentElement.dataset.id;
    const item = reviewData.items.find((it) => it.id === id);
    if (btn.dataset.choice === "edit") {
      const slot = btn.parentElement.nextElementSibling;
      slot.innerHTML = editForm(item, item.review?.value || item.a || item.b);
      return;
    }
    reviewDecide(id, btn.dataset.choice, null);
    return;
  }
  if (e.target.closest(".review-edit-cancel")) e.target.closest(".review-edit-slot").innerHTML = "";
});

$("#review-items").addEventListener("submit", (e) => {
  e.preventDefault();
  const form = e.target.closest(".review-edit");
  const item = reviewData.items.find((it) => it.id === form.dataset.id);
  reviewDecide(item.id, "edit", readForm(form, item, item.review?.value || item.a || item.b));
});

$("#review-back").addEventListener("click", reviewLoadCases);
$("#review-show-all").addEventListener("change", () => reviewSlug && reviewLoadCase(reviewSlug));
$("#review-by").addEventListener("change", () => { try { localStorage.setItem("reviewBy", $("#review-by").value.trim()); } catch (e) { /* no storage */ } });
if (!$("#view-review").classList.contains("hidden")) reviewLoadCases();

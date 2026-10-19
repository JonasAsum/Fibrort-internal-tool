// Checks view: plain-language PASS / WARN / FAIL rows from /api/checks. Shares $, api and esc with app.js.

const CHECK_CLASS = { pass: "ok", warn: "warn", fail: "fail" };
const CHECK_WORD = { pass: "Pass", warn: "Warning", fail: "Fail" };

function checkRow(r) {
  const fix = r.fix ? `<div class="fix"><b>How to fix:</b> ${esc(r.fix)}</div>` : "";
  return `<li class="${CHECK_CLASS[r.status]}" data-check="${esc(r.id)}" data-status="${esc(r.status)}">
    <div><div class="t">${esc(r.title)}</div><div class="d">${esc(r.detail)}</div>${fix}</div>
    <span class="badge ${CHECK_CLASS[r.status]}">${CHECK_WORD[r.status]}</span></li>`;
}

let checksBusy = false;

async function checksRun() {
  if (checksBusy) return;
  checksBusy = true;
  const btn = $("#btn-checks-run");
  btn.disabled = true;
  btn.textContent = "Checking…";
  const summary = $("#checks-summary");
  try {
    const d = await api("/api/checks");
    $("#checks-env").innerHTML = d.environment.map(checkRow).join("");
    $("#checks-self").innerHTML = d.selftest.map(checkRow).join("");
    const all = [...d.environment, ...d.selftest];
    const n = (s) => all.filter((r) => r.status === s).length;
    summary.className = `banner ${n("fail") ? "error" : n("warn") ? "warn" : "ok"}`;
    summary.textContent = `${n("pass")} pass, ${n("warn")} warning${n("warn") === 1 ? "" : "s"}, ${n("fail")} fail. `
      + (n("fail") ? "Fix the failing lines first." : n("warn") ? "Warnings do not stop an import." : "Everything this site needs is in place.");
  } catch (e) {
    summary.className = "banner error";
    summary.textContent = "The checks could not run: " + e.message;
  } finally {
    btn.disabled = false;
    btn.textContent = "Check again";
    checksBusy = false;
  }
}
$("#btn-checks-run").addEventListener("click", checksRun);

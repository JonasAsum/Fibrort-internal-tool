import os
import re

REAL_ROOT = "/Users/jonas/Mac Documents/Shanghai Company/CRM Excel Lists 2026-10-11"
LABELS = {"found": "visits found", "visits": "visits to import", "unsure": "unsure (not imported)",
          "companies": "companies", "contacts": "contacts", "tasks": "tasks", "weeks": "weekly reports"}


def real_files(*prefixes):
    out = []
    for prefix in prefixes:
        for dirpath, _, names in os.walk(REAL_ROOT):
            if not os.path.relpath(dirpath, REAL_ROOT).startswith(prefix):
                continue
            out += [os.path.join(dirpath, n) for n in names
                    if n.lower().endswith((".xlsx", ".xls")) and not n.startswith(("~$", "."))]
    return sorted(out)


def real_id(path):
    return os.path.relpath(path, REAL_ROOT)


def fresh_page(page, base):
    """The upload step with no earlier upload on the server (a reload would show that one)."""
    page.request.post(base + "/api/clear")
    page.goto(base)


def upload(page, base, paths):
    """Chooses the files in the real page's file input; returns "review" or "error" once the page settled."""
    fresh_page(page, base)
    with page.expect_response(lambda r: r.url.endswith("/api/upload"), timeout=120_000):
        page.set_input_files("#file-input", [str(p) for p in paths])
    page.wait_for_function(
        "document.querySelector('#step-2.active') || !document.querySelector('#upload-error').classList.contains('hidden')",
        timeout=120_000)
    return "review" if page.locator("#step-2.active").count() else "error"


def stats(page):
    """The numbers the review step shows, as ints keyed like LABELS."""
    shown = {}
    for stat in page.locator("#summary-grid .stat").all():
        shown[stat.locator(".label").inner_text().strip()] = int(stat.locator(".n").inner_text())
    return {key: shown[label] for key, label in LABELS.items()}


def page_text(page):
    """Everything a person can read on the visible page."""
    return re.sub(r"[ \t]+", " ", page.locator("body").inner_text())


def entry_rows(page):
    """The front door's table as a list of dicts: file, what, target, rows, badge, outcome."""
    rows = []
    for tr in page.locator("#entry-body tr").all():
        cells = tr.locator("td")
        rows.append({"file": cells.nth(0).locator(":scope").inner_text().split("\n")[0].strip(),
                     "what": cells.nth(1).inner_text().strip(), "target": cells.nth(2).inner_text().strip(),
                     "rows": cells.nth(3).inner_text().strip(), "badge": tr.locator(".badge").first.inner_text().strip(),
                     "outcome": tr.locator(".outcome").inner_text().strip(), "status": tr.get_attribute("data-status")})
    return rows


def crm_counts(page, base):
    """What the in-memory CRM of tests/e2e/site_server.py holds."""
    return page.request.get(base + "/__test/crm").json()


def watch_console(page):
    """A list that fills with every console error and uncaught page error from now on."""
    errors = []
    page.on("console", lambda m: errors.append(m.text) if m.type == "error" else None)
    page.on("pageerror", lambda e: errors.append(str(e)))
    return errors


def shot(page, name):
    """Saves a screenshot under $E2E_SHOTS when it is set (the tests never need one)."""
    import os
    folder = os.environ.get("E2E_SHOTS")
    if folder:
        os.makedirs(folder, exist_ok=True)
        page.screenshot(path=os.path.join(folder, name + ".png"), full_page=True)


def wait_import_result(page, timeout=300_000):
    page.wait_for_selector("#import-done:not(.hidden), #import-error:not(.hidden)", timeout=timeout)
    assert page.locator("#import-done:not(.hidden)").count(), page.locator("#error-text").inner_text()

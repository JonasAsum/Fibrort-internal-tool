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

"""Generic truths about the extraction output (real-archive truths live in the gitignored checks_private.py). extract.py runs them after every extraction; a failure exits non-zero.

    .venv/bin/python -m mail.checks          # re-check staging/extracted.json without re-extracting

Add a line here whenever you correct something by hand, so it can never silently come back.
"""
import json
import os
import re
import sys

try:
    from .checks_private import run as private_run
except ImportError:  # not in the public repo: real names and addresses
    private_run = None

HERE = os.path.dirname(os.path.abspath(__file__))

BAD_COMPANY_WORD = re.compile(r"我|您|请|尊敬|菲博|FIBRO|Fibro|Läpple|Laepple")
BAD_PERSON = re.compile(r"(\d|^[a-z0-9._-]+$|^(info|sales\w*|service|admin|ctc)$)")


def run(out: dict) -> list[str]:
    rows = out["relevance"]
    fails = []

    def check(ok: bool, what: str) -> None:
        print(("  ok    " if ok else "  FAIL  ") + what)
        if not ok:
            fails.append(what)

    print("checks:")
    for c in out["companies"]:
        if not c["keep"] or c.get("relationship") == "FIBRO group":
            continue
        for v in (c["name"], c.get("legal_name"), c.get("legal_name_en")):
            if v and BAD_COMPANY_WORD.search(v):
                check(False, f"company {c.get('key')} has a name that is a sentence or our own: {v!r}")
    bad_people = [c["full_name"] for c in out["contacts"]
                  if c["keep"] and not c["mailbox"] and BAD_PERSON.match(c["full_name"])]
    check(not bad_people, f"no contact is named by a bare handle {bad_people[:6]}")
    junk_titles = [c["title"] for c in out["contacts"] if c["keep"] and c["title"] and re.search(
        r"Dr\. Michael|Board of|Directors?:|Sincerely|</|please|if you|如有|页面", c["title"], re.I)]
    check(not junk_titles, f"no footer or disclaimer line is used as a job title {junk_titles[:3]}")
    robots = [c["full_name"] for c in out["contacts"] if c["keep"] and not c["mailbox"]
              and re.match(r"(srm|erp|rpa|system)[-._]", c["key"])]
    check(not robots, f"system mailboxes are not people {robots}")
    banners = [m["subject"][:40] for m in out["messages"]
               if re.search(r"LearnAboutSenderIdentification|erhalten nicht häufig E-Mails|don't often get email", m["top"])]
    check(not banners, f"Outlook's 'you don't often get email from' warning is not kept as the mail's text {banners[:3]}")
    if private_run:  # checks on named people and mailboxes live in checks_private.py (gitignored)
        private_run(out, check, rows)
    print(f"{len(fails)} check(s) failed" if fails else "all checks passed")
    return fails


if __name__ == "__main__":
    with open(os.path.join(HERE, "staging", "extracted.json")) as f:
        sys.exit(1 if run(json.load(f)) else 0)

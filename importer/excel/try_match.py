"""Try the company matcher on names you type, against the real company list.

    .venv/bin/python -m excel.try_match "Wuhan Jiashengyuan Mechanical and Electrical Equipment Co., Ltd."
    .venv/bin/python -m excel.try_match            # then type one name per line
"""
import sys

from .companies import CJK_RE, Master, _cn_core, _en_core, _places, _prepare
from .run import IMPORTER_DIR, best_match

VERDICT = {"exact": "SAME (exact name)", "code": "SAME (customer code)", "contains": "SAME (one name inside the other)",
           "pinyin": "SAME (pinyin)", "fuzzy": "SAME (close spelling)", "new": "NEW company (nothing like it in the list)",
           "review": "ASK A PERSON (unsure)", "ambiguous": "ASK A PERSON (several exact matches)", "none": "NOTHING TO MATCH ON"}


def show(master: Master, name: str) -> None:
    text = _prepare(name)
    m = best_match(master, name, "")
    print(f"\n{name}")
    print(f"  compared as: Chinese core {_cn_core(''.join(CJK_RE.findall(text))) or '-'!r}, "
          f"English core {_en_core(CJK_RE.sub(' ', text)) or '-'!r}, places {sorted(sum(map(list, _places(text)), [])) or '-'}")
    print(f"  -> {VERDICT.get(m.method, m.method)}  (score {m.score:.2f})")
    if m.code:
        c = master.get(m.code)
        print(f"     {m.code}  {c['cn'] or c['en']}  {c['en'] if c['cn'] else ''}  USCC {c['uscc'] or '-'}")
    for code, label, score in m.candidates:
        c = master.get(code) or {}
        print(f"     candidate {score:.0%}  {code}  {label}  {c.get('en', '')}  USCC {c.get('uscc') or '-'}")


if __name__ == "__main__":
    master = Master.load(IMPORTER_DIR)
    print(f"{master.size} companies loaded")
    for n in (sys.argv[1:] or (line.strip() for line in sys.stdin)):
        if n:
            show(master, n)

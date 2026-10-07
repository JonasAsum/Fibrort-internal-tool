"""Layer 1c: one company per organisation domain, named from countable evidence only.

A name is copied from the data, never composed:
  - Chinese legal name: a whole signature line ending in 有限公司/有限责任公司/股份有限公司/集团, in a mail the company
    itself sent, that is either tied to the company's domain (the same signature block shows the domain, as in
    "示例智能技术（上海）有限公司 / www.gamma.com") or seen in at least two of its mails.
  - English legal name: the same rule for "... Co., Ltd / GmbH / Inc" lines, plus contract parties.
  - Two candidates with equal evidence: no name, listed as a conflict for review - unless one mail says the company
    was renamed (更名为 X).
Names that appear in FIBRO staff signatures (菲博精密部件..., our own) can never name an outside company.
"""
import re
from collections import Counter, defaultdict

from .people import FREEMAIL, clean_company, domain_company_name, is_group, is_internal, org_domain

ZH_LINE = re.compile(r"^(?:[A-Za-z0-9&·\-]{1,20}[-\s]?)?[一-鿿（）()·]{4,38}(?:有限责任公司|股份有限公司|有限公司|集团)$")
ZH_SENTENCE = re.compile(r"[我您请的贵，。：:；!！?？]|尊敬|附件|提供|更名|以下|如下|关于|开具")
EN_LINE = re.compile(r"(Co\.,?\s*Ltd\.?|Co\.\s*Ltd|Ltd\.?|GmbH(?: & Co\.? KG)?|Inc\.?|LLC|Corporation|Corp\.?|\bAG|S\.A\.|"
                     r"S\.L\.|B\.V\.|S\.p\.A\.|Pvt\.? Ltd\.?)\s*$")
LEGAL_FORM = re.compile(r"[\s,.]*(Co\.?,?\s*Ltd\.?|Co\.|Ltd\.?|GmbH(?: & Co\.? KG)?|Inc\.?|LLC|Corporation|Corp\.?|AG|"
                        r"S\.A\.|S\.L\.|B\.V\.|S\.p\.A\.|Pvt\.? Ltd\.?)\.?$", re.I)
OURS = re.compile(r"菲博|FIBRO|Fibro|Läpple|Laepple|莱普", re.I)


def _sig_lines(top: str) -> list[str]:
    lines = [re.sub(r"\s+", " ", l).strip(" \t|/") for l in top.split("\n")]
    return [l for l in lines if l][-15:]


def zh_candidates(top: str) -> list[str]:
    out = []
    for l in _sig_lines(top):
        t = re.sub(r"\s+", "", l)
        t = re.sub(r"^(公司名称|公司|单位|Company)[:：]", "", t)
        if ZH_LINE.fullmatch(t) and not ZH_SENTENCE.search(t):
            out.append(t)
    return out


def en_candidates(top: str) -> list[str]:
    out = []
    for l in _sig_lines(top):
        if len(l) > 80 or not EN_LINE.search(l) or re.search(r"[一-鿿]", l):
            continue
        c = clean_company(l)
        if c and len(re.findall(r"[A-Za-z]", c)) >= 4 and not re.search(r"©|copyright|registered|confidential", c, re.I):
            out.append(c)
    return out


def _label(org: str) -> str:
    return org.split(".")[0]


def _tied(top: str, org: str) -> bool:
    block = "\n".join(_sig_lines(top)).lower()
    return _label(org) in block


def _pick(cands: dict[str, list], renamed: set[str]) -> tuple[str | None, list[str]]:
    """cands: name -> [messages, tied messages]. Returns (name, conflicts)."""
    scored = sorted(((t, n, name) for name, (n, t) in cands.items() if t >= 1 or n >= 2), reverse=True)
    if not scored:
        return None, []
    for t, n, name in scored:
        if name in renamed:
            return name, []
    if len(scored) == 1:
        return scored[0][2], []
    (t1, n1, top), (t2, n2, _) = scored[0], scored[1]
    if t1 >= 1 and n1 >= 3 * n2 and t1 > t2:  # clearly the company's own name, the other a quoted one-off
        return top, []
    return None, [name for _, _, name in scored]


def strip_legal_form(name: str) -> str:
    prev = None
    while prev != name:
        prev, name = name, LEGAL_FORM.sub("", name).strip(" ,.")
    return name or prev


def build_companies(messages: list[dict], contacts: list[dict], folder_companies: list[dict],
                    docs: dict[str, dict]) -> dict[str, dict]:
    """Returns {key: company}; key = organisation domain, or 'folder:<name>' for a folder with no domain,
    or 'party:<name>' for a contract party nobody mailed from."""
    internal_zh, internal_en = Counter(), Counter()
    zh, en = defaultdict(lambda: defaultdict(lambda: [0, 0])), defaultdict(lambda: defaultdict(lambda: [0, 0]))
    renamed: dict[str, set] = defaultdict(set)
    senders: dict[str, set] = defaultdict(set)  # who writes from a domain
    signers: dict[tuple, set] = defaultdict(set)  # who signs with a given legal name
    for m in messages:
        a = m["from"]["email"]
        if not a:
            continue
        dom = a.split("@")[-1]
        if is_internal(a) or is_group(dom):
            # a forward by a colleague carries the customer's signature too, so a single sighting proves nothing
            internal_zh.update(set(zh_candidates(m["top"])))
            internal_en.update(set(en_candidates(m["top"])))
            continue
        if dom in FREEMAIL:
            continue
        org = org_domain(dom)
        tied = _tied(m["top"], org)
        senders[org].add(a)
        for name in set(zh_candidates(m["top"])):
            zh[org][name][0] += 1
            zh[org][name][1] += tied
            signers[(org, name)].add(a)
        for name in set(en_candidates(m["top"])):
            en[org][name][0] += 1
            en[org][name][1] += tied
            signers[(org, name)].add(a)
        for x in re.findall(r"更名为\s*([一-鿿（）()]{4,38}?(?:有限责任公司|股份有限公司|有限公司|集团))", m["top"]):
            renamed[org].add(x)

    ours_zh = {n for n, k in internal_zh.items() if k >= 3}
    ours_en = {n for n, k in internal_en.items() if k >= 3}

    # contract / order parties: counterparties are English names ("Suzhou Acme Automation Technology Co.,Ltd")
    party_of: dict[str, set] = defaultdict(set)
    parties_by_org: dict[str, set] = defaultdict(set)
    orgs = {c["org"] for c in contacts if c.get("org")}
    for m in messages:
        ext = {org_domain(p["email"].split("@")[-1]) for p in [m["from"]] + m["to"] + m["cc"]
               if p["email"] and not is_internal(p["email"]) and not is_group(p["email"].split("@")[-1])
               and p["email"].split("@")[-1] not in FREEMAIL}
        for a in m["attachments"]:
            parsed = (docs.get(a["sha256"]) or {}).get("parsed") or {}
            party = (parsed.get("counterparty") or "").strip()
            if not party or OURS.search(party):
                continue
            folded = re.sub(r"[^a-z0-9]", "", party.lower())
            match = {o for o in orgs if len(_label(o)) >= 4 and _label(o).replace("-", "") in folded}
            # a filed folder without an outside domain ('Eta') owns the party whose name contains it
            # ('Hebei Eta Machinery Manufacturing Co., Ltd.')
            match |= {"folder:" + fc["name"] for fc in folder_companies
                      if not fc.get("domains") and fc.get("customer_type") != "internal"
                      and len(re.sub(r"[^a-z0-9]", "", fc["name"].lower())) >= 4
                      and re.sub(r"[^a-z0-9]", "", fc["name"].lower()) in folded}
            if not match and len(ext) == 1:
                match = ext
            for o in match:
                party_of[o].add(party)
                parties_by_org[o].add(party)
            if not match:
                party_of["party:" + party].add(party)

    out: dict[str, dict] = {}
    folder_by_domain = {d: fc for fc in folder_companies for d in fc.get("domains", [])}
    for fc in folder_companies:
        if fc.get("customer_type") == "internal":
            continue
        key = fc["domains"][0] if fc.get("domains") else "folder:" + fc["name"]
        out[key] = {**fc, "key": key, "folder_name": fc["name"]}
    for org in sorted(orgs | set(zh) | set(en)):
        if org in folder_by_domain and folder_by_domain[org]["domains"][0] != org:
            continue  # a second domain of a folder company (cn.epsilon.com and epsilon.com fold already)
        out.setdefault(org, {"name": None, "key": org, "customer_type": None, "business_line": None,
                             "domains": [org], "folders": [], "messages": 0, "origin": "domain",
                             "folder_name": None})
    for key in [k for k in party_of if k.startswith("party:")]:
        party = key[6:]
        out.setdefault(key, {"name": None, "key": key, "customer_type": None, "business_line": None, "domains": [],
                             "folders": [], "messages": 0, "origin": "document", "folder_name": None})

    for key, co in out.items():
        doms = co["domains"] or []
        zc, ec = defaultdict(lambda: [0, 0]), defaultdict(lambda: [0, 0])
        # a large domain is a group of legal entities (Delta Nanjing, Suzhou, Taicang ...): one person's
        # signature names their entity, not the whole company, so there a name needs two signers
        big = sum(len(senders[d]) for d in doms) >= 8

        def representative(d: str, n: str) -> bool:
            return not big or len(signers[(d, n)]) >= 2

        for d in doms:
            for n, v in zh[d].items():
                if n not in ours_zh and not OURS.search(n) and representative(d, n):
                    zc[n][0] += v[0]
                    zc[n][1] += v[1]
            for n, v in en[d].items():
                if n not in ours_en and not OURS.search(n) and representative(d, n):
                    ec[n][0] += v[0]
                    ec[n][1] += v[1]
        for d in doms + [key]:
            for p in party_of.get(d, ()):
                ec[p][0] += 2  # a signed contract names the party: as strong as two signatures
                ec[p][1] += 1
        ren = set().union(*(renamed[d] for d in doms)) if doms else set()
        co["legal_name"], zh_conf = _pick(zc, ren)
        co["legal_name_en"], en_conf = _pick(ec, set())
        co["name_conflicts"] = {"zh": zh_conf, "en": en_conf}
        co["name_evidence"] = {"zh": {n: v for n, v in zc.items()}, "en": {n: v for n, v in ec.items()}}
        if co.get("folder_name"):
            co["name"] = co["folder_name"]
        elif co["legal_name_en"]:
            co["name"] = strip_legal_form(co["legal_name_en"])
        elif key.startswith("party:"):
            co["name"] = strip_legal_form(key[6:])
        else:
            co["name"] = domain_company_name(key)
        co["group"] = any(is_group(d) for d in doms)
        co["parties"] = sorted(set().union(*(parties_by_org[d] for d in doms + [key]))
                               | ({key[6:]} if key.startswith("party:") else set()))

    # two companies may not share a display name: the domain tells them apart
    seen = Counter(co["name"].lower() for co in out.values())
    for co in out.values():
        if seen[co["name"].lower()] > 1 and co["domains"]:
            co["name"] = f"{co['name']} ({co['domains'][0]})"
    return out

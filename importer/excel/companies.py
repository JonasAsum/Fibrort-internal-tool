"""Customer master index and company-name matching (see excel/SPEC.md, "Companies")."""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import xlrd
from rapidfuzz import fuzz

MASTER_GLOB = "mail/staging/files/*/*Customer_List_20260104*.xlsx"
ERP_DIR_NAME = "Excel Files from Sales Meetings"  # sibling of the importer dir
ERP_GLOB = "*Customer List_*Region*.xls"

CODE_LABELS = {"customer code", "客户编码"}
CN_LABELS = {"cn name", "客户名称"}
EN_LABELS = {"en name"}
SHORT_LABELS = {"客户简称"}
HEADER_SCAN_ROWS = 5

AUTO_SCORE = 0.9
FUZZY_MIN = 92
FUZZY_GAP = 5

_BRACKET_PREFIX_RE = re.compile(r"^\s*[(\[（【][^)\]）】]*[)\]）】]")
_PUNCT_RE = re.compile(r"[^0-9a-z一-鿿]+")
_SUFFIX_RE = re.compile(
    r"有限责任公司|股份有限公司|有限公司|股份|集团|(?<![a-z])(?:co ltd|gmbh|inc|ltd)(?![a-z])"
)


def norm(text) -> str:
    """Matching key: NFKC, lowercase, no leading city bracket, no legal suffix, no punctuation or spaces."""
    if text is None:
        return ""
    s = unicodedata.normalize("NFKC", str(text)).lower()
    s = _BRACKET_PREFIX_RE.sub(" ", s)
    s = _PUNCT_RE.sub(" ", s)
    s = _SUFFIX_RE.sub(" ", s)
    return "".join(s.split())


def _label(value) -> str:
    s = unicodedata.normalize("NFKC", str(value or "")).replace("\n", " ").lower()
    return " ".join(s.split()).strip("：: ")


def _cell(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def parse_rows(rows, source: str) -> list[dict]:
    """Find the header row (first 5 rows) and return one record per data row. No header -> []."""
    for h, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        cols: dict[str, int] = {}
        for idx, value in enumerate(row):
            lab = _label(value)
            if lab in CODE_LABELS:
                cols.setdefault("code", idx)
            elif lab in CN_LABELS:
                cols.setdefault("cn", idx)
            elif lab in EN_LABELS:
                cols.setdefault("en", idx)
            elif lab in SHORT_LABELS:
                cols.setdefault("short", idx)
        if "code" not in cols or not ({"cn", "en", "short"} & cols.keys()):
            continue
        out = []
        for row in rows[h + 1:]:
            rec = {"source": source}
            for field_name in ("code", "cn", "en", "short"):
                i = cols.get(field_name)
                rec[field_name] = _cell(row[i]) if i is not None and i < len(row) else ""
            if any(rec[k] for k in ("code", "cn", "en", "short")):
                out.append(rec)
        return out
    return []


@dataclass
class Company:
    code: str
    cn: str = ""
    en: str = ""
    short: str = ""
    aliases: list[str] = field(default_factory=list)

    def names(self) -> list[str]:
        return [n for n in [self.cn, self.en, self.short, *self.aliases] if n]

    def absorb(self, rec: dict) -> None:
        for name_field in ("cn", "en", "short"):
            value = rec.get(name_field, "")
            if not value:
                continue
            if not getattr(self, name_field):
                setattr(self, name_field, value)
            elif value != getattr(self, name_field) and value not in self.aliases:
                self.aliases.append(value)


@dataclass(frozen=True)
class Match:
    code: str | None
    cn: str
    en: str
    score: float
    method: str  # exact | contains | fuzzy | ambiguous | none
    candidates: list = field(default_factory=list)  # [(code, name, score)], only when score < AUTO_SCORE


def _read_xlsx(path: Path) -> list[list]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return [list(rows) for rows in (ws.iter_rows(values_only=True) for ws in wb.worksheets)]
    finally:
        wb.close()


def _read_xls(path: Path) -> list[list]:
    wb = xlrd.open_workbook(str(path))
    return [[sh.row_values(r) for r in range(sh.nrows)] for sh in wb.sheets()]


class Master:
    def __init__(self, companies: list[Company], sources: dict[str, int], skipped: list[str]):
        self.companies = companies
        self.sources = sources
        self.skipped = skipped
        self._exact: dict[str, list[int]] = {}
        self._keys: list[tuple[int, str]] = []
        for idx, company in enumerate(companies):
            for key in {norm(n) for n in company.names()} - {""}:
                self._exact.setdefault(key, []).append(idx)
                self._keys.append((idx, key))

    @property
    def size(self) -> int:
        return len(self.companies)

    @classmethod
    def from_records(cls, records: list[dict]) -> "Master":
        companies: list[Company] = []
        by_code: dict[str, Company] = {}
        sources: dict[str, int] = {}
        for rec in records:
            sources[rec["source"]] = sources.get(rec["source"], 0) + 1
            code = rec.get("code", "")
            if code and code in by_code:
                by_code[code].absorb(rec)
                continue
            company = Company(code=code, cn=rec.get("cn", ""), en=rec.get("en", ""), short=rec.get("short", ""))
            companies.append(company)
            if code:
                by_code[code] = company
        return cls(companies, sources, [])

    @classmethod
    def load(cls, importer_dir) -> "Master":
        root = Path(importer_dir).resolve()
        records: list[dict] = []
        skipped: list[str] = []
        sources = {"master": 0, "erp": 0}
        files = [(p, "master", _read_xlsx) for p in sorted(root.glob(MASTER_GLOB))]
        erp_dir = root.parent / ERP_DIR_NAME
        if erp_dir.is_dir():
            files += [(p, "erp", _read_xls) for p in sorted(erp_dir.rglob(ERP_GLOB))
                      if not p.name.startswith("~$")]
        for path, source, reader in files:
            try:
                recs = []
                for rows in reader(path):
                    recs.extend(parse_rows(rows, source))
            except Exception:
                skipped.append(path.name)
                continue
            records.extend(recs)
            sources[source] += len(recs)
        master = cls.from_records(records)
        master.sources = sources
        master.skipped = skipped
        return master

    def summary(self) -> dict:
        return {
            "records": dict(self.sources),
            "companies": self.size,
            "with_code": sum(1 for c in self.companies if c.code),
            "with_cn": sum(1 for c in self.companies if c.cn),
            "with_en": sum(1 for c in self.companies if c.en),
            "skipped": list(self.skipped),
        }

    def match(self, name) -> Match:
        q = norm(name)
        if not q:
            return Match(None, "", "", 0.0, "none")

        hits = sorted(set(self._exact.get(q, [])))
        if len(hits) == 1:
            return self._found(hits[0], 1.0, "exact")
        if len(hits) > 1:
            return Match(None, "", "", 0.0, "ambiguous",
                         [self._cand(i, 1.0) for i in hits[:3]])

        if len(q) >= 4:
            contains = {i for i, key in self._keys if len(key) >= 4 and (q in key or key in q)}
            if len(contains) == 1:
                return self._found(contains.pop(), AUTO_SCORE, "contains")

        best: dict[int, float] = {}
        for i, key in self._keys:
            s = max(fuzz.token_set_ratio(q, key), fuzz.partial_ratio(q, key))
            if s > best.get(i, -1.0):
                best[i] = s
        ranked = sorted(best.items(), key=lambda kv: (-kv[1], kv[0]))
        top = [self._cand(i, s / 100) for i, s in ranked[:3]]
        if ranked:
            top_idx, top_s = ranked[0]
            second = ranked[1][1] if len(ranked) > 1 else 0.0
            if top_s >= FUZZY_MIN and top_s - second >= FUZZY_GAP:
                return self._found(top_idx, round(top_s * 0.95 / 100, 4), "fuzzy", top)
        return Match(None, "", "", 0.0, "none", top)

    def _found(self, idx: int, score: float, method: str, candidates=None) -> Match:
        c = self.companies[idx]
        if score >= AUTO_SCORE:
            candidates = []
        elif candidates is None:
            candidates = [self._cand(idx, score)]
        return Match(c.code or None, c.cn, c.en, score, method, candidates)

    def _cand(self, idx: int, score: float) -> tuple:
        c = self.companies[idx]
        return (c.code, c.cn or c.en or c.short, round(score, 4))

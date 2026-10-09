"""Customer master index and company-name matching (see excel/SPEC.md, "Companies")."""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import xlrd
from pypinyin import lazy_pinyin
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
FUZZY_MIN = 90
FUZZY_GAP = 5
REVIEW_MIN = 60
CN_MIN = 2  # shortest CJK core allowed to match by itself
EN_MIN = 4  # shortest latin core allowed to match by itself
PINYIN_PREFIX_MIN = 6

_BRACKET_PREFIX_RE = re.compile(r"^\s*[(\[（【][^)\]）】]*[)\]）】]")
_PUNCT_RE = re.compile(r"[^0-9a-z一-鿿]+")
_SUFFIX_RE = re.compile(
    r"有限责任公司|股份有限公司|有限公司|股份|集团|(?<![a-z])(?:co ltd|gmbh|inc|ltd)(?![a-z])"
)
CJK_RE = re.compile(r"[㐀-䶿一-鿿豈-﫿]+")
CODE_RE = re.compile(r"(?<![A-Za-z0-9])0*([A-Za-z]{2,4}\d{5,})(?!\d)")

CN_LEGAL = ("有限责任公司", "股份有限公司", "有限公司", "股份", "集团")
CN_GENERIC = ("责任", "科技", "技术", "精密", "机械", "设备", "装备", "智能", "自动化", "数控", "制造",
              "工业", "发展", "贸易", "实业", "机床", "国际")
CN_STRIP = tuple(sorted(CN_LEGAL + CN_GENERIC, key=len, reverse=True))

LAT_STRIP = frozenset({
    "co", "ltd", "limited", "inc", "corp", "corporation", "gmbh", "llc",
    "company", "group", "technology", "tech", "precision", "machinery", "machine",
    "equipment", "intelligent", "automation", "cnc", "manufacturing", "industrial",
    "industry", "international",
})

CN_PLACES = (
    "上海", "北京", "天津", "重庆", "宁夏", "银川", "沈阳", "大连", "广州", "深圳", "东莞", "佛山",
    "珠海", "中山", "惠州", "江门", "汕头", "杭州", "宁波", "温州", "嘉兴", "绍兴", "台州", "金华",
    "南京", "苏州", "无锡", "常州", "昆山", "常熟", "江阴", "张家港", "南通", "扬州", "徐州", "青岛",
    "济南", "烟台", "潍坊", "临沂", "武汉", "长沙", "成都", "西安", "郑州", "合肥", "南昌", "福州",
    "厦门", "泉州", "哈尔滨", "长春", "石家庄", "唐山", "保定", "太原", "呼和浩特", "兰州", "西宁",
    "乌鲁木齐", "拉萨", "昆明", "贵阳", "南宁", "海口", "香港", "澳门", "台湾", "江苏", "浙江",
    "广东", "山东", "河北", "河南", "湖北", "湖南", "四川", "陕西", "辽宁", "吉林", "黑龙江", "安徽",
    "江西", "福建", "山西", "云南", "贵州", "广西", "海南", "甘肃", "青海", "新疆", "西藏", "内蒙古",
)
CN_PLACES_SORTED = tuple(sorted(set(CN_PLACES), key=len, reverse=True))
EN_PLACES = frozenset({
    "shanghai", "beijing", "tianjin", "chongqing", "ningxia", "yinchuan", "shenyang", "dalian",
    "guangzhou", "shenzhen", "dongguan", "foshan", "zhuhai", "zhongshan", "huizhou", "jiangmen",
    "shantou", "hangzhou", "ningbo", "wenzhou", "jiaxing", "shaoxing", "taizhou", "jinhua", "nanjing",
    "suzhou", "wuxi", "changzhou", "kunshan", "changshu", "jiangyin", "zhangjiagang", "nantong",
    "yangzhou", "xuzhou", "qingdao", "jinan", "yantai", "weifang", "linyi", "wuhan", "changsha",
    "chengdu", "xian", "zhengzhou", "hefei", "nanchang", "fuzhou", "xiamen", "quanzhou", "harbin",
    "changchun", "shijiazhuang", "tangshan", "baoding", "taiyuan", "hohhot", "lanzhou", "xining",
    "urumqi", "lhasa", "kunming", "guiyang", "nanning", "haikou", "hong kong", "macau", "taiwan",
    "jiangsu", "zhejiang", "guangdong", "shandong", "hebei", "henan", "hubei", "hunan", "sichuan",
    "shaanxi", "liaoning", "jilin", "heilongjiang", "anhui", "jiangxi", "fujian", "shanxi", "yunnan",
    "guizhou", "guangxi", "hainan", "gansu", "qinghai", "xinjiang", "tibet", "inner mongolia",
}) | {"".join(lazy_pinyin(p)) for p in CN_PLACES}


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


def _prepare(text) -> str:
    s = unicodedata.normalize("NFKC", str(text or "")).replace("\n", " ")
    return _BRACKET_PREFIX_RE.sub(" ", s).strip()


def _canon_code(code: str) -> str:
    return code.upper().lstrip("0")


def _extract_codes(text: str) -> tuple[list[str], str]:
    codes = [_canon_code(m.group(1)) for m in CODE_RE.finditer(text)]
    return codes, CODE_RE.sub(" ", text)


def _near_contains(a: str, b: str) -> bool:
    """One core is inside the other and covers at least 80% of it (no short word inside a long name)."""
    short, long_ = sorted((a, b), key=len)
    return short in long_ and len(short) * 5 >= len(long_) * 4


def _pinyin(text: str) -> str:
    return "".join(lazy_pinyin(text)).lower()


def _cn_core(cjk: str) -> str:
    """CJK name without legal suffixes, generic industry words and one leading province/city."""
    s = cjk
    for word in CN_STRIP:
        s = s.replace(word, "")
    stripped = True
    while stripped:
        stripped = False
        for place in CN_PLACES_SORTED:
            if s.startswith(place) and len(s) - len(place) >= CN_MIN:
                s = s[len(place):]
                stripped = True
                break
    return s


def _en_core(latin: str) -> str:
    """Latin name letters/digits only, without legal suffixes, generic words and one leading place."""
    toks = [t for t in re.sub(r"[^0-9a-z]+", " ", latin.lower()).split() if t not in LAT_STRIP]
    for n in (2, 1):
        if len(toks) > n and " ".join(toks[:n]) in EN_PLACES and len("".join(toks[n:])) >= EN_MIN:
            toks = toks[n:]
            break
    return "".join(toks)


class Master:
    def __init__(self, companies: list[Company], sources: dict[str, int], skipped: list[str]):
        self.companies = companies
        self.sources = sources
        self.skipped = skipped
        self._exact: dict[str, list[int]] = {}
        self._by_code: dict[str, int] = {}
        self._cn: list[tuple[int, str]] = []  # CJK cores of CN / short names
        self._en: list[tuple[int, str]] = []  # latin cores of EN names
        self._py: list[tuple[int, str]] = []  # pinyin of CN cores
        for idx, company in enumerate(companies):
            if company.code:
                self._by_code.setdefault(_canon_code(company.code), idx)
            for name in company.names():
                key = norm(name)
                if key:
                    self._exact.setdefault(key, []).append(idx)
                text = _prepare(name)
                cjk = "".join(CJK_RE.findall(text))
                cn = _cn_core(cjk) if cjk else ""
                if cn:
                    self._cn.append((idx, cn))
                    self._py.append((idx, _pinyin(cn)))
                en = _en_core(CJK_RE.sub(" ", text))
                if en:
                    self._en.append((idx, en))

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
        text = _prepare(name)
        codes, text = _extract_codes(text)
        for code in codes:
            idx = self._by_code.get(code)
            if idx is not None:
                return self._found(idx, 1.0, "code")

        q = norm(text)
        cjk = "".join(CJK_RE.findall(text))
        cq = _cn_core(cjk) if cjk else ""
        eq = _en_core(CJK_RE.sub(" ", text))
        if not cq and not eq or self.size == 0:
            return Match(None, "", "", 0.0, "none")
        usable_cn = len(cq) >= CN_MIN
        usable_en = len(eq) >= EN_MIN

        hits = sorted(set(self._exact.get(q, []))) if q else []
        if len(hits) == 1:
            return self._found(hits[0], 1.0, "exact")
        if len(hits) > 1:
            return Match(None, "", "", 0.0, "ambiguous",
                         [self._cand(i, 1.0) for i in hits[:3]])

        contains: set[int] = set()
        if usable_cn:
            contains |= {i for i, k in self._cn if len(k) >= CN_MIN and _near_contains(cq, k)}
        if usable_en:
            contains |= {i for i, k in self._en if len(k) >= EN_MIN and _near_contains(eq, k)}
        if len(contains) == 1:
            return self._found(contains.pop(), AUTO_SCORE, "contains")

        if usable_en:
            equal_py = sorted({i for i, p in self._py if p == eq})
            if len(equal_py) == 1:
                return self._found(equal_py[0], 0.95, "pinyin")
            if len(equal_py) > 1:
                return Match(None, "", "", 0.0, "ambiguous",
                             [self._cand(i, 0.95) for i in equal_py[:3]])
            prefix = {i for i, p in self._py
                      if min(len(p), len(eq)) >= PINYIN_PREFIX_MIN and (p.startswith(eq) or eq.startswith(p))}
            if len(prefix) == 1:
                return self._found(prefix.pop(), AUTO_SCORE, "pinyin")

        best: dict[int, tuple[float, bool]] = {}
        for is_cn, keys, query in ((True, self._cn, cq), (False, self._en, eq)):
            if not query:
                continue
            for i, k in keys:
                s = fuzz.ratio(query, k)
                if s > best.get(i, (-1.0, False))[0]:
                    best[i] = (s, is_cn)
        ranked = sorted(best.items(), key=lambda kv: (-kv[1][0], kv[0]))
        top = [self._cand(i, s / 100) for i, (s, _) in ranked[:3]]
        usable = usable_cn or usable_en
        if not ranked:
            return Match(None, "", "", 1.0, "new") if usable else Match(None, "", "", 0.0, "review")
        top_idx, (top_s, top_is_cn) = ranked[0]
        second = ranked[1][1][0] if len(ranked) > 1 else 0.0
        if top_s < REVIEW_MIN:
            if usable:
                return Match(None, "", "", 1.0, "new")
            return Match(None, "", "", round(top_s / 100, 4), "review", top)
        top_usable = usable_cn if top_is_cn else usable_en
        if top_usable and top_s >= FUZZY_MIN and top_s - second >= FUZZY_GAP:
            return self._found(top_idx, round(top_s * 0.95 / 100, 4), "fuzzy", top)
        return Match(None, "", "", round(top_s / 100, 4), "review", top)

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

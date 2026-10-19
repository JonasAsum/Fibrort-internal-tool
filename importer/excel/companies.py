"""Customer master index and company-name matching (see excel/SPEC.md, "Companies")."""
from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher
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
TAX_LABELS = {"纳税人登记号"}  # in the ERP lists: the 18-character unified social credit code (USCC)
LEGAL_LABELS = {"法人"}
USCC_RE = re.compile(r"[0-9A-Z]{18}")
HEADER_SCAN_ROWS = 5

AUTO_SCORE = 0.9
FUZZY_MIN = 90
FUZZY_GAP = 5
REVIEW_MIN = 60
SHOW_MIN = 70  # a candidate below this is shown only when it shares a real run of letters with the name
CN_MIN = 2  # shortest CJK core allowed to match by itself
EN_MIN = 4  # shortest latin core allowed to match by itself
PINYIN_PREFIX_MIN = 6
# A shorter core than this can't auto-match by containment or fuzzy ratio: one letter of
# difference is 10-20% of the name, and a longer name by one letter is often another company.
EN_SAFE = 8
CN_SAFE = 3
SHARED_RUN_EN = 4  # a review question needs at least this many letters in a row in common
SHARED_RUN_CN = 2

_BRACKET_PREFIX_RE = re.compile(r"^\s*[(\[（【][^)\]）】]*[)\]）】]")
_PUNCT_RE = re.compile(r"[^0-9a-z一-鿿]+")
_SUFFIX_RE = re.compile(
    r"有限责任公司|股份有限公司|有限公司|股份|集团|(?<![a-z])(?:co ltd|gmbh|inc|ltd)(?![a-z])"
)
CJK_RE = re.compile(r"[㐀-䶿一-鿿豈-﫿]+")
CODE_RE = re.compile(r"(?<![A-Za-z0-9])0*([A-Za-z]{2,4}\d{5,})(?!\d)")

CN_LEGAL = ("有限责任公司", "股份有限公司", "有限公司", "股份", "集团")
CN_GENERIC = ("责任", "科技", "技术", "精密", "机械", "设备", "装备", "智能", "自动化", "数控", "制造",
              "工业", "发展", "贸易", "实业", "机床", "国际",
              "机电", "电气", "电子", "工程", "系统", "材料", "能源", "仪器")
CN_STRIP = tuple(sorted(CN_LEGAL + CN_GENERIC, key=len, reverse=True))

LAT_STRIP = frozenset({
    "co", "ltd", "limited", "inc", "corp", "corporation", "gmbh", "llc",
    "company", "group", "technology", "tech", "precision", "machinery", "machine",
    "equipment", "intelligent", "automation", "cnc", "manufacturing", "industrial",
    "industry", "international",
    # Industry words: two companies that share only these are not alike ("Jiashengyuan Mechanical and
    # Electrical" vs "Oulang Mechanical and Electrical" is 81% by letters, 0% by name).
    "mechanical", "electrical", "electric", "electronic", "electronics", "electromechanical", "mechatronics",
    "and", "of", "the", "tool", "tools",  # 机床 is stripped from the Chinese side
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
CN_PLACES_SET = frozenset(CN_PLACES)
CN_WEAK = frozenset({"大学", "学院", "研究院", "研究所"})  # what is left of a school's name once place and industry are gone
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


def _fold_accents(s: str) -> str:
    """"Müller" -> "Muller": accents are dropped, every other letter (CJK, Hangul) is kept as it is."""
    d = unicodedata.normalize("NFKD", s)
    return unicodedata.normalize("NFC", "".join(ch for ch in d if unicodedata.category(ch) != "Mn"))


def _dkey(name) -> str:
    """The key an answer is stored under: the matching key, or for a name with no letters at all
    ("---") the name itself, so even that can be answered once and stays answered."""
    return norm(name) or " ".join(str(name or "").split()).lower()


def norm(text) -> str:
    """Matching key: NFKC, lowercase, no accents, no leading city bracket, no legal suffix, no punctuation or spaces."""
    if text is None:
        return ""
    s = _fold_accents(unicodedata.normalize("NFKC", str(text)).lower())
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
            elif lab in TAX_LABELS:
                cols.setdefault("uscc", idx)
            elif lab in LEGAL_LABELS:
                cols.setdefault("legal", idx)
        if "code" not in cols or not ({"cn", "en", "short"} & cols.keys()):
            continue
        out = []
        for row in rows[h + 1:]:
            rec = {"source": source}
            for field_name in ("code", "cn", "en", "short"):
                i = cols.get(field_name)
                rec[field_name] = _cell(row[i]) if i is not None and i < len(row) else ""
            for field_name in ("uscc", "legal"):  # only the ERP lists have them
                i = cols.get(field_name)
                if i is not None:
                    rec[field_name] = _cell(row[i]) if i < len(row) else ""
            if "uscc" in rec:  # an old 15-digit number is no USCC
                rec["uscc"] = rec["uscc"].upper() if USCC_RE.fullmatch(rec["uscc"].upper()) else ""
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
    uscc: str = ""  # unified social credit code: the one number that says "same legal entity"
    legal: str = ""  # legal representative

    def names(self) -> list[str]:
        return [n for n in [self.cn, self.en, self.short, *self.aliases] if n]

    def absorb(self, rec: dict) -> None:
        self.uscc = self.uscc or rec.get("uscc", "")
        self.legal = self.legal or rec.get("legal", "")
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
    s = _fold_accents(unicodedata.normalize("NFKC", str(text or ""))).replace("\n", " ")
    return _BRACKET_PREFIX_RE.sub(" ", s).strip()


def _canon_code(code: str) -> str:
    return code.upper().lstrip("0")


def _extract_codes(text: str) -> tuple[list[str], str]:
    codes = [_canon_code(m.group(1)) for m in CODE_RE.finditer(text)]
    return codes, CODE_RE.sub(" ", text)


def _real_code(raw: str) -> str:
    """The customer code in an ERP code cell, or "" when there is none ("New", empty, a name with a code glued on)."""
    raw = (raw or "").strip()
    if CJK_RE.search(raw):
        found = CODE_RE.search(raw)
        return found.group(0) if found else ""
    return "" if raw.lower() == "new" else raw


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
    for place in CN_PLACES_SORTED:  # "珩述自动化科技(上海)有限公司": a trailing city is no part of the name, as in English
        if s.endswith(place) and len(s) - len(place) >= CN_MIN:
            return s[:-len(place)]
    return s


def _en_core(latin: str) -> str:
    """Latin name letters/digits only, without legal suffixes, generic words and one leading place."""
    toks = [t for t in re.sub(r"[^0-9a-z]+", " ", latin.lower()).split() if t not in LAT_STRIP]
    for n in (2, 1):
        if len(toks) > n and " ".join(toks[:n]) in EN_PLACES and len("".join(toks[n:])) >= EN_MIN:
            toks = toks[n:]
            break
    for n in (2, 1):  # "Johnson Electric (Shenzhen)": a trailing place is not part of the name either
        if len(toks) > n and " ".join(toks[-n:]) in EN_PLACES and len("".join(toks[:-n])) >= EN_MIN:
            toks = toks[:-n]
            break
    return "".join(toks)


def _places(text) -> tuple[frozenset, frozenset]:
    """(latin places, CJK places) at the start or end of a name. Used to tell branches apart."""
    toks = [t for t in re.sub(r"[^0-9a-z]+", " ", CJK_RE.sub(" ", unicodedata.normalize("NFKC", str(text or "")).lower())).split()
            if t not in LAT_STRIP]
    lat: set[str] = set()
    for side in (lambda n: toks[:n], lambda n: toks[-n:]):
        for n in (2, 1):
            if len(toks) > n and " ".join(side(n)) in EN_PLACES:
                lat.add(" ".join(side(n)))
                break
    cjk = "".join(CJK_RE.findall(str(text or "")))
    for word in CN_LEGAL:
        cjk = cjk.replace(word, "")
    cn = {pl for pl in CN_PLACES_SORTED if len(cjk) > len(pl) and (cjk.startswith(pl) or cjk.endswith(pl))}
    return frozenset(lat), frozenset(cn)


CN_PLACE_EXONYMS = {"香港": "hongkong", "澳门": "macau", "台湾": "taiwan", "西藏": "tibet", "内蒙古": "innermongolia"}


def _place_keys(lat, cn) -> frozenset:
    """Places of a name as comparable keys whatever the script: "上海" and "Shanghai" are one key."""
    return frozenset({p.replace(" ", "") for p in lat} |
                     {CN_PLACE_EXONYMS.get(p) or "".join(lazy_pinyin(p)) for p in cn})


def _review(score: float, candidates) -> "Match":
    """A question for a person. Its score stays below the auto-accept line, however alike the candidates are."""
    return Match(None, "", "", min(round(score, 4), 0.89), "review", candidates)


def lookalike(a, b) -> bool:
    """Two names that may be one company written two ways: a shortened form, extra words, a small spelling change."""
    ta, tb = _prepare(a), _prepare(b)
    cja, cjb = "".join(CJK_RE.findall(ta)), "".join(CJK_RE.findall(tb))
    ca, cb = (_cn_core(cja) if cja else ""), (_cn_core(cjb) if cjb else "")
    if len(ca) >= CN_MIN and len(cb) >= CN_MIN and (ca == cb or _near_contains(ca, cb) or fuzz.ratio(ca, cb) >= 85):
        return True
    ea, eb = _en_core(CJK_RE.sub(" ", ta)), _en_core(CJK_RE.sub(" ", tb))
    return len(ea) >= EN_MIN and len(eb) >= EN_MIN and (ea == eb or _near_contains(ea, eb) or fuzz.ratio(ea, eb) >= 85)


def _real_run(score, is_cn, a, b, kind) -> bool:
    """Two keys alike by letters, not by chance. Chinese: 3 characters in a row, or 2 with a score of 75 (two of
    three is chance), or the shorter name whole inside the longer. Pinyin: a score of 80 or 6 letters in a row
    (any two names share "yuan"). English: 4 letters in a row."""
    if is_cn:
        return (_shares_run(a, b, 3) or (score >= 75 and _shares_run(a, b, SHARED_RUN_CN))
                or (min(len(a), len(b)) >= CN_MIN and min(a, b, key=len) not in CN_WEAK
                    and _shares_run(a, b, min(len(a), len(b)))))
    if kind == "py":
        return score >= 80 or _shares_run(a, b, 6)
    return _shares_run(a, b, SHARED_RUN_EN)


def _shares_run(a: str, b: str, need: int) -> bool:
    return SequenceMatcher(None, a, b, autojunk=False).find_longest_match(0, len(a), 0, len(b)).size >= need


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
        self.skipped_decisions: list[str] = []  # saved answers that could not be applied (and why)
        self.origin = "files"  # "crm" when the companies were read from the CRM (excel/crm_master.py)
        self._new_exact: set[str] = set()  # names a person answered "new company" to
        self._new_cn: list[str] = []
        self._new_en: list[str] = []
        self._places: dict[int, tuple[set, set]] = {}
        self._cores: dict[int, list[str]] = {}
        for idx, company in enumerate(companies):
            if company.code:
                self._by_code.setdefault(_canon_code(company.code), idx)
            places = (set(), set())
            for name in company.names():
                lat, cn_places = _places(name)
                places[0].update(lat)
                places[1].update(cn_places)
                self._places[idx] = places
                key = norm(name)
                if key:
                    self._exact.setdefault(key, []).append(idx)
                text = _prepare(name)
                cjk = "".join(CJK_RE.findall(text))
                cn = _cn_core(cjk) if cjk else ""
                if cn and cn not in CN_PLACES_SET:  # a name that is only a city is no key (沈阳机床 matches by exact name)
                    self._cores.setdefault(idx, []).append(cn)
                    self._cn.append((idx, cn))
                    self._py.append((idx, _pinyin(cn)))
                en = _en_core(CJK_RE.sub(" ", text))
                if en and en not in EN_PLACES:
                    self._cores.setdefault(idx, []).append(en)
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
            code = _real_code(rec.get("code", ""))  # "New" or a code with text glued on is no code: its own company
            if code and code in by_code:
                by_code[code].absorb(rec)
                continue
            company = Company(code=code, cn=rec.get("cn", ""), en=rec.get("en", ""), short=rec.get("short", ""),
                              uscc=rec.get("uscc", ""), legal=rec.get("legal", ""))
            companies.append(company)
            if code:
                by_code[code] = company
        return cls(companies, sources, [])

    @staticmethod
    def read_files(importer_dir) -> tuple[list[dict], list[str], dict]:
        """(records, skipped file names, records per source) of the customer list files on disk."""
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
        return records, skipped, sources

    def with_decisions(self, importer_dir) -> "Master":
        """Applies the saved "same company?" answers; a stale one is named in skipped_decisions, never fatal."""
        from .decisions import DecisionStore, default_path
        self.skipped_decisions = self.apply_decisions(DecisionStore(default_path(importer_dir)).as_dict(), strict=False)
        return self

    @classmethod
    def load(cls, importer_dir) -> "Master":
        records, skipped, sources = cls.read_files(importer_dir)
        master = cls.from_records(records)
        master.sources = sources
        master.skipped = skipped
        return master.with_decisions(importer_dir)

    def summary(self) -> dict:
        return {
            "records": dict(self.sources),
            "companies": self.size,
            "with_code": sum(1 for c in self.companies if c.code),
            "with_cn": sum(1 for c in self.companies if c.cn),
            "with_en": sum(1 for c in self.companies if c.en),
            "skipped": list(self.skipped),
        }

    def apply_decisions(self, decisions: dict, strict: bool = True) -> list[str]:
        """Answers to "same company?": {name: customer code} links the name to that customer,
        {name: None} makes it a known new company. strict: a bad entry raises ValueError (an answer
        a person just gave); otherwise it is skipped and named in the returned list, so one stale or
        hand-edited entry can never switch off all matching."""
        skipped: list[str] = []
        for name, code in decisions.items():
            problem = None
            if not isinstance(name, str):
                problem = f"decision with a name that is not text: {name!r}"
            elif code is not None and not isinstance(code, str):
                problem = f"decision for {name!r}: {code!r} is not a customer code or null"
            elif code is not None and self._by_code.get(_canon_code(code)) is None:
                problem = f"decision for {name!r}: unknown company code {code}"
            if problem:
                if strict:
                    raise ValueError(problem)
                skipped.append(problem)
                continue
            key = _dkey(name)
            if code is not None:
                idx = self._by_code[_canon_code(code)]
                if idx not in self._exact.setdefault(key, []):
                    self._exact[key].append(idx)
                if name not in self.companies[idx].aliases:
                    self.companies[idx].aliases.append(name)
                continue
            text = _prepare(name)
            self._new_exact.add(key)
            cjk = "".join(CJK_RE.findall(text))
            cn, en = (_cn_core(cjk) if cjk else ""), _en_core(CJK_RE.sub(" ", text))
            if len(cn) >= CN_MIN:
                self._new_cn.append(cn)
            if len(en) >= EN_MIN:
                self._new_en.append(en)
        return skipped

    def _decided_new(self, q: str, cq: str, eq: str) -> bool:
        if q in self._new_exact:
            return True
        return (len(cq) >= CN_MIN and any(_near_contains(cq, k) for k in self._new_cn)) or \
               (len(eq) >= EN_MIN and any(_near_contains(eq, k) for k in self._new_en))

    def match(self, name) -> Match:
        text = _prepare(name)
        codes, text = _extract_codes(text)
        cjk = "".join(CJK_RE.findall(text))
        cq = _cn_core(cjk) if cjk else ""
        eq = _en_core(CJK_RE.sub(" ", text))
        for code in codes:
            idx = self._by_code.get(code)
            if idx is not None:
                if self._contradicts(idx, cq, eq):  # a code-shaped word in a sentence about another company
                    return Match(None, "", "", 0.8, "review", [self._cand(idx, 0.8)])
                return self._found(idx, 1.0, "code")

        q = norm(text)
        places = _places(text)
        if not cq and not eq or self.size == 0:
            if _dkey(name) in self._new_exact:
                return Match(None, "", "", 1.0, "new")
            return Match(None, "", "", 0.0, "none")
        hits = sorted(set(self._exact.get(q, []))) if q else []
        if len(hits) == 1:
            return self._found(hits[0], 1.0, "exact")
        if len(hits) > 1:
            return Match(None, "", "", 0.0, "ambiguous",
                         [self._cand(i, 1.0) for i in hits[:3]])

        if self._decided_new(q, cq, eq):
            return Match(None, "", "", 1.0, "new")

        if cq in CN_PLACES_SET:  # "武汉机电" is a city and an industry: no name left to match on
            cq = ""
        if eq in EN_PLACES:
            eq = ""
        if not cq and not eq:  # nothing to compare, so no candidates to offer either: a new company, not a question
            return Match(None, "", "", 1.0, "new")
        usable_cn = len(cq) >= CN_MIN
        usable_en = len(eq) >= EN_MIN

        contains: set[int] = set()
        if usable_cn:
            contains |= {i for i, k in self._cn if (cq == k or min(len(cq), len(k)) >= CN_SAFE) and _near_contains(cq, k)}
        if usable_en:
            contains |= {i for i, k in self._en if (eq == k or min(len(eq), len(k)) >= EN_SAFE) and _near_contains(eq, k)}
        if len(contains) == 1:
            return self._found(contains.pop(), AUTO_SCORE, "contains", places=places)

        if usable_en:
            equal_py = sorted({i for i, p in self._py if p == eq})
            if len(equal_py) == 1:
                return self._found(equal_py[0], 0.95, "pinyin", places=places)
            if len(equal_py) > 1:
                return Match(None, "", "", 0.0, "ambiguous",
                             [self._cand(i, 0.95) for i in equal_py[:3]])
            prefix = {i for i, p in self._py
                      if min(len(p), len(eq)) >= PINYIN_PREFIX_MIN and (p.startswith(eq) or eq.startswith(p))}
            if len(prefix) == 1:
                return self._found(prefix.pop(), AUTO_SCORE, "pinyin", places=places)

        best: dict[int, tuple[float, bool, str, str]] = {}
        for kind, keys, query in (("cn", self._cn, cq), ("en", self._en, eq), ("py", self._py, eq)):
            if not query:
                continue
            for i, k in keys:
                s = fuzz.ratio(query, k)
                if s > best.get(i, (-1.0, False, "", "", ""))[0]:
                    best[i] = (s, kind == "cn", query, k, kind)
        # A middling score between two short names is chance, not resemblance: it only
        # counts when they share a real run of letters (else a question for every new company).
        run = {i: v[0] >= FUZZY_MIN or _real_run(*v) for i, v in best.items()}
        near = {i: v for i, v in best.items() if v[0] >= SHOW_MIN or run[i]}
        best = {i: v for i, v in best.items() if run[i]}
        # The three closest by score are shown with every question ("is it one of these?"), but never a chance
        # resemblance. Another city is a worse candidate, the same city a better one (a score only orders them
        # here; accepting is still decided on the plain score and _place_conflict).
        shown = [self._cand(i, s / 100) for i, (s, *_) in
                 sorted(near.items(), key=lambda kv: (-self._place_adjusted(kv[0], kv[1][0], places), kv[0]))[:3]]
        ranked = sorted(best.items(), key=lambda kv: (-kv[1][0], kv[0]))
        top = [self._cand(i, s / 100) for i, (s, *_) in ranked[:3]]
        usable = usable_cn or usable_en
        if not ranked:
            return Match(None, "", "", 1.0, "new") if usable else Match(None, "", "", 0.0, "review")
        top_idx, (top_s, top_is_cn, *_rest) = ranked[0]
        second = ranked[1][1][0] if len(ranked) > 1 else 0.0
        if top_s < REVIEW_MIN:
            if usable:
                return Match(None, "", "", 1.0, "new")
            return _review(top_s / 100, shown)
        top_usable = usable_cn if top_is_cn else usable_en
        long_enough = top_is_cn or len(eq) >= EN_SAFE
        if top_usable and long_enough and top_s >= FUZZY_MIN and top_s - second >= FUZZY_GAP:
            return self._found(top_idx, round(top_s * 0.95 / 100, 4), "fuzzy", top, places=places)
        return _review(top_s / 100, shown)

    def get(self, code) -> dict | None:
        idx = self._by_code.get(_canon_code(str(code or "")))
        if idx is None:
            return None
        c = self.companies[idx]
        return {"code": c.code, "cn": c.cn, "en": c.en, "short": c.short, "uscc": c.uscc, "legal": c.legal}

    def has_code(self, code) -> bool:
        return _canon_code(str(code or "")) in self._by_code

    def search(self, query: str, limit: int = 8) -> list[dict]:
        """Companies whose code or any name looks like what a person typed (for the "existing company" list)."""
        q = norm(query)
        if not q:
            return []
        scored = []
        for idx, c in enumerate(self.companies):
            best = 0.0
            if c.code and q == _canon_code(c.code).lower():
                best = 100.0
            for name in c.names():
                key = norm(name)
                if not key:
                    continue
                if q in key:
                    best = max(best, 90.0 + 10.0 * len(q) / len(key))
                elif len(q) >= 5:  # a typo-tolerant hit only for a query long enough to mean something
                    pr = fuzz.partial_ratio(q, key)
                    if pr >= 85:
                        best = max(best, pr * 0.85)
            if best >= 70:
                scored.append((best, idx))
        scored.sort(key=lambda t: (-t[0], t[1]))
        return [{"code": self.companies[i].code, "cn": self.companies[i].cn, "en": self.companies[i].en,
                 "short": self.companies[i].short, "uscc": self.companies[i].uscc} for _, i in scored[:limit]]

    def _contradicts(self, idx: int, cq: str, eq: str) -> bool:
        """The rest of the cell (besides a customer code) names something else than that customer."""
        usable = [(q, CN_MIN) for q in (cq,) if q] + [(q, EN_MIN) for q in (eq,) if q]
        usable = [q for q, need in usable if len(q) >= need]
        if not usable:
            return False
        cores = self._cores.get(idx, [])
        return not any(fuzz.ratio(q, k) >= 50 or _near_contains(q, k) for q in usable for k in cores)

    def place_check(self, match: "Match", text: str) -> "Match":
        """The match, or a question when the whole cell names another place than the company it found (a hit found
        through one language part of "JUGANG 宁夏巨冈" must still agree with the Ningxia in the other part)."""
        idx = self._by_code.get(_canon_code(str(match.code or ""))) if match.code else None
        if idx is None or match.method == "code" or not self._place_conflict(idx, _places(text)):
            return match
        return Match(None, "", "", min(match.score, 0.89), "review", [self._cand(idx, min(match.score, 0.89))])

    def _place_adjusted(self, idx: int, score: float, places) -> float:
        if self._place_conflict(idx, places):
            return score - 15
        c_lat, c_cn = self._places.get(idx, (set(), set()))
        return score + 3 if _place_keys(*places) & _place_keys(c_lat, c_cn) else score

    def _place_conflict(self, idx: int, places) -> bool:
        lat, cn = places
        c_lat, c_cn = self._places.get(idx, (set(), set()))
        if bool(lat and c_lat and not lat & c_lat) or bool(cn and c_cn and not cn & c_cn):
            return True
        mine, theirs = _place_keys(lat, cn), _place_keys(c_lat, c_cn)  # one name in each script
        return bool(mine and theirs and not mine & theirs)

    def _found(self, idx: int, score: float, method: str, candidates=None, places=None) -> Match:
        c = self.companies[idx]
        if places is not None and self._place_conflict(idx, places):
            # Same name core, other city: a branch or a sister company. Never merge silently.
            return Match(None, "", "", min(score, 0.89), "review", [self._cand(idx, min(score, 0.89))])
        if score >= AUTO_SCORE:
            candidates = []
        elif candidates is None:
            candidates = [self._cand(idx, score)]
        return Match(c.code or None, c.cn, c.en, score, method, candidates)

    def _cand(self, idx: int, score: float) -> tuple:
        c = self.companies[idx]
        return (c.code, c.cn or c.en or c.short, round(score, 4))

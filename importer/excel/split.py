"""Splits a customer cell into the company name and its address, and turns phone
cells into text. A cell with a newline gives name then address; a legal suffix
(Ltd., GmbH, 有限公司) ends the name; an address marker is a guess and is flagged.
Placeholders ("/", "none", "无" ...) are empty everywhere."""
import re
import unicodedata
from datetime import date, datetime

SUFFIX = re.compile(r"股份有限公司|有限公司|Co\.?,?\s*Ltd\.?|Ltd\.?|GmbH", re.I)
ADDRESS_MARK = re.compile(r"省|市|区|路|号|No\.|Road|District")
EN_ADDRESS = re.compile(r"\b(?:Village|Town|City|Zone|Park|Street|Province|County|Avenue|Lane|Building|Floor|Block)\b", re.I)


def looks_like_address(text) -> bool:
    """A line that reads like a postal address: a Chinese or English address word, or a house number."""
    t = "" if text is None else str(text)
    return bool(ADDRESS_MARK.search(t) or EN_ADDRESS.search(t) or re.search(r"\d", t))
LATIN_ALIAS = re.compile(
    r"^(?P<cn>.*[\u4e00-\u9fff].*?(?:股份有限公司|有限公司))\s*[(（]\s*"
    r"(?P<en>[A-Za-z][A-Za-z0-9 .,&'\-]*?)\s*[)）]?\s*$")
SUFFIX_END = re.compile(r"(?:股份有限公司|有限责任公司|有限公司|Co\.?,?\s*Ltd\.?|Ltd\.?|GmbH)\s*\.?$", re.I)
TRIM = " ,;，、()（）"
SHORT_NAME = 40
LONG_NAME = 80
NAME_BEFORE_MARK = 6
PLACEHOLDERS = {"/", "-", "—", "not offer", "none", "n/a", "无", "暂无", "没有"}
SEPARATORS = re.compile(r"[/,;，、;\s]+")
NUMBER_DIGITS = 6  # a token with fewer digits is a prefix (+86, 0755) of the next one
LONG_RUN = 13


def is_placeholder(value) -> bool:
    """True for "/", "-", "none", "无" and the other empty stand-ins, any case."""
    if value is None:
        return False
    text = re.sub(r"\s+", " ", str(value)).strip().casefold()
    return text in PLACEHOLDERS


def cell_text(value, *, keep_lines: bool = False) -> str:
    """The cell as tidy text. Placeholders ("/", "none", "无" ...) come back empty; dates come back
    ISO; a whole number loses its ".0". keep_lines keeps line breaks (to-do text)."""
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    if keep_lines:
        lines = [" ".join(line.split()) for line in unicodedata.normalize("NFC", str(value)).replace("\r", "").split("\n")]
        text = "\n".join(line for line in lines if line)
    else:
        text = " ".join(str(value).split())
    return "" if is_placeholder(text) else text


def split_customer(cell) -> tuple[str, str, bool]:
    """(name, address, sure)."""
    text = "" if cell is None else str(cell).strip()
    if not text or is_placeholder(text):
        return "", "", True
    if "\n" in text:
        name, _, rest = text.partition("\n")
        rest = " ".join(rest.split()).strip(TRIM)
        # A second line that is itself a company name ("wendaxin" over "文达新", or ending
        # in 有限公司) belongs to the name; only a line that reads like an address is split off.
        if rest and (SUFFIX.search(rest) and SUFFIX.search(rest).end() >= len(rest) - 1
                     or not ADDRESS_MARK.search(rest) and not re.search(r"\d", rest)):
            return f"{name.strip(TRIM)} {rest}".strip(), "", True
        return name.strip(TRIM), rest, True
    if m := SUFFIX.search(text):
        tail = text[m.end():].strip(TRIM)
        if tail and SUFFIX.fullmatch(tail.strip(" .,")):  # "… Co., Ltd. Co.,Ltd.": the suffix typed twice
            tail = ""
        return text[:m.end()].strip(TRIM), tail, True
    if m := ADDRESS_MARK.search(text, NAME_BEFORE_MARK):
        return text[:m.start()].strip(TRIM), text[m.start():].strip(TRIM), False
    # A longer name with several words and no digit or address marker is still a name.
    return text, "", len(text) <= SHORT_NAME or (len(text) <= LONG_NAME and " " in text and not re.search(r"\d", text))


LATIN_ALIAS_ANY = re.compile(
    r"^(?P<cn>.*?[\u4e00-\u9fff].*?(?:股份有限公司|有限责任公司|有限公司))\s*[(（]\s*(?P<en>[A-Za-z].*?)\s*[)）]?\s*$")
ADDR_LABEL = re.compile(r"^\s*(?:Addr(?:ess)?|地址)\s*[:：]\s*", re.I)
CJK = re.compile(r"[\u4e00-\u9fff]")


CUSTOMER_CODE = re.compile(r"(?<![A-Za-z0-9])0*CR\d{5,}(?!\d)")


def split_customer_code(raw) -> tuple[str, str]:
    """(ERP customer code, the cell without it). "CR0701256⏎四川…" and "四川…⏎CR0701256" both give
    ("CR0701256", "四川…"); a cell with no code comes back unchanged with code "". The code is the
    customer's ERP number, not part of its name, so it goes to its own field."""
    text = "" if raw is None else str(raw)
    m = CUSTOMER_CODE.search(text)
    if not m:
        return "", text
    rest = (text[:m.start()] + "\n" + text[m.end():]).strip(" \t,;，、:：")
    lines = [ln.strip(" \t,;，、:：") for ln in rest.split("\n")]
    return m.group(0), "\n".join(ln for ln in lines if ln)


def split_two_lines(raw) -> tuple[str, str, str] | None:
    """(name, english name, address) for a two or three line customer cell of a shape that the flattened
    text mangles, else None. `raw` keeps its line breaks (cell_text(..., keep_lines=True)).

        中文名(English)   |  中文名      |  中文名           |  中文名
        Addr:重庆市…      |  Addr:重庆市… |  English Co., Ltd. |  (English Co., Ltd.)
                          |              |                   |  Addr:重庆市…

    Any other cell, including one with a customer code (CR…) line, returns None and goes through the
    usual path unchanged."""
    lines = [ln.strip() for ln in str(raw or "").split("\n") if ln.strip()]
    if len(lines) == 1 and CJK.search(lines[0]):  # "中文有限公司 ENGLISH NAME": a Latin remainder after a Chinese legal-suffix name
        m = re.search(r"(?:股份有限公司|有限责任公司|有限公司)", lines[0])
        if m:
            tail = lines[0][m.end():].strip(TRIM + ".")
            if tail and not CJK.search(tail) and re.search(r"[A-Za-z]{3}", tail) and not looks_like_address(tail):
                return lines[0][:m.end()].strip(TRIM), tail, ""
    if len(lines) == 1:  # "English Co., Ltd.中文名": a Chinese remainder after a Latin legal-suffix name
        m = SUFFIX.search(lines[0])
        if m and not CJK.search(lines[0][:m.end()]):
            tail = lines[0][m.end():].strip(TRIM + ".")
            if CJK.search(tail) and not looks_like_address(tail):
                return lines[0][:m.end()].strip(TRIM), tail, ""
    if len(lines) == 2 and not CJK.search(lines[0]) and SUFFIX_END.search(lines[0]) and CJK.search(lines[1]) \
            and not ADDR_LABEL.match(lines[1]) and is_other_name(lines[1]):
        return lines[0].strip(TRIM), lines[1].strip(TRIM), ""  # "English Co., Ltd.⏎中文名": the other-language name
    if not 1 <= len(lines) <= 3 or not CJK.search(lines[0]):
        return None
    m = LATIN_ALIAS_ANY.match(lines[0])  # also an English name with brackets of its own: "(… (Chongqing) Co., Ltd.)"
    name, english = (m.group("cn"), m.group("en")) if m else (lines[0], "")
    address = ""
    for ln in lines[1:]:
        if ADDR_LABEL.match(ln) and not address:
            address = ADDR_LABEL.sub("", ln).strip()
        elif not english and not CJK.search(ln) and (
                re.fullmatch(r"[(（]\s*[A-Za-z].*[)）]", ln) or SUFFIX_END.search(ln)
                or (SUFFIX_END.search(name) and is_other_name(ln))):
            english = ln.strip(TRIM)
        else:
            return None
    return (name.strip(TRIM), english, address) if english or address else None

def is_other_name(text: str) -> bool:
    """A second line that is a company name in its own right (ends in a legal suffix, or
    carries no address marker or house number)."""
    text = " ".join(str(text).split())
    return bool(SUFFIX_END.search(text)) or not looks_like_address(text)


def latin_alias(cell) -> tuple[str, str]:
    """("上海某某有限公司(Some Name)") -> ("上海某某有限公司", "Some Name"); other cells unchanged."""
    text = "" if cell is None else str(cell).strip()
    if m := LATIN_ALIAS.match(text):
        return m.group("cn").strip(TRIM), m.group("en").strip(TRIM)
    return text, ""


def _groups(text: str) -> list[str]:
    """Numbers in a phone cell: tokens split on / , ; space, with short prefixes joined to the next token."""
    groups, pending = [], []
    for token in (t for t in SEPARATORS.split(text) if t):
        pending.append(token)
        if sum(ch.isdigit() for ch in token) >= NUMBER_DIGITS:
            groups.append(" ".join(pending))
            pending = []
    if pending:  # trailing words without digits ("(Mr. Li)") stay with the number before them
        if groups:
            groups[-1] = f"{groups[-1]} {' '.join(pending)}"
        else:
            groups.append(" ".join(pending))
    out = []
    for g in groups:
        if g.isdigit() and len(g) > LONG_RUN:  # two numbers glued together
            cut = len(g) // 2 if len(g) % 2 == 0 else 11
            out += [g[:cut], g[cut:]]
        else:
            out.append(g)
    return out


def phone_numbers(value) -> list[str]:
    """Every number in a phone cell, in order; [] when the cell is empty or a placeholder."""
    if value is None or is_placeholder(value):
        return []
    if isinstance(value, float) and value.is_integer():
        return [str(int(value))]
    text = str(value).strip()
    if not text:
        return []
    return _groups(text) or [text]


def phone_text(value) -> str:
    """A phone cell as text; numbers stored as floats lose their ".0"."""
    if value is None or is_placeholder(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


MOBILE_CN = re.compile(r"1[3-9]\d{9}")


def phone_e164(raw) -> str:
    """Phone as +countrycode digits. No country code: an 11-digit 13-19 mobile is Chinese, and a
    number starting 0 is a Chinese landline (its 0 dropped). Anything with a + (+49, +86) is kept as written.
    Returns "" for a placeholder and the cleaned text when the rule does not apply."""
    text = phone_text(raw)
    if not text:
        return ""
    s = re.sub(r"[\s\-().（）/]", "", text)
    s = re.sub(r"^(?:tel|t|f|m|mob|fax)[:：]?(?=[+\d])", "", s, flags=re.I)  # a label such as T / Tel / M
    if s.startswith("00"):
        s = "+" + s[2:]
    if s.startswith("+"):
        return "+" + re.sub(r"\D", "", s)
    if re.fullmatch(r"86\d{11}", s):  # 86 without the +
        return "+" + s
    if MOBILE_CN.fullmatch(s):
        return "+86" + s
    if re.fullmatch(r"0\d{9,11}", s):  # Chinese landline with its trunk 0
        return "+86" + s[1:]
    return text


def phone_type(number: str) -> str:
    """The CRM label of an E.164 number, decided by the number alone: "mobile" for a Chinese mobile (+86 1[3-9]...),
    "work" for everything else (landlines such as Beijing +8610..., 400/800 lines, foreign numbers)."""
    return "mobile" if re.fullmatch(r"\+86" + MOBILE_CN.pattern, number or "") else "work"


def looks_like_phone(text) -> bool:
    """A cell that is a phone number (digits and separators, at most a short label), not an address or a name."""
    t = "" if text is None else str(text)
    return sum(ch.isdigit() for ch in t) >= 6 and sum(ch.isalpha() for ch in t) <= 4


HONORIFICS = frozenset("mr mrs ms miss dr prof sir".split())
DATE_LIKE = re.compile(r"^\s*(?:\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2}(?:[-/.]\d{2,4})?)\s*$")


def looks_like_person_name(text) -> bool:
    """False for a contact cell that cannot be a person: empty, a date, digits only, or just a title ("Mr.")."""
    t = "" if text is None else str(text).strip()
    if not t or DATE_LIKE.match(t) or not any(ch.isalpha() for ch in t):
        return False
    words = re.findall(r"[^\W\d_]+", t.lower())
    return not all(w in HONORIFICS for w in words)


def visit_kind(cell) -> str:
    """"internal" or "exhibition" for a row that is not a customer visit, "" otherwise."""
    text = "" if cell is None else str(cell)
    if EXHIBITION.search(text):  # before the to-do verbs: "Attend the Dalian exhibition" is an exhibition
        return "exhibition"
    if INTERNAL.search(text):
        return "internal"
    if _is_sentence(text):
        return "internal"
    return ""


# Grammar words, not business words: a company name has at most one of them ("Institute of Technology"), a sentence
# written in the customer cell ("On the way to the ...", "Get ready for the ...") has several.
FUNCTION_WORDS = frozenset("the a an to for in on at with from by is are be will".split())
SENTENCE_MIN_FUNCTION_WORDS = 2


def _is_sentence(text: str) -> bool:
    """A customer cell that reads as a sentence, not a name: no legal suffix and several grammar words."""
    if SUFFIX.search(text) or CJK.search(text):  # a Chinese name inside is a customer with a prefix: left for a person
        return False
    words = re.findall(r"[a-z]+", text.lower())
    return sum(w in FUNCTION_WORDS for w in words) >= SENTENCE_MIN_FUNCTION_WORDS


INTERNAL = re.compile(r"holiday|假期|休息|meeting online|internal meeting|\boffice\b|training|培训|team building"
                      # a to-do written in the customer cell (marketing rows: "Register a Tik Tok account.")
                      r"|^\s*(?:register|upload|get ready|attend|prepare|collect|arrange)\b", re.I)
EXHIBITION = re.compile(r"exhibition|展会|\bexpo\b|\bfair\b|\bmesse\b", re.I)


def fold_key(text) -> str:
    """Matching key for names: NFKC, casefolded, no spaces or punctuation."""
    return re.sub(r"[\W_]+", "", unicodedata.normalize("NFKC", str(text or "")).casefold())

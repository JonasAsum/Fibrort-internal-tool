"""Splits a customer cell into the company name and its address, and turns phone
cells into text. A cell with a newline gives name then address; a legal suffix
(Ltd., GmbH, 有限公司) ends the name; an address marker is a guess and is flagged.
Placeholders ("/", "none", "无" ...) are empty everywhere."""
import re

SUFFIX = re.compile(r"股份有限公司|有限公司|Co\.?,?\s*Ltd\.?|Ltd\.?|GmbH", re.I)
ADDRESS_MARK = re.compile(r"省|市|区|路|号|No\.|Road|District")
LATIN_ALIAS = re.compile(
    r"^(?P<cn>.*[\u4e00-\u9fff].*?(?:股份有限公司|有限公司))\s*[(（]\s*"
    r"(?P<en>[A-Za-z][A-Za-z0-9 .,&'\-]*?)\s*[)）]?\s*$")
SUFFIX_END = re.compile(r"(?:股份有限公司|有限责任公司|有限公司|Co\.?,?\s*Ltd\.?|Ltd\.?|GmbH)\s*\.?$", re.I)
TRIM = " ,;，、()（）"
SHORT_NAME = 40
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
        return text[:m.end()].strip(TRIM), text[m.end():].strip(TRIM), True
    if m := ADDRESS_MARK.search(text, NAME_BEFORE_MARK):
        return text[:m.start()].strip(TRIM), text[m.start():].strip(TRIM), False
    return text, "", len(text) <= SHORT_NAME


def is_other_name(text: str) -> bool:
    """A second line that is a company name in its own right (ends in a legal suffix, or
    carries no address marker or house number)."""
    text = " ".join(str(text).split())
    return bool(SUFFIX_END.search(text)) or not (ADDRESS_MARK.search(text) or re.search(r"\d", text))


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


def visit_kind(cell) -> str:
    """"internal" or "exhibition" for a row that is not a customer visit, "" otherwise."""
    text = "" if cell is None else str(cell)
    if INTERNAL.search(text):
        return "internal"
    if EXHIBITION.search(text):
        return "exhibition"
    return ""


INTERNAL = re.compile(r"holiday|假期|休息|meeting online|internal meeting|\boffice\b|training|培训|team building", re.I)
EXHIBITION = re.compile(r"exhibition|展会|\bexpo\b", re.I)

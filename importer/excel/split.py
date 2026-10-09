"""Splits a customer cell into the company name and its address, and turns phone
cells into text. A cell with a newline gives name then address; a legal suffix
(Ltd., GmbH, 有限公司) ends the name; an address marker is a guess and is flagged."""
import re

SUFFIX = re.compile(r"股份有限公司|有限公司|Co\.?,?\s*Ltd\.?|Ltd\.?|GmbH", re.I)
ADDRESS_MARK = re.compile(r"省|市|区|路|号|No\.|Road|District")
TRIM = " ,;，、()（）"
SHORT_NAME = 40


def split_customer(cell) -> tuple[str, str, bool]:
    """(name, address, sure)."""
    text = "" if cell is None else str(cell).strip()
    if not text:
        return "", "", True
    if "\n" in text:
        name, _, address = text.partition("\n")
        return name.strip(TRIM), address.strip(TRIM), True
    if m := SUFFIX.search(text):
        return text[:m.end()].strip(TRIM), text[m.end():].strip(TRIM), True
    if m := ADDRESS_MARK.search(text, 2):
        return text[:m.start()].strip(TRIM), text[m.start():].strip(TRIM), False
    return text, "", len(text) <= SHORT_NAME


def phone_text(value) -> str:
    """A phone cell as text; numbers stored as floats lose their ".0"."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()

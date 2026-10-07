"""Consolidates the spelling variants a visit list uses for one real company
or contact into a single canonical record, the same "matching step" the
mock fixture's ground truth exists to test (mock_visits/README.md).

The CRM's own importer already resolves a contact's employer by exact
(case/whitespace-folded) name match against the estate (csvemployer.go), and
dedupes a company against the estate the same way the UI's merge screen
would. Neither of those collapses five spellings of ONE NEW company inside a
single file into one row - that has to happen here, before the file is
built.
"""
import re
from collections import Counter
from dataclasses import dataclass, field

from .reader import Visit

_PUNCT = re.compile(r"[.,/\\]")
_SPACE = re.compile(r"\s+")
_SALUTATION = re.compile(r"^(mr|ms|mrs|mister|miss)\.?\s*", re.I)
_LEGAL_SUFFIXES = {"co", "ltd", "inc", "corp", "corporation", "limited", "gmbh", "llc", "plc", "group", "company"}

# Common transliterated Chinese surnames, used only to split a run-together
# rendering like "Zhangwei" into its surname ("zhang") for grouping. Anyone
# not on this list falls back to the whole token, which only costs a split
# between two renderings of the same contact that a human corrects via
# "Merge contact" - the CRM's own fallback for a fuzzy match (see
# docs/handbook/records.md).
SURNAMES = ["zhang", "wang", "li", "liu", "chen", "yang", "huang", "zhao", "zhou", "wu", "xu",
            "sun", "ma", "zhu", "hu", "guo", "he", "lin", "gao", "luo", "qian", "tang", "song"]
# Longest first: "lin" and "huang" must win over the "li" and "hu" they start with.
SURNAMES.sort(key=len, reverse=True)


def _normalize(name: str) -> list[str]:
    tokens = _SPACE.sub(" ", _PUNCT.sub(" ", name.lower())).strip().split()
    while tokens and tokens[-1] in _LEGAL_SUFFIXES:
        tokens.pop()
    return tokens


def _name_tokens(name: str) -> tuple[tuple[str, ...], tuple[bool, ...]]:
    """The company-identifying words of `name`, and which of them were
    written abbreviated ("Auto." in `Hengrui Auto.`), so a match may treat
    those as a prefix and every other word as exact."""
    words = _SPACE.sub(" ", name.replace(",", " ")).strip().split()
    while words and _normalize(words[-1]) == []:  # trailing "Co.", "Ltd"
        words.pop()
    tokens, abbreviated = [], []
    for w in words:
        folded = "".join(_normalize(w))
        if folded:
            tokens.append(folded)
            abbreviated.append(w.endswith("."))
    return tuple(tokens), tuple(abbreviated)


def _could_be(tokens: tuple[str, ...], abbreviated: tuple[bool, ...], full: tuple[str, ...]) -> bool:
    if len(tokens) > len(full):
        return False
    return all(t == f or (abbr and f.startswith(t)) for t, abbr, f in zip(tokens, abbreviated, full))


class _CompanyNamer:
    """Resolves every way a visit row writes a customer to one company key.

    A name written out in full (two or more words, none abbreviated) is its
    own key, compared word for word: `Hengrui Intelligent Equipment` and
    `Hengrui Intelligent Manufacturing` are two companies. A shortened form
    (`Hengrui Auto.`, `HENGRUI`, `Hengrui Co.`) is resolved to a full name the
    file spells out elsewhere, by evidence on its own row, in this order: the
    contact met is one already met at exactly one candidate; the row's city is
    the home city of exactly one candidate; there is only one candidate and
    the row names no city. Anything else becomes a company of its own, keyed
    by the short form and city: a duplicate a human merges in one click
    ("Merge company", docs/handbook/records.md) costs far less than two real
    customers' visits silently filed under one name.
    """

    def __init__(self, visits: list[Visit]):
        self._full_by_first: dict[str, set[tuple[str, ...]]] = {}
        self._cities: dict[tuple[str, ...], Counter[str]] = {}
        self._surnames: dict[tuple[str, ...], set[str]] = {}
        self.decisions: list[dict] = []
        for v in _customer_rows(visits):
            tokens, abbreviated = _name_tokens(v.customer_raw)
            if len(tokens) < 2 or any(abbreviated):
                continue
            self._full_by_first.setdefault(tokens[0], set()).add(tokens)
            if v.city:
                self._cities.setdefault(tokens, Counter())[v.city] += 1
            if v.contact_raw:
                self._surnames.setdefault(tokens, set()).add(_surname_key(v.contact_raw))

    def key_for(self, customer_raw: str, city: str = "", contact_raw: str = "", record: bool = True) -> str:
        tokens, abbreviated = _name_tokens(customer_raw)
        if not tokens:
            return customer_raw.strip().lower()
        if len(tokens) >= 2 and not any(abbreviated):
            return ":".join(tokens)
        candidates = sorted(f for f in self._full_by_first.get(tokens[0], ())
                            if _could_be(tokens, abbreviated, f))
        pick, why = self._pick(candidates, city, contact_raw)
        key = ":".join(pick) if pick else f"{tokens[0]}@{city.strip().lower()}"
        if record:
            self.decisions.append({"written": customer_raw, "city": city, "contact": contact_raw,
                                   "candidates": [":".join(c) for c in candidates], "key": key, "why": why})
        return key

    def _pick(self, candidates: list[tuple[str, ...]], city: str, contact_raw: str):
        if not candidates:
            return None, "no full name in the file starts this way"
        # A surname alone is weak evidence - "Mr. Li" works at half the
        # estate - so it only breaks a tie among candidates the city allows.
        in_city = [c for c in candidates if not city or (self._cities.get(c) and self._cities[c].most_common(1)[0][0] == city)]
        if contact_raw:
            surname = _surname_key(contact_raw)
            by_contact = [c for c in in_city if surname in self._surnames.get(c, ())]
            if len(by_contact) == 1:
                return by_contact[0], "same contact met there"
        if city:
            if len(in_city) == 1:
                return in_city[0], "same city"
            return None, "kept separate: no candidate in this city" if not in_city else "kept separate: ambiguous"
        if len(candidates) == 1:
            return candidates[0], "only candidate"
        return None, "kept separate: ambiguous"


def _customer_rows(visits: list[Visit]):
    return (v for v in visits if v.kind in ("customer", "dealer") and v.customer_raw)


def _surname_key(contact_raw: str) -> str:
    stripped = _SALUTATION.sub("", contact_raw.strip())
    token = "".join(_normalize(stripped))
    for surname in SURNAMES:
        if token.startswith(surname):
            return surname
    return token


def _is_spelled_out(raw_name: str) -> bool:
    """True when `raw_name` carries more than a bare surname: a given name as
    its own word (`Zhang Wei`), or run together after a recognised surname
    (`Zhangwei`). Used to prefer the fuller rendering as the canonical
    display name over a salutation-only one (`Mr.Zhang`) that a small
    sample can otherwise make look more common by chance."""
    stripped = _SALUTATION.sub("", raw_name.strip())
    tokens = _normalize(stripped)
    if len(tokens) >= 2:
        return True
    if not tokens:
        return False
    token = tokens[0]
    return any(token.startswith(s) and len(token) > len(s) for s in SURNAMES)


def _canonical(counter: Counter) -> str:
    spelled_out = Counter({name: n for name, n in counter.items() if _is_spelled_out(name)})
    return (spelled_out or counter).most_common(1)[0][0]


@dataclass
class Company:
    key: str
    display_name: str
    city: str
    is_dealer: bool = False
    # The newest value the visits state, by visit date: a rating that drifts
    # from C to A over the year lands as A, and every older rating stays on
    # the visit that recorded it.
    potential: str = ""
    customer_type: str = ""
    owners: list[str] = field(default_factory=list)
    first_visit: str = ""
    last_visit: str = ""
    visit_count: int = 0


@dataclass
class Contact:
    key: str
    company_key: str
    full_name: str
    phones: list[str] = field(default_factory=list)  # mobile column, every distinct value in file order
    office_phones: list[str] = field(default_factory=list)
    emails: list[str] = field(default_factory=list)
    spellings: list[str] = field(default_factory=list)
    owners: list[str] = field(default_factory=list)


@dataclass
class MatchResult:
    companies: dict[str, Company] = field(default_factory=dict)
    contacts: dict[str, Contact] = field(default_factory=dict)
    # visit_key -> (company_key | None, contact_key | None)
    links: dict[str, tuple[str | None, str | None]] = field(default_factory=dict)

    def resolve_company(self, customer_raw: str, city: str = "") -> str | None:
        """A name written outside the visit rows (the week's key plan) folded
        the same way the rows were. None when no visit names that company."""
        key = self._namer.key_for(customer_raw, city, record=False) if self._namer else None
        return key if key in self.companies else None

    _namer: "_CompanyNamer | None" = None


def _add_distinct(values: list[str], value: str) -> None:
    value = value.strip()
    if value and value not in values:
        values.append(value)


def match(visits: list[Visit]) -> MatchResult:
    result = MatchResult()
    namer = _CompanyNamer(visits)
    result._namer = namer
    raw_names: dict[str, Counter] = {}
    cities: dict[str, Counter] = {}
    dealer_keys: set[str] = set()
    contact_raw_names: dict[str, Counter] = {}
    contacts: dict[str, Contact] = {}
    latest: dict[str, dict[str, tuple[str, str]]] = {}  # company -> field -> (date, value)
    owners: dict[str, list[str]] = {}
    dates: dict[str, list[str]] = {}

    for v in visits:
        if v.kind not in ("customer", "dealer") or not v.customer_raw:
            result.links[v.visit_key] = (None, None)
            continue
        company_key = namer.key_for(v.customer_raw, v.city, v.contact_raw)
        raw_names.setdefault(company_key, Counter())[v.customer_raw] += 1
        if v.city:
            cities.setdefault(company_key, Counter())[v.city] += 1
        if v.kind == "dealer":
            dealer_keys.add(company_key)
        _add_distinct(owners.setdefault(company_key, []), v.owner)
        when = v.date_iso or ""
        if when:
            dates.setdefault(company_key, []).append(when)
        potential = v.potential.strip().upper()
        # The CRM picklist only accepts A/B/C; "A+", "B-", "TBD" would 422 the company PATCH.
        potential = potential if potential in ("A", "B", "C") else ""
        for name, value in (("potential", potential), ("customer_type", v.customer_type)):
            seen = latest.setdefault(company_key, {}).get(name)
            if value and (seen is None or when >= seen[0]):
                latest[company_key][name] = (when, value)

        contact_key = None
        if v.contact_raw:
            contact_key = f"{company_key}::{_surname_key(v.contact_raw)}"
            contact_raw_names.setdefault(contact_key, Counter())[v.contact_raw] += 1
            c = contacts.setdefault(contact_key, Contact(key=contact_key, company_key=company_key, full_name=""))
            _add_distinct(c.phones, v.phone_raw)
            _add_distinct(c.office_phones, v.office_phone)
            _add_distinct(c.emails, v.email)
            _add_distinct(c.spellings, v.contact_raw)
            _add_distinct(c.owners, v.owner)
        result.links[v.visit_key] = (company_key, contact_key)

    for key, counter in raw_names.items():
        # A bare mention ("Senlan") can outnumber the spelled-out name by
        # chance on a company with only a handful of visits; prefer the most
        # common SPELLED-OUT rendering when the group has one at all.
        seen_dates = sorted(dates.get(key, []))
        result.companies[key] = Company(
            key=key, display_name=_canonical(counter),
            city=cities[key].most_common(1)[0][0] if key in cities else "",
            is_dealer=key in dealer_keys,
            potential=latest.get(key, {}).get("potential", ("", ""))[1],
            customer_type=latest.get(key, {}).get("customer_type", ("", ""))[1],
            owners=owners.get(key, []),
            first_visit=seen_dates[0] if seen_dates else "", last_visit=seen_dates[-1] if seen_dates else "",
            visit_count=sum(counter.values()),
        )

    for key, counter in contact_raw_names.items():
        contacts[key].full_name = _canonical(counter)
        result.contacts[key] = contacts[key]

    return result

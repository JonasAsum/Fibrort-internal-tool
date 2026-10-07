"""Layer 1b: contacts and companies from the message headers and signatures.

Hard data only - addresses, names, phone numbers, titles. Nothing is guessed from
meaning: a phone number is taken from a signature line that labels it, a company
comes from the folder the mail was filed in or from the address domain.
"""
import re
from collections import Counter, defaultdict

from .reader import INTERNAL_DOMAINS, OWNER_EMAIL, OWNER_NAME

FREEMAIL = {"gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "live.com", "yahoo.com", "qq.com",
            "163.com", "126.com", "foxmail.com", "sina.com", "icloud.com", "msn.com", "aliyun.com", "139.com",
            "189.cn", "sohu.com", "21cn.com", "yeah.net", "vip.163.com", "vip.qq.com",
            # US internet providers: a private address, the person is not the provider's customer contact
            "comcast.net", "charter.net", "att.net", "verizon.net", "sbcglobal.net", "bellsouth.net", "aol.com",
            "yahoo.com.cn", "yahoo.co.jp", "gmx.de", "gmx.net", "web.de", "t-online.de", "naver.com", "daum.net",
            "hanmail.net", "protonmail.com", "mail.ru", "yandex.ru"}
SYSTEM_LOCALPARTS = re.compile(r"^(no[-_.]?reply|do[-_.]?not[-_.]?reply|mailer-daemon|postmaster|notifications?|"
                               r"bounce|noreply|system|alert|newsletter)", re.I)
TLD_PARTS = {"com", "cn", "de", "org", "net", "co", "edu", "gov", "uk", "us", "eu", "hk", "tw", "jp", "kr", "at",
             "ch", "fr", "it", "es", "nl", "info", "biz", "io", "mail", "www"}

TITLE_WORDS = re.compile(
    r"(Manager|Director|Engineer|Head of|Sales|Purchas|Procurement|Managing|CEO|President|Product Management|"
    r"Order management|Auftragsmanagement|After Sales|Support|Assistant|Specialist|Controller|Buyer|Technician|"
    r"Leiter|Vertrieb|Einkauf|Geschäftsführer|总经理|经理|总监|工程师|采购|销售|主管|专员|助理|订单管理|研发|"
    r"Research & Development|Maintenance|维修)", re.I)
# footer lines that contain a title word but are an imprint, a disclaimer or a sentence, not a job title
TITLE_REJECT = re.compile(r":\s*(Dr\.|Prof\.)|Board of|Directors?\s*:|Managing Directors|Geschäftsführer|Vorstand|"
                          r"Sincerely|</?\w+>|please|kindly|if you|contact your|如有|联系|请|页面|持续|[。！？!?]|"
                          r"\w+:\s*[A-Z][a-z]+ [A-Z][a-z]+,", re.I)
COMPANY_WORDS = re.compile(r"(GmbH|AG\b|Co\.|Ltd|Inc\.?|Corp|LLC|S\.A\.|B\.V\.|有限公司|股份|集团)")
SIGNOFF = re.compile(r"^\s*(best regards|kind regards|regards|best rgds|rgds|best wishes|br,|mit freundlichen gr|freundliche gr|viele gr|"
                     r"亲切的问候|此致|thanks|thank you|br\b|cheers|saludos|cordialement)", re.I)
NOISE_LINE = re.compile(r"(USt-ID|Steuer|HRB|Registergericht|Geschäftsführer|统一社会信用代码|Bank|IBAN|BIC|"
                        r"confidential|vertraulich|www\.|http|@)", re.I)
PHONE_RE = re.compile(r"(?<![\w/])(\+?\d[\d\s\-\(\)]{6,}\d)(?![\w])")
LABEL_RE = re.compile(r"^\s*(?:(?P<label>T|Tel|Telefon|Phone|Mobile|Mob|Mobil|M|Fon|Fax|F|电话|手机|移动电话|直线|传真|"
                      r"Direct|Cell)\b\.?\s*[:：]?)", re.I)
CJK = re.compile(r"[一-鿿]{2,4}")


# vn.cncmaint@, cn.sales@: a region code in front makes an inbox, not 'Vn Cncmaint'
REGION_PREFIX = {"cn", "vn", "sh", "bj", "gz", "sz", "de", "us", "uk", "in", "jp", "kr", "th", "my", "sg", "hk", "tw",
                 "eu", "ap", "hq", "it", "fr", "es", "mx", "br", "ru"}


def clean_name(name: str, addr: str) -> str:
    name = re.sub(r"\(.*?\)|\[.*?\]", "", name or "").strip().strip('"\' ')
    name = re.sub(r"\s+", " ", name)
    if "@" in name or not name:
        local = addr.split("@")[0]
        if re.fullmatch(r"[a-z]+[._-][a-z]+", local) and re.split(r"[._-]", local)[0] not in REGION_PREFIX:
            return " ".join(p.capitalize() for p in re.split(r"[._-]", local))
        return ""
    if name.count(",") == 1 and not CJK.search(name):
        last, first = [p.strip() for p in name.split(",")]
        if last and first and len(first.split()) <= 3 and len(last.split()) <= 3:
            return f"{first} {last}"
    return name


def _name_key_words(name: str) -> list[str]:
    return re.findall(r"[a-z]+|[一-鿿]+", (name or "").lower())


def _name_key(name: str) -> str:
    """Word order does not matter: 'Costa Ana' and 'Ana Costa' are one person."""
    words = re.findall(r"[a-z]+|[一-鿿]+", name.lower())
    return "".join(sorted(words))


def is_internal(addr: str) -> bool:
    return addr.split("@")[-1] in INTERNAL_DOMAINS


def _is_real_address(addr: str) -> bool:
    return "@" in addr and not addr.startswith("imceaex-") and not is_system(addr)


def is_system(addr: str) -> bool:
    """noreply@, postmaster@, notifications@ ...: a real sender of a mail, never a person."""
    return bool(SYSTEM_LOCALPARTS.match(addr.split("@")[0]))


def resolve_addresses(messages: list[dict]) -> None:
    """Exchange writes some senders as IMCEAEX-...@domain (a legacy directory path, not a mailbox).
    The same person nearly always appears elsewhere with the real address; map by name."""
    known: dict[str, Counter] = defaultdict(Counter)
    for m in messages:
        for p in [m["from"]] + m["to"] + m["cc"]:
            nm = clean_name(p["name"], p["email"])
            if nm and _is_real_address(p["email"]):
                known[_name_key(nm)][p["email"]] += 1
    for m in messages:
        for key in ("to", "cc"):
            fixed = []
            for p in m[key]:
                fixed_p = _fix(p, known)
                if fixed_p:
                    fixed.append(fixed_p)
            m[key] = fixed
        m["from"] = _fix(m["from"], known) or {"name": "", "email": ""}


def _fix(p: dict, known) -> dict | None:
    """The address as it should be stored on the message: real and system mailboxes as written, an Exchange
    IMCEAEX path resolved to the person's real address, None when it cannot be resolved."""
    if "@" in p["email"] and is_system(p["email"]):
        return p
    return _resolve(p, known)


def _resolve(p: dict, known) -> dict | None:
    if _is_real_address(p["email"]):
        return p
    if p["email"].startswith("imceaex-"):
        nm = clean_name(p["name"], "")
        if not nm:  # IMCEAEX embeds the display name after CN=RECIPIENTS ...-LAST+2C+20FIRST
            m = re.search(r"-([A-Z0-9+]+)@", p["email"].upper())
            if m:
                tail = m.group(1).replace("+2C", ",").replace("+20", " ")
                nm = clean_name(tail.title(), "")
        cands = known.get(_name_key(nm)) if nm else None
        if cands:
            return {"name": p["name"] or nm, "email": cands.most_common(1)[0][0]}
        if nm and _name_key(nm) == _name_key(OWNER_NAME):
            return {"name": OWNER_NAME, "email": OWNER_EMAIL}
    return None


def signature_block(top: str) -> list[str]:
    lines = [l.rstrip() for l in top.split("\n")]
    idx = None
    for i, l in enumerate(lines):
        if SIGNOFF.match(l):
            idx = i
    seg = lines[idx + 1:] if idx is not None else lines[-18:]
    return [l for l in seg if l.strip()][:22]


def parse_signature(top: str, display_name: str) -> dict:
    out = {"phones": [], "title": None, "company": None, "native_name": None}
    block = signature_block(top)
    first_names = set(display_name.lower().split())
    pending_label = ""
    for line in block:
        # signatures often put the label ("M", "T") alone on one line and the number on the next
        if re.fullmatch(r"\s*(T|Tel|M|Mob|Mobile|Fax|F)\s*[:：]?\s*", line, re.I):
            pending_label = line.strip(" :：").lower()
            continue
        if NOISE_LINE.search(line) and not LABEL_RE.match(line):
            pending_label = ""
            continue
        lab = LABEL_RE.match(line)
        label = (lab.group("label") or "").lower() if lab else pending_label
        if not LABEL_RE.match(line) and not PHONE_RE.search(line):
            pending_label = ""
        if label in ("fax", "f", "传真"):
            continue
        for m in PHONE_RE.finditer(line):
            raw = m.group(1).strip()
            digits = re.sub(r"\D", "", raw)
            if not 9 <= len(digits) <= 15 or not (label or raw.startswith("+")):
                continue
            kind = "mobile" if label in ("m", "mob", "mobile", "mobil", "cell", "手机", "移动电话") else "work"
            out["phones"].append((raw, kind))
        if not out["title"] and len(line) < 90 and TITLE_WORDS.search(line) and not LABEL_RE.match(line):
            t = line.strip(" /|")
            if display_name and t.lower().startswith(display_name.split()[0].lower()) and "|" in t:
                t = t.split("|", 1)[1].strip()  # 'Dustin J. Stewart | Business Development Manager'
            if not TITLE_REJECT.search(t):
                out["title"] = t
        if not out["company"] and len(line) < 90 and COMPANY_WORDS.search(line):
            out["company"] = line.strip(" /")
        if not out["native_name"] and display_name and (set(line.lower().split()) & first_names or len(block) < 4):
            m = CJK.search(line)
            if m and len(line) < 40:
                out["native_name"] = m.group(0)
    return out


def e164(raw: str) -> str | None:
    digits = re.sub(r"\D", "", raw)
    s = raw.strip()
    if s.startswith("+"):
        num = "+" + digits
    elif digits.startswith("0086") and len(digits) >= 13:
        num = "+" + digits[2:]
    elif len(digits) == 11 and digits.startswith("1"):
        num = "+86" + digits
    elif digits.startswith("86") and len(digits) == 13:
        num = "+" + digits
    elif digits.startswith("0") and 10 <= len(digits) <= 12:
        num = "+86" + digits[1:]
    else:
        return None
    # "+86 021 3416 1967" is a common mis-write: the trunk 0 stays after the country code
    num = re.sub(r"^\+860(\d)", r"+86\1", num)
    num = re.sub(r"^\+(49|43|41)0(\d)", r"+\1\2", num)
    return num if re.fullmatch(r"\+\d{8,15}", num) else None


def domain_company_name(domain: str) -> str:
    parts = [p for p in domain.lower().split(".") if p not in TLD_PARTS]
    return (parts[-1] if parts else domain).replace("-", " ").title()


def org_domain(domain: str) -> str:
    """mail.delta.com -> delta.com, iota.com.cn stays: one organisation, one domain."""
    parts = domain.lower().split(".")
    if len(parts) >= 3 and parts[-2] in {"com", "net", "org", "co", "gov", "edu", "ac"} and len(parts[-1]) == 2:
        return ".".join(parts[-3:])
    return ".".join(parts[-2:])


COMPANY_END = re.compile(r"^(.*?(?:有限公司|GmbH(?: & Co\.? KG)?|Co\.,? ?Ltd\.?|Ltd\.?|Inc\.?|LLC|Corporation|"
                         r"Corp\.?|AG\b|S\.A\.|B\.V\.))")


def clean_company(line: str | None) -> str | None:
    """A signature line that names a company, cut after its legal form; None for addresses, copyright
    lines and sentences, which also match COMPANY_WORDS."""
    if not line:
        return None
    s = re.sub(r"\[.*?\]", "", line).strip(" ,;/")
    m = COMPANY_END.search(s)
    if m:
        s = m.group(1).strip(" ,")
    if (not s or len(s) > 70 or re.match(r"^\W*\d", s) or len(re.findall(r"\d", s)) > 3
            or re.search(r"©|版权|请|？|\?|！|，|：|地址|^另", s)):
        return None
    return s


def _domain_matches(company: str, domain: str) -> bool:
    """acme.com and zeta-drive.cn belong to the 'Acme'/'ZETA' folders; epsilonhuayu-steering.com does NOT
    belong to 'Epsilon' (a separate legal entity): a token of the domain must equal the folder name."""
    folded = re.sub(r"[^a-z0-9]", "", company.lower())
    if len(folded) < 3:
        return False
    label = org_domain(domain).split(".")[0]
    return folded == label.replace("-", "") or folded in label.split("-")


GROUP_DOMAIN = re.compile(r"(?:^|\.)(?:fibro[a-z0-9-]*|laepple[a-z0-9-]*)\.[a-z.]+$")


def is_group(domain: str) -> bool:
    """FIBRO / Läpple group companies (FIBRO China, FIBRO Inc, FIBRO India, Läpple ...): our own side."""
    d = domain.split("@")[-1].lower()
    return d in INTERNAL_DOMAINS or bool(GROUP_DOMAIN.search(d))


# --- person names: only what has the shape of a name is used as a name -------------------------------
SURNAMES = set("王李张刘陈杨黄赵吴周徐孙马朱胡郭何高林罗郑梁谢宋唐许韩冯邓曹彭曾肖田董袁潘于蒋蔡余杜叶程苏魏吕丁任沈"
               "姚卢姜崔钟谭陆汪范金石廖贾夏韦付方白邹孟熊秦邱江尹薛闫段雷侯龙史陶黎贺顾毛郝龚邵万钱严覃武戴莫孔向汤"
               "常温康施文牛樊葛邢安齐易乔伍庞颜倪庄聂章鲁岳翟殷詹申欧耿关兰焦俞左柳甘祝包宁尚符舒阮柯纪梅童凌毕单季"
               "裴霍涂成苗谷盛曲翁冉骆蓝路游辛靳管柴蒙鲍华喻祁蒲房滕屈饶解牟艾尤阳时穆农司卓古吉缪简车项连芦麦褚娄"
               "窦戚岑景党宫费卜冷晏席卫米柏宗瞿桂全佟应臧闵苟邬边卞姬师和仇栾隋商刁沙荣巫寇桑郎甄丛仲虞敖巩明佘池"
               "查麻苑迟邝蔺")
COMPOUND_SURNAMES = {"欧阳", "司马", "诸葛", "上官", "东方", "皇甫", "慕容", "令狐", "司徒", "夏侯"}
ROLE_WORDS = re.compile(
    r"\b(service|services|center|centre|team|support|sales|info|invoice|billing|accounts?|accounting|department|"
    r"dept|office|admin|purchasing|procurement|logistics|customer|hotline|notification|system|user|group|china|"
    r"shanghai|suzhou|gmbh|ltd|inc|co|company|express|import|export|airimport|inhouse|mailbox|reception|"
    r"management|consulting|automation|technology|fibro)\b|"
    r"部|中心|团队|客服|发票|结算|公司|集团|商旅|管理|报告|系统|速运|官网", re.I)
ROLE_LOCAL = re.compile(
    r"^(sales|info|service|support|admin|office|contact|enquir|inquir|invoice|billing|account|finance|hr|purchas|"
    r"marketing|export|import|logistic|order|shipping|customer|kefu|fapiao|bill|cs|team|hello|mail|webmaster|"
    r"reception|bearing|airimport|wcn-|ncs|newcustomers|srm|erp|rpa|sap|system|robot|auto)", re.I)
LATIN_NAME = re.compile(r"[A-Za-zÀ-ÿ][A-Za-zÀ-ÿ'’\-]*\.?(?: [A-Za-zÀ-ÿ][A-Za-zÀ-ÿ'’\-]*\.?){1,3}")
HAN_NAME = re.compile(r"[一-鿿]{2,4}")


HONORIFIC = re.compile(r"(小姐|先生|女士|老师|经理|总监|总|工|姐|哥)$")  # 王小姐 = 'Miss Wang', a greeting, not a name


def is_han_name(n: str) -> bool:
    return bool(HAN_NAME.fullmatch(n)) and (n[0] in SURNAMES or n[:2] in COMPOUND_SURNAMES) and not ROLE_WORDS.search(n) \
        and not HONORIFIC.search(n)


def is_latin_name(n: str) -> bool:
    return bool(LATIN_NAME.fullmatch(n)) and not ROLE_WORDS.search(n) and not re.search(r"\d", n)


# A role word as a whole later part of the address: cn-finance@, it-helpdesk@, sh.sales@ (a prefix match there
# would catch people: billy.zhang@, hrishi.k@)
ROLE_PART = {"finance", "helpdesk", "sales", "info", "service", "services", "support", "admin", "office", "invoice",
             "invoices", "billing", "account", "accounts", "accounting", "hr", "purchase", "purchasing", "marketing",
             "export", "import", "logistics", "order", "orders", "shipping", "customer", "team", "reception", "it"}


def is_role_mailbox(addr: str) -> bool:
    """sales@, info@, sales7@, 10086@, kbc100@, xyd881024@, pd@, cn-finance@, it-helpdesk@: an inbox, not a person."""
    local = addr.split("@")[0].lower()
    return bool(ROLE_LOCAL.match(local) or len(re.findall(r"\d", local)) >= 2
                or set(re.split(r"[._-]", local)[1:]) & ROLE_PART or re.split(r"[._-]", local)[0] == "it")


def _titled(n: str) -> str:
    """'Jane Doe' / 'LI Wei' stay as written unless the whole name is lower case."""
    return " ".join(w.capitalize() for w in n.split()) if n == n.lower() else n


def person_name(display: str, addr: str) -> tuple[str | None, str | None]:
    """(name, native) from a header display name; (None, None) when it is not shaped like a person's name.
    'Paul Wu', 'Rao, Raj' -> 'Raj Rao', '陈小明/Xiaoming Chen' -> ('Xiaoming Chen', '陈小明'),
    'Xiaoli Wang (王小丽)' -> ('Xiaoli Wang', '王小丽'), 'jane.doe' -> 'Jane Doe'."""
    raw = (display or "").strip().strip('"\' ')
    native = None
    m = re.search(r"[(（]([一-鿿]{2,4})[)）]", raw)
    if m and is_han_name(m.group(1)):
        native = m.group(1)
    local = addr.split("@")[0].lower()
    if raw.lower() == local or raw.lower() == addr.lower():
        raw = ""
    # Epsilon writes 'WANG Ming (JnaP/TEF JnaP/TEF1)': the bracket holds department codes, not names
    raw = re.sub(r"\s*[(（][^)）]*[)）]", "", raw).strip()
    both = re.fullmatch(r"([A-Za-z][A-Za-z'\-]+(?: [A-Za-z][A-Za-z'\-]+){1,2})\s+([一-鿿]{2,4})", raw)
    if both and is_han_name(both.group(2)) and is_latin_name(both.group(1)):  # 'Wang Zonghao 王宗浩'
        return _titled(both.group(1)), both.group(2)
    if "/" in raw:
        parts = [p.strip() for p in raw.split("/") if p.strip()]
        lat = next((p for p in parts if is_latin_name(clean_name(p, addr))), None)
        han = next((p for p in parts if is_han_name(p)), None)
        if lat:
            return _titled(clean_name(lat, addr)), han or native
        if han:
            return han, None
    n = clean_name(raw, addr)
    if is_han_name(n):
        return n, None
    if is_latin_name(n):
        return _titled(n), native
    return None, None


def signature_person(top: str) -> tuple[str | None, str | None, bool]:
    """(latin, han, strong) from the first lines after the sign-off: 'Mary 王芳', 'Paul Zhao/赵波 (GM)',
    'Mary' + next line '王芳'. strong = a two-word or bilingual name (a lone first name needs 2 sightings)."""
    lines = signature_block(top)[:3]
    latin = han = None
    for line in lines[:2]:
        t = re.sub(r"\s*[(（].*$", "", line).strip(" ,")
        for part in re.split(r"\s*/\s*|\s{2,}|\t", t):
            part = part.strip()
            # 'Mary 王芳', and the all-capitals way many sign: 'NGUYEN VAN AN'
            mm = re.fullmatch(r"((?:[A-Z][a-zà-ÿ]+|[A-Z]{2,})(?: (?:[A-Z][a-zà-ÿ]+|[A-Z]{2,})){0,2})\s*([一-鿿]{2,4})?", part)
            if mm and mm.group(1).isupper() and " " not in mm.group(1):
                mm = None  # one capitalised word is a code ('YtP/HAM', 'UAC'), not a name
            if mm and not latin and not ROLE_WORDS.search(mm.group(1)) and not SIGNOFF.match(mm.group(1)) \
                    and not re.search(r"\b(best|regards|rgds|thanks|dear|hello|hi)\b", mm.group(1), re.I):
                latin = mm.group(1)
                if mm.group(2) and is_han_name(mm.group(2)):
                    han = mm.group(2)
            elif is_han_name(part) and not han:
                han = part
        if latin or han:
            if latin and han:
                break
            continue
        break
    strong = bool(han) or bool(latin and " " in latin)
    return latin, han, strong


def internal_company(addr: str, sig_company: str | None) -> str:
    if sig_company and re.search(r"Rundtische|Weinsberg", sig_company, re.I):
        return "FIBRO Rundtische GmbH"
    if sig_company and re.search(r"Shanghai|菲博", sig_company, re.I):
        return "FIBRO Precision Components (Shanghai) Co., Ltd."
    return "FIBRO Precision Components (Shanghai) Co., Ltd." if addr.endswith(".org") else "FIBRO Rundtische GmbH"


def _choose_name(c: dict, primary: str, colleagues: frozenset = frozenset()) -> tuple[str | None, bool]:
    """(name, is_mailbox). Header names of persons first, then the sender's own signature (a lone first name
    only when seen twice), then a first.last address. A role inbox (sales@, info@, 10086@) or a handle with no
    usable name stays a mailbox: its 'name' is the address itself, nothing is invented.
    colleagues: name keys of FIBRO staff - a signature with one of those under an outside address is a quoted
    reply of ours, never that person's name (an outside address was once named after a colleague)."""
    outside = not is_internal(primary)
    sigs = [(k, n) for k, n in c["sig_persons"].most_common()
            if not (outside and _name_key(k[0] or k[1] or "") in colleagues)]
    if c["names"]:
        header = c["names"].most_common(1)[0][0]
        # the signature spells the same person out in full: 'An Nguyen' in the header, 'NGUYEN VAN AN' signed
        words = set(_name_key_words(header))
        for (lat, han, strong), n in sigs:
            if lat and strong and words and words < set(_name_key_words(lat)):
                return lat, False
        return header, False
    role = is_role_mailbox(primary)
    for (lat, han, strong), n in sigs:
        # a shared inbox (sales@) is signed by whoever answers: only a full name is trusted there
        if strong or (n >= 2 and not role and lat and len(lat) > 3):
            return (f"{lat} ({han})" if lat and han else lat or han), False
    if not is_role_mailbox(primary):
        n = clean_name("", primary)
        if n and is_latin_name(n):
            return n, False
    return None, True


def build_people(messages: list[dict]) -> tuple[list[dict], list[dict]]:
    """Returns (contacts, companies); annotates each message with `company` and `contact_emails`."""
    resolve_addresses(messages)

    dom_votes: dict[str, Counter] = defaultdict(Counter)
    addr_votes: dict[str, Counter] = defaultdict(Counter)
    company_meta: dict[str, dict] = {}
    for m in messages:
        f = m["folder"]
        comp = f["company"]
        if not comp:
            continue
        meta = company_meta.setdefault(comp, {"name": comp, "customer_types": Counter(), "business_lines": Counter(),
                                              "domains": Counter(), "folders": set(), "messages": 0})
        meta["customer_types"][f["customer_type"]] += 1
        if f["business_line"]:
            meta["business_lines"][f["business_line"]] += 1
        meta["folders"].add(f["path"])
        meta["messages"] += 1
        for p in [m["from"]] + m["to"] + m["cc"]:
            a = p["email"]
            if not a or is_internal(a) or is_system(a):
                continue
            addr_votes[a][comp] += 1
            d = a.split("@")[-1]
            if d not in FREEMAIL:
                dom_votes[org_domain(d)][comp] += 1
                meta["domains"][org_domain(d)] += 1

    # A folder holds the customer's mails but also its forwarders, consultants and FIBRO colleagues. A
    # domain belongs to the folder's company when it carries the company's name, or, failing that, when it
    # is the folder's most frequent domain and this folder is where that domain appears most.
    accepted: dict[str, str] = {}
    for comp, meta in company_meta.items():
        named = [d for d in meta["domains"] if _domain_matches(comp, d)]
        if not named and meta["domains"] and not comp.lower().startswith("fibro"):
            top = meta["domains"].most_common(1)[0][0]
            if not is_group(top) and dom_votes[top].most_common(1)[0][0] == comp:
                named = [top]
        for d in named:
            accepted.setdefault(d, comp)
    for comp, meta in company_meta.items():
        meta["domains"] = Counter({d: n for d, n in meta["domains"].items() if accepted.get(d) == comp})

    people: dict[str, dict] = {}
    sig_by_addr: dict[str, list] = defaultdict(list)
    for m in messages:
        for role in ("from", "to", "cc"):
            for p in ([m["from"]] if role == "from" else m[role]):
                a = p["email"]
                if not a or is_system(a):
                    continue
                c = people.setdefault(a, {"emails": {a}, "names": Counter(), "natives": Counter(), "messages": 0,
                                          "sent": 0, "first_seen": m["date_iso"], "last_seen": m["date_iso"],
                                          "folders": set(), "sig_persons": Counter()})
                nm, native = person_name(p["name"], a)
                if nm and not is_role_mailbox(a):
                    c["names"][nm] += 1
                if native:
                    c["natives"][native] += 1
                c["messages"] += 1
                c["folders"].add(m["folder"]["path"])
                if m["date_iso"]:
                    c["first_seen"] = min(c["first_seen"] or m["date_iso"], m["date_iso"])
                    c["last_seen"] = max(c["last_seen"] or m["date_iso"], m["date_iso"])
                if role == "from":
                    c["sent"] += 1
                    sig_by_addr[a].append(parse_signature(m["top"], nm or ""))
                    lat, han, strong = signature_person(m["top"])
                    if lat or han:
                        c["sig_persons"][(lat, han, strong)] += 1

    # Outside people seen only in the quoted header lines of a forwarded thread (Tom Smith was on an
    # earlier Cc of An Nguyen's mail). Name as written there; their company decides later whether they are kept.
    for m in messages:
        for b in m.get("quoted", []):
            addrs = b["from"] + b["to"] + b["cc"]
            if len({a.split("@")[-1] for a in addrs}) >= 8:  # a quoted mass mailing
                continue
            for a in addrs:
                if a in people or is_internal(a) or is_system(a) or a.startswith("imceaex"):
                    continue
                c = people.setdefault(a, {"emails": {a}, "names": Counter(), "natives": Counter(), "messages": 0,
                                          "sent": 0, "first_seen": m["date_iso"], "last_seen": m["date_iso"],
                                          "folders": set(), "sig_persons": Counter(), "quoted_only": True})
                nm, native = person_name((b.get("names") or {}).get(a, ""), a)
                if nm and not is_role_mailbox(a):
                    c["names"][nm] += 1
                if native:
                    c["natives"][native] += 1
                c["messages"] += 1
                c["folders"].add(m["folder"]["path"])

    # merge internal staff who use both .com and .org addresses: same person, same name
    merged: dict[str, dict] = {}
    for a, c in people.items():
        best = c["names"].most_common(1)[0][0] if c["names"] else ""
        # "Hans CN Muster" (h.muster_cn@) and "Hans US Muster" are regional mailboxes of Hans Muster
        best_key = _name_key(re.sub(r"\b([A-Z]{2}|Org|Com)\b", "", best))
        key = f"internal:{best_key}" if is_internal(a) and best_key else a
        if key in merged:
            t = merged[key]
            t["emails"] |= c["emails"]
            t["names"] += c["names"]
            t["natives"] += c["natives"]
            t["sig_persons"] += c["sig_persons"]
            t["messages"] += c["messages"]
            t["sent"] += c["sent"]
            t["folders"] |= c["folders"]
            t["first_seen"] = min(x for x in (t["first_seen"], c["first_seen"]) if x)
            t["last_seen"] = max(x for x in (t["last_seen"], c["last_seen"]) if x)
            t["sigs"] += sig_by_addr.get(a, [])
        else:
            c["sigs"] = list(sig_by_addr.get(a, []))
            merged[key] = c

    def primary_of(c: dict) -> str:
        return sorted(c["emails"], key=lambda e: (not e.endswith(".com") and is_internal(e), e))[0]

    colleagues = frozenset({_name_key(n) for c in merged.values() if any(is_internal(e) for e in c["emails"])
                            for n in c["names"]} | {_name_key(OWNER_NAME)})
    contacts = []
    phone_counts: dict[str, Counter] = {}  # contact key -> how often each number was in its signatures
    for key, c in merged.items():
        primary = primary_of(c)
        name, mailbox = _choose_name(c, primary, colleagues)
        sigs = c["sigs"]
        phones = []
        counts = phone_counts.setdefault(primary, Counter())
        for s in sigs:
            for raw, kind in s["phones"]:
                n = e164(raw)
                if not n:
                    continue
                counts[n] += 1
                if n not in {p["phone"] for p in phones}:
                    phones.append({"phone": n, "phone_type": kind, "is_primary": not phones, "position": len(phones)})
        title = Counter(s["title"] for s in sigs if s["title"]).most_common(1)
        sig_company = Counter(s["company"] for s in sigs if s["company"]).most_common(1)
        native = c["natives"].most_common(1) or [(h, n) for (l, h, st), n in c["sig_persons"].most_common()
                                                 if h and l and name and l.split()[0] == name.split()[0]][:1]
        internal = any(is_internal(e) for e in c["emails"])
        if internal:
            company = internal_company(primary, sig_company[0][0] if sig_company else None)
            company_source = "internal"
        else:
            domain = primary.split("@")[-1]
            od = org_domain(domain)
            if domain not in FREEMAIL and od in accepted:
                company, company_source = accepted[od], "folder"
            elif domain in FREEMAIL and addr_votes.get(primary):
                company, company_source = addr_votes[primary].most_common(1)[0][0], "folder"
            else:  # named later by companies.py from the domain's own evidence (or dropped by relevance.py)
                company, company_source = None, None
        full_name = name or primary
        if name and native and not CJK.search(name):
            full_name = f"{name} ({native[0][0]})"
        dom = primary.split("@")[-1]
        contacts.append({
            "key": primary, "full_name": full_name, "mailbox": mailbox,
            "org": None if (internal or dom in FREEMAIL) else org_domain(dom), "group": is_group(dom) and not internal, "emails": sorted(c["emails"], key=lambda e: (e != primary, e)),
            "phones": phones, "title": title[0][0] if title else None, "company": company,
            "company_source": company_source, "internal": internal, "messages": c["messages"],
            "sent": c["sent"], "first_seen": c["first_seen"], "last_seen": c["last_seen"],
            "folders": sorted(c["folders"]), "aliases": sorted(set(c["names"]) - {name}),
            "quoted_only": bool(c.get("quoted_only")),
        })

    # A shared inbox signed by a person who also writes from an address of their own (a shared inbox signed
    # by a colleague who also has a personal address): the inbox is no second person. It stays a nameless
    # mailbox; the signature's phones and title belong to the person.
    by_name: dict[str, list] = defaultdict(list)
    for c in contacts:
        if not c["mailbox"]:
            by_name[_name_key(c["full_name"])].append(c)
    for group in by_name.values():
        if len(group) < 2:
            continue
        words = [w for w in _name_key_words(group[0]["full_name"]) if len(w) > 1 and w.isascii()]
        own = [c for c in group if any(w in re.sub(r"[^a-z]", "", e.split("@")[0]) for e in c["emails"] for w in words)]
        if not own:
            continue
        for c in group:
            if c in own:
                continue
            person = own[0]
            have = {p["phone"] for p in person["phones"]}
            person["phones"] += [p for p in c["phones"] if p["phone"] not in have]
            person["title"] = person["title"] or c["title"]
            c.update(full_name=c["key"], mailbox=True, phones=[], title=None)

    # A number in several people's signatures: shared within one company it is a switchboard and stays
    # with all of them; across companies it was quoted from someone else's mail and stays only with the
    # person whose signatures carry it most often (with nobody on a tie).
    holders: dict[str, list] = defaultdict(list)
    for c in contacts:
        for p in c["phones"]:
            holders[p["phone"]].append(c)
    for number, cs in holders.items():
        if len(cs) < 2 or len({c["company"] for c in cs}) == 1:
            continue
        ranked = sorted(cs, key=lambda c: -phone_counts[c["key"]][number])
        top = phone_counts[ranked[0]["key"]][number]
        owner = ranked[0] if top > phone_counts[ranked[1]["key"]][number] else None
        for c in cs:
            if c is not owner:
                c["phones"] = [p for p in c["phones"] if p["phone"] != number]
    for c in contacts:
        for i, p in enumerate(c["phones"]):
            p["position"], p["is_primary"] = i, i == 0

    # companies
    companies: dict[str, dict] = {}
    for comp, meta in company_meta.items():
        companies[comp] = {
            "name": comp,
            "customer_type": meta["customer_types"].most_common(1)[0][0],
            "business_line": meta["business_lines"].most_common(1)[0][0] if meta["business_lines"] else None,
            "domains": sorted(meta["domains"]), "folders": sorted(meta["folders"]), "messages": meta["messages"],
            "origin": "folder",
        }
    for c in contacts:
        name = c["company"]
        if not name or not (c["internal"] or c["company_source"] == "folder"):
            continue
        comp = companies.setdefault(name, {"name": name, "customer_type": "internal" if c["internal"] else None,
                                           "business_line": None, "domains": [], "folders": [], "messages": 0,
                                           "origin": c["company_source"]})
        dom = c["key"].split("@")[-1]
        if not c["internal"] and dom not in FREEMAIL and org_domain(dom) not in comp["domains"]:
            comp["domains"].append(org_domain(dom))

    contact_company = {a: c["company"] for c in contacts for a in c["emails"]}
    for m in messages:
        folder_company = m["folder"]["company"]
        externals = [contact_company.get(p["email"]) for p in [m["from"]] + m["to"] + m["cc"]
                     if p["email"] and not is_internal(p["email"])]
        externals = [x for x in externals if x]
        m["company"] = folder_company or (Counter(externals).most_common(1)[0][0] if externals else None)
        m["contact_emails"] = sorted({p["email"] for p in [m["from"]] + m["to"] + m["cc"] if p["email"]})
    return sorted(contacts, key=lambda c: -c["messages"]), sorted(companies.values(), key=lambda c: -c["messages"])

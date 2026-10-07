#!/usr/bin/env python3
"""
Generates fictional weekly visit-report workbooks that mimic the layout of the
printed Excel lists (owner/timing/summary header, Mon-Fri key plan, visit table
with merged cells and colour-highlighted next-step lines), plus a ground truth
(expected parser output) as CSV.

All companies, people, phone numbers and e-mail addresses are invented.
Phone numbers use a 0000 block (not allocated), e-mails use example.com.

Usage:
  python3 generate_mock_visits.py --owners 4 --weeks 14 --year 2025 --seed 7 --out .
"""
import argparse, csv, json, os, random
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
DAYLABEL = ["Monday", "Tue", "Wed", "Thu", "Fri"]  # sic, as in the original sheets

CITIES = {
    "east": ["Shanghai", "Suzhou", "Wuxi", "Kunshan", "Changzhou"],
    "coast": ["Ningbo", "Hangzhou", "Jiaxing", "Nanjing", "Shaoxing"],
    "south": ["Shenzhen", "Dongguan", "Guangzhou", "Foshan", "Zhuhai"],
    "west": ["Chengdu", "Wuhan", "Chongqing", "Tianjin", "Qingdao"],
}
OWNER_POOL = [("Liu Yang", "east"), ("Sun Mei", "coast"), ("Zhao Lei", "south"), ("Chen Hao", "west")]
EXTRA_OWNER_NAMES = ["Wu Jing", "Xu Tao", "Ma Fang", "Zhu Qiang", "Guo Yan", "He Jun", "Lin Xin", "Gao Ming"]

PREFIX = ["Hengrui", "Tianyu", "Mingxin", "Lanxi", "Jinhai", "Ruifeng", "Zhongke", "Boyuan", "Hongda", "Kaiwei",
          "Xinda", "Yuanchuang", "Shengtai", "Huaxin", "Dongfang", "Jiahe", "Weida", "Anda", "Haoran", "Chuangyi",
          "Senlan", "Tongli", "Yuxin", "Guoteng", "Leibo", "Xingchen", "Hanyu", "Qiming", "Zhuoyue", "Bangda",
          "Lixing", "Kangtai", "Mingda", "Yongsheng", "Fuhua"]
SUFFIX = ["Automation", "Robotics", "Machinery", "Precision", "Intelligent Equipment", "Tech", "Packaging Machinery",
          "Electronics", "Intelligent Manufacturing", "Industrial Equipment"]
SURNAME = ["Wang", "Li", "Zhang", "Liu", "Chen", "Yang", "Huang", "Zhao", "Zhou", "Wu", "Xu", "Sun", "Ma", "Zhu",
           "Hu", "Guo", "He", "Lin", "Gao", "Luo"]
GIVEN_M = ["Wei", "Lei", "Hao", "Tao", "Qiang", "Jun", "Ming", "Haifeng", "Mingxin", "Shengxiang", "Jianguo", "Bin"]
GIVEN_F = ["Na", "Jing", "Yan", "Mei", "Fang", "Xin", "Lan", "Xiaoyu", "Yuting", "Qian"]
ROLES = ["purchasing manager", "design engineer", "project manager", "general manager", "technical director",
         "overseas business development manager", "sales manager", "production manager"]

INDUSTRIES = ["packaging and printing", "3C electronics", "new energy battery", "automotive BIW", "machine tool",
              "medical device", "semiconductor", "motor and engine assembly", "laboratory equipment", "welding"]
APPLICATIONS = ["assembly line", "fixture turning system", "laser cutting", "inspection station", "welding cell",
                "coil winding line", "packaging machine", "pick and place", "testing equipment"]
COMPETITORS = ["Weiss", "Sandex", "Huilin", "Gamma", "Taiwan DEX", "Genlong", "local brand"]
MODELS = ["ER.10", "ER.11", "ER.13", "ER15 D4", "ER11 D2", "VR.NC 18", "AT1600", "orange table"]
PARTS = ["motor", "gearbox", "encoder cable", "brake", "sensor"]
PROBLEMS = ["noise at high speed", "position drift after 3 months", "oil leakage at housing", "encoder error"]
PROJECT_STATES = ["still in design stage", "delayed because of budget", "running well", "waiting for approval"]
MONTHS_LONG = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
               "November", "December"]
HQ_NAMES = ["Mr.Becker", "Mr.Hartmann", "Ms.Vogel"]
DEPTS = ["BIW business", "overseas business", "R&D", "service"]
EXHIBITIONS = ["Industrial Automation Expo", "Packaging and Printing Expo", "Machine Tool Fair", "Assembly Technology Show"]

TYPO = {"focus": "foucs", "company": "compnay", "business": "buesiness", "because": "becuase",
        "customer": "custmer", "their": "thier", "quotation": "quatation", "product": "prodcut",
        "application": "aplication", "delivery": "delievery", "difficult": "diffcult"}

# public holidays in China 2025 (only used if --year 2025)
HOLIDAYS_2025 = []
for a, b in [((2025, 1, 28), (2025, 2, 4)), ((2025, 4, 4), (2025, 4, 6)), ((2025, 5, 1), (2025, 5, 5)),
             ((2025, 5, 31), (2025, 6, 2)), ((2025, 10, 1), (2025, 10, 8))]:
    d = date(*a)
    while d <= date(*b):
        HOLIDAYS_2025.append(d)
        d += timedelta(1)

TEAL = PatternFill("solid", fgColor="1E8070")
YELLOW = PatternFill("solid", fgColor="FFFF00")
HDR = PatternFill("solid", fgColor="D9D9D9")
thin = Side(style="thin", color="000000")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def ordinal(n):
    return "%d%s" % (n, "th" if 11 <= n % 100 <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th"))


DATE_STYLES = [
    lambda d: "%d-%s" % (d.day, MON[d.month - 1]),
    lambda d: "%s,%s" % (ordinal(d.day), MON[d.month - 1]),
    lambda d: "%s %d" % (MON[d.month - 1], d.day),
    lambda d: "%02d/%02d" % (d.day, d.month),
]
SHEET_STYLES = [lambda cw: "CW%d" % cw, lambda cw: "CW %d" % cw, lambda cw: "Week %d" % cw, lambda cw: "W%d" % cw]


class Gen:
    def __init__(self, a):
        self.a = a
        self.r = random.Random(a.seed)
        self.cust = []
        self.contacts = []
        self.owners = []
        self.visit_rows, self.note_rows, self.week_rows = [], [], []
        self.pending = {}  # (owner_idx, cw) -> [cust]

    # ---------- master data ----------
    def build_master(self):
        r = self.a
        names = EXTRA_OWNER_NAMES[:]
        self.r.shuffle(names)
        regions = list(CITIES)
        for i in range(self.a.owners):
            if i < len(OWNER_POOL):
                nm, reg = OWNER_POOL[i]
            else:
                nm, reg = names[i - len(OWNER_POOL)], regions[i % 4]
            self.owners.append({"name": nm, "cities": CITIES[reg], "idx": i, "dstyle": i % 4, "sstyle": i % 4,
                                "phone_style": 0})
        combos = [(p, s) for p in PREFIX for s in SUFFIX]
        self.r.shuffle(combos)
        n_cust = min(len(combos), 28 * self.a.owners)
        for i in range(n_cust):
            p, s = combos[i]
            canon = "%s %s" % (p, s)
            abbr = "%s %s." % (p, s.split()[0][:4])
            variants = [canon, p, abbr, p.upper(), "%s Co." % p]
            o = self.owners[i % len(self.owners)]
            owners = [o["idx"]]
            if self.r.random() < 0.12 and len(self.owners) > 1:  # shared customer
                owners.append(self.r.choice([x["idx"] for x in self.owners if x["idx"] != o["idx"]]))
            roll = self.r.random()
            status = "new" if roll < 0.45 else ("exist" if roll < 0.92 else "dealer")
            pr = self.r.random()
            self.cust.append({
                "id": "C%03d" % (i + 1), "name": canon, "prefix": p, "variants": variants, "owners": owners,
                "city": self.r.choice(o["cities"]), "industry": self.r.choice(INDUSTRIES),
                "potential": "A" if pr < 0.15 else ("B" if pr < 0.45 else "C"), "status": status,
                "parent": None, "visits": 0, "contacts": []})
        for c in self.cust:
            if self.r.random() < 0.12:
                c["parent"] = self.r.choice([x for x in self.cust if x is not c])
            for _ in range(self.r.choice([1, 1, 2, 3])):
                fem = self.r.random() < 0.35
                given = self.r.choice(GIVEN_F if fem else GIVEN_M)
                sur = self.r.choice(SURNAME)
                pid = "P%03d" % (len(self.contacts) + 1)
                ct = {"id": pid, "cust": c, "sur": sur, "given": given, "fem": fem,
                      "role": self.r.choice(ROLES), "phone": "%s%s0000%04d" % (
                          "1", self.r.choice(["38", "58", "36", "86", "91", "89"]), self.r.randint(0, 9999)),
                      "office": ("021-0000%04d" % self.r.randint(0, 9999)) if self.r.random() < 0.3 else ""}
                c["contacts"].append(ct)
                self.contacts.append(ct)

    # ---------- cell value helpers ----------
    def contact_display(self, ct):
        t = self.r.random()
        sal = "Ms." if ct["fem"] else "Mr."
        if t < 0.40:
            return "%s%s" % (sal, ct["sur"])
        if t < 0.80:
            return "%s %s" % (ct["sur"], ct["given"])
        if t < 0.92:
            return "%s%s" % (ct["sur"], ct["given"].lower())
        return "%s %s" % (sal, ct["sur"])

    def phone_display(self, ph):
        t = self.r.random()
        a, b, c = ph[:3], ph[3:7], ph[7:]
        if t < 0.80:
            return ph
        if t < 0.88:
            return "%s %s %s" % (a, b, c)
        if t < 0.94:
            return "%s-%s-%s" % (a, b, c)
        return "+86 %s %s %s" % (a, b, c)

    def email_for(self, ct):
        dom = ct["cust"]["prefix"].lower()
        return "%s.%s@%s.example.com" % (ct["given"].lower(), ct["sur"].lower(), dom)

    def typo(self, text):
        if self.r.random() > 0.13:
            return text
        for k, v in TYPO.items():
            if k in text:
                return text.replace(k, v, 1)
        return text

    # ---------- notes ----------
    def make_notes(self, cust, ct, owner, cw, year_cw_hint):
        r = self.r
        n = r.choice([1, 2, 2, 3, 3, 4, 4, 5])
        comp, model = r.choice(COMPETITORS), r.choice(MODELS)
        qty = r.randint(1, 20)
        sal = ("Ms." if ct["fem"] else "Mr.") + ct["sur"]
        pct = r.choice([15, 20, 30, 40, 50])
        k = cw + r.randint(1, 4)
        month = r.choice(MONTHS_LONG)
        info_new = [
            "%s is supplier in %s area, mostly %s." % (cust["prefix"], cust["industry"], r.choice(APPLICATIONS)),
            "%s is %s. He was introduced us to the %s team." % (sal, ct["role"], r.choice(DEPTS)) if not ct["fem"]
            else "%s is %s. She was introduced us to the %s team." % (sal, ct["role"], r.choice(DEPTS)),
            "They used %s table for %s. Mostly size is about %d-%dmm." % (comp, r.choice(APPLICATIONS), 200, r.choice([300, 500])),
            "The price of %s product is almost %d%% cheaper than our table, it is look like difficult to change." % (comp, pct),
            "Their business is %s than last year, many update requirement project this year." % r.choice(["better", "decreased"]),
            "They also have a %s department, but the leader is busy." % r.choice(DEPTS),
            "No business in %s area, mostly in %s." % (r.choice(["automobile", "medical", "3C"]), r.choice(["3C", "packaging", "battery"])),
            "Their customer is tier 1 and have good business in oversea.",
            "They will inquire us if they have new project.",
            "We had cooperation with another department of the group before.",
        ]
        if cust["parent"]:
            info_new.append("%s belong to %s group." % (cust["prefix"], cust["parent"]["prefix"]))
        info_exist = [
            "Project for %s is %s. Need %d pc %s." % (r.choice(APPLICATIONS), r.choice(PROJECT_STATES), qty, model),
            "Customer asked for quotation of %d pc %s, delivery time is important." % (qty, model),
            "%s offered %d%% lower price. Customer ask for discount." % (comp, pct),
            "The replacement %s has not arrived yet." % r.choice(PARTS),
            "Technical problem with %s: %s. Need support from HQ." % (model, r.choice(PROBLEMS)),
            "They plan to buy %d pcs %s per year." % (qty, model),
            "Select %d pcs %s for the new project, will get order in %s." % (qty, model, month),
            "Used %d - %d pcs %s table per year." % (qty, qty + 5, comp),
        ]
        actions = [
            ("Will send quotation for %d pc %s this week." % (qty, model), "action"),
            ("Arrange visit in CW %d." % k, "action"),
            ("Promote Fibro %s table and start contact with the customers." % model, "action"),
            ("Send comparison with %s to %s." % (comp, sal), "action"),
            ("Will go to the enduser site together with distributor in CW %d." % k, "action"),
            ("Arrange appointment with %s in %s." % (r.choice(HQ_NAMES), month), "action"),
            ("Customer provide the parameter, wait for solution and quotation from HQ.", "wait"),
            ("Wait for drawing confirmation from customer.", "wait"),
            ("Waiting for PO, customer say budget approval in %s." % month, "wait"),
        ]
        out = []
        pool_info = info_new if cust["status"] == "new" else info_exist + info_new[3:6]
        n_info = max(1, n - r.choice([0, 1, 1]))
        for t in r.sample(pool_info, min(n_info, len(pool_info))):
            out.append([t, "info"])
        while len(out) < n:
            t, kind = r.choice(actions)
            out.append([t, kind])
        r.shuffle(out)
        result = []
        for t, kind in out:
            dl = ""
            fill = ""
            if kind == "action":
                fill = "green" if r.random() < 0.7 else ""
            elif kind == "wait":
                fill = "yellow" if r.random() < 0.8 else ""
            if kind in ("action", "wait") and r.random() < 0.35:
                dl = r.choice(["Send quotation before end of this week.", "Follow up in CW %d." % k,
                               "%s will follow up this customer." % owner["name"].replace(" ", ""),
                               "Appointment should be arrange in this month.",
                               "Contact with %s." % cust["prefix"], "Reply to customer in 3 days."])
            result.append((self.typo(t), fill, dl))
        return result

    def purpose(self, cust, kind, followup):
        r = self.r
        if kind == "exhibition":
            return "Visit to get information about %s industry." % r.choice(INDUSTRIES)
        if kind == "internal":
            return r.choice(["Come to office for team meeting.", "Weekly sales meeting at office.", "Training at office."])
        if kind == "dealer":
            return r.choice(["Dealer meeting: stock and price update.", "Training for dealer sales team.",
                             "Visit dealer and check market feedback."])
        if followup:
            return r.choice(["Follow up last visit.", "Follow up quotation.", "Visit again and discuss next step."])
        if cust["status"] == "new":
            return r.choice(["Visit new customer to promote Fibro product.", "Promote VR.NC product and orange table.",
                             "Roadshow", "Visit to get information about %s industry." % cust["industry"]])
        return r.choice(["Support customer to change motor for 1 pc ER.13",
                         "New project need %d pc ER table, visit to do selection." % r.randint(2, 8),
                         "Technical exchange about %s." % r.choice(APPLICATIONS),
                         "Collect requirements for %s project." % r.choice(APPLICATIONS),
                         "Cutting customer, but there is also some automation project, visit and promote VR.NC product."])

    # ---------- workbook writing ----------
    def style_cell(self, c, fill=None, bold=False, wrap=True, center=False):
        c.alignment = Alignment(wrap_text=wrap, vertical="center" if center else "top",
                                horizontal="center" if center else None)
        c.border = BORDER
        if fill:
            c.fill = fill
        if bold:
            c.font = Font(bold=True)

    def run(self):
        self.build_master()
        os.makedirs(os.path.join(self.a.out, "excel"), exist_ok=True)
        os.makedirs(os.path.join(self.a.out, "ground_truth"), exist_ok=True)
        for o in self.owners:
            self.write_owner(o)
        self.write_truth()

    def write_owner(self, o):
        r, a = self.r, self.a
        wb = Workbook()
        wb.remove(wb.active)
        start_cw = r.randint(14, 20)
        fname = "visit_list_%s_%d.xlsx" % (o["name"].replace(" ", "_"), a.year)
        my_cust = [c for c in self.cust if o["idx"] in c["owners"]]
        dfmt, sfmt = DATE_STYLES[o["dstyle"]], SHEET_STYLES[o["sstyle"]]
        counter = 0
        for w in range(a.weeks):
            cw = start_cw + w
            if cw > 52:
                break
            monday = date.fromisocalendar(a.year, cw, 1)
            days = [monday + timedelta(i) for i in range(5)]
            if a.year == 2025:
                days = [d for d in days if d not in HOLIDAYS_2025]
            if not days:
                continue  # whole week off -> no sheet (gap in CW numbering)
            slots = []
            for c in self.pending.pop((o["idx"], cw), []):
                slots.append(("customer", c, True))
            target = max(2, round(r.randint(4, 8) * len(days) / 5))
            while len(slots) < target:
                t = r.random()
                if t < 0.05 and not any(s[0] == "exhibition" for s in slots):
                    slots.append(("exhibition", None, False))
                elif t < 0.22 and not any(s[0] == "internal" for s in slots):
                    slots.append(("internal", None, False))
                else:
                    c = r.choice(my_cust)
                    if any(s[1] is c for s in slots):
                        continue
                    slots.append(("dealer" if c["status"] == "dealer" else "customer", c, False))
            planned = []
            for kind, c, fu in slots:
                d = r.choice(days)
                if kind == "internal":
                    d = days[-1]
                planned.append((d, kind, c, fu))
            planned.sort(key=lambda x: (x[0], x[2]["city"] if x[2] else "~"))
            sname = sfmt(cw)
            ws = wb.create_sheet(sname)
            widths = [11, 13, 24, 11, 20, 15, 16, 14, 28, 40, 70, 36]
            for i, wd in enumerate(widths, 1):
                ws.column_dimensions[chr(64 + i)].width = wd
            # header block
            ws["A1"] = "Owner :"
            ws.merge_cells("B1:C1")
            ws["B1"] = o["name"]
            ws["E1"] = "Timing :"
            ws["F1"] = "CW%d" % cw
            ws.merge_cells("J1:L1")
            ws["J1"] = "Summary"
            ws["A2"] = "Item"
            ws.merge_cells("B2:I2")
            ws["B2"] = "Key Plan"
            for i in range(1, 13):
                for rr in (1, 2):
                    self.style_cell(ws.cell(rr, i), HDR if rr == 2 or i in (1, 5, 10) else None, bold=True, wrap=False)
            by_day = {}
            for d, kind, c, fu in planned:
                by_day.setdefault(d, []).append((kind, c))
            plan_text = {}
            for i, d in enumerate(days_of_week(monday)):
                rr = 3 + i
                ws.cell(rr, 1, DAYLABEL[i])
                ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=9)
                txt = ""
                if d in by_day:
                    parts = []
                    for kind, c in by_day[d]:
                        if kind == "internal":
                            parts.append("Office")
                        elif kind == "exhibition":
                            parts.append("Exhibition")
                        else:
                            parts.append("%s: %s" % (c["city"], r.choice(c["variants"][:2])))
                    txt = "; ".join(parts)
                ws.cell(rr, 2, txt)
                plan_text[DAYLABEL[i]] = txt
                for cc in range(1, 10):
                    self.style_cell(ws.cell(rr, cc), HDR if cc == 1 else None, bold=(cc == 1), wrap=False)
            header_row = 9
            heads = ["Date", "Location/ City", "Customer", "Potential (A, B, C)", "Customer (new, exist, dealer)",
                     "Name", "Phone", "Office", "Email", "Purpose for Visiting", "Next Step / Action",
                     "Dead line and follow up"]
            for i, h in enumerate(heads, 1):
                self.style_cell(ws.cell(header_row, i, h), HDR, bold=True, center=True)
            # visits
            row = header_row + 1
            n_visits = len(planned)
            n_new = n_a = 0
            green = 0
            topics = []
            for d, kind, c, fu in planned:
                counter += 1
                vid = "%s-%dCW%02d-%02d" % ("".join(x[0] for x in o["name"].split()), a.year, cw, counter % 100)
                status = "done"
                if kind == "customer" or kind == "dealer":
                    ct = r.choice(c["contacts"])
                    c["visits"] += 1
                    if c["status"] == "new" and c["visits"] > 1 and r.random() < 0.3:
                        c["status"] = "exist"
                    if r.random() < 0.10:
                        c["potential"] = {"A": "B", "B": r.choice("AC"), "C": "B"}[c["potential"]]
                    cust_cell = r.choice(c["variants"]) if r.random() < 0.35 else c["name"]
                    city = c["city"]
                    typ_norm = c["status"]
                    typ_raw = {"new": r.choice(["New", "new", "New", "new"]),
                               "exist": r.choice(["Exist", "exist", "Exist", "Exist-motion customer",
                                                  "Exist-cutting customer", "exist"]),
                               "dealer": r.choice(["Dealer", "dealer"])}[c["status"]]
                    pot = c["potential"]
                    name_cell = self.contact_display(ct)
                    phone_cell = self.phone_display(ct["phone"]) if r.random() < 0.65 else ""
                    office = ct["office"] if r.random() < 0.5 else ""
                    mail = self.email_for(ct) if r.random() < 0.25 else ""
                    purpose = self.purpose(c, kind, fu)
                    notes = self.make_notes(c, ct, o, cw, 0)
                    cust_id, cont_id = c["id"], ct["id"]
                    if r.random() < 0.05:
                        status = "cancelled"
                        k = cw + r.randint(1, 3)
                        notes = [("Cancelled by customer.", "green", ""), ("Arrange visit in CW %d" % k, "green", "")]
                    if any("Arrange visit in CW" in t for t, _, _ in notes):
                        kk = [int(x) for x in "".join(ch if ch.isdigit() else " " for ch in
                              [t for t, _, _ in notes if "Arrange visit in CW" in t][0]).split()][-1]
                        self.pending.setdefault((o["idx"], kk), []).append(c)
                    topics.append(c["industry"])
                elif kind == "exhibition":
                    cust_cell = r.choice(EXHIBITIONS)
                    city, typ_norm, typ_raw, pot = r.choice(o["cities"]), "exhibition", "New exhibition", r.choice("AB")
                    name_cell = phone_cell = office = mail = ""
                    purpose = self.purpose(None, kind, False)
                    notes = [("There are approximately dozens of equipment manufacturers doing the %s machine. Most of them used %s table." % (
                        r.choice(["packaging", "printing", "assembly"]), r.choice(COMPETITORS)), "", ""),
                        ("Promote Fibro %s table and start contact with the customers. Will arrange visit in CW %d and CW %d." % (
                            r.choice(MODELS), cw + 2, cw + 3), "green", "")]
                    cust_id = cont_id = ""
                    status = "exhibition"
                else:
                    cust_cell, city, typ_norm, typ_raw, pot = "Office " + o["cities"][0], o["cities"][0], "internal", "", ""
                    name_cell = phone_cell = office = mail = ""
                    purpose = self.purpose(None, kind, False)
                    notes = []
                    cust_id = cont_id = ""
                    status = "internal"
                if pot == "A":
                    n_a += 1
                if typ_norm == "new":
                    n_new += 1
                height = max(len(notes) + r.choice([0, 1, 2]), 2 if kind == "internal" else 3)
                r0, r1 = row, row + height - 1
                vals = [dfmt(d), city, cust_cell, pot, typ_raw, name_cell, phone_cell, office, mail, purpose]
                for ci, v in enumerate(vals, 1):
                    if height > 1:
                        ws.merge_cells(start_row=r0, start_column=ci, end_row=r1, end_column=ci)
                    ws.cell(r0, ci, v if v != "" else None)
                    for rr in range(r0, r1 + 1):
                        self.style_cell(ws.cell(rr, ci), center=ci in (1, 2, 4, 5))
                for rr in range(r0, r1 + 1):
                    self.style_cell(ws.cell(rr, 11))
                    self.style_cell(ws.cell(rr, 12))
                for li, (t, fill, dl) in enumerate(notes):
                    cell = ws.cell(r0 + li, 11, t)
                    if fill == "green":
                        cell.fill = TEAL
                        green += 1
                    elif fill == "yellow":
                        cell.fill = YELLOW
                    ws.cell(r0 + li, 12, dl if dl else None)
                    self.note_rows.append({
                        "note_id": "%s-N%d" % (vid, li + 1), "visit_id": vid, "line_no": li + 1, "text": t,
                        "highlight": fill, "deadline_followup": dl, "excel_row": r0 + li})
                phone_norm = "+86" + "".join(ch for ch in phone_cell if ch.isdigit())[-11:] if phone_cell else ""
                self.visit_rows.append({
                    "visit_id": vid, "workbook": fname, "sheet": sname, "year": a.year, "cw": cw, "owner": o["name"],
                    "date_raw": vals[0], "date_iso": d.isoformat(), "city": city, "customer_raw": cust_cell,
                    "customer_id": cust_id, "potential": pot, "customer_type_raw": typ_raw,
                    "customer_type": typ_norm, "contact_raw": name_cell, "contact_id": cont_id,
                    "phone_raw": phone_cell, "phone_e164": phone_norm, "office_phone": office, "email": mail,
                    "purpose": purpose, "status": status, "excel_row_start": r0, "excel_row_end": r1})
                row = r1 + 1
            # summary lines (J2:L7 merged per line)
            lines = ["%d visits, %d new customers, %d with potential A." % (n_visits, n_new, n_a),
                     "Open actions: %d." % green]
            if topics:
                lines.append("Main topics: %s." % ", ".join(sorted(set(topics))[:3]))
            if r.random() < 0.5:
                lines.append("Market feedback: local brands are about %d%% cheaper than our tables." % r.choice([30, 40, 50]))
            for i in range(5):
                rr = 2 + i
                ws.merge_cells(start_row=rr, start_column=10, end_row=rr, end_column=12)
                ws.cell(rr, 10, lines[i] if i < len(lines) else None)
                for cc in range(10, 13):
                    self.style_cell(ws.cell(rr, cc), HDR if rr == 2 else None, wrap=True)
            ws.cell(2, 10).value = "Summary" if False else ws.cell(2, 10).value
            self.week_rows.append({"workbook": fname, "sheet": sname, "year": a.year, "cw": cw, "owner": o["name"],
                                   "monday": monday.isoformat(), "key_plan": json.dumps(plan_text, ensure_ascii=False),
                                   "summary": json.dumps(lines, ensure_ascii=False)})
        wb.save(os.path.join(a.out, "excel", fname))

    def write_truth(self):
        gt = os.path.join(self.a.out, "ground_truth")

        def dump(name, rows, cols):
            with open(os.path.join(gt, name), "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=cols)
                w.writeheader()
                w.writerows(rows)
        dump("visits.csv", self.visit_rows, list(self.visit_rows[0]))
        dump("notes.csv", self.note_rows, list(self.note_rows[0]))
        dump("weeks.csv", self.week_rows, list(self.week_rows[0]))
        dump("customers.csv", [{
            "customer_id": c["id"], "canonical_name": c["name"], "aliases": "|".join(c["variants"][1:]),
            "city": c["city"], "industry": c["industry"], "parent_customer_id": c["parent"]["id"] if c["parent"] else "",
            "owners": "|".join(self.owners[i]["name"] for i in c["owners"])} for c in self.cust],
            ["customer_id", "canonical_name", "aliases", "city", "industry", "parent_customer_id", "owners"])
        dump("contacts.csv", [{
            "contact_id": p["id"], "customer_id": p["cust"]["id"], "canonical_name": "%s %s" % (p["sur"], p["given"]),
            "gender": "f" if p["fem"] else "m", "role": p["role"], "phone_e164": "+86" + p["phone"]} for p in self.contacts],
            ["contact_id", "customer_id", "canonical_name", "gender", "role", "phone_e164"])


def days_of_week(monday):
    return [monday + timedelta(i) for i in range(5)]


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--owners", type=int, default=4)
    ap.add_argument("--weeks", type=int, default=14)
    ap.add_argument("--year", type=int, default=2025)
    ap.add_argument("--seed", type=int, default=7)
    ap.add_argument("--out", default=".")
    a = ap.parse_args()
    g = Gen(a)
    g.run()
    print("owners", len(g.owners), "customers", len(g.cust), "visits", len(g.visit_rows), "notes", len(g.note_rows),
          "sheets", len(g.week_rows))

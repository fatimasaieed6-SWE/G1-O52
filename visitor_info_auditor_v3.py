
"""
SW350 - Group 2 | Cross-Platform Visitor Information Auditor (v3)

What the tool does
  Reads the raw dataset (Data sheet: one row per attraction x information category, the value shown on
  Google Maps, Visit Saudi and Tripadvisor, and the official value), applies the written rules of
  Table 2 (Phase 2 report), and labels every cell:
      MISSING        the platform does not provide the information (rule 1 of Table 2)
      MATCH/MISMATCH platform value vs official value (rules 2-5)
      NOT_COMPARABLE the comparison cannot be made under Table 2
  and every attraction x category across the three platforms:
      CONSISTENT / CONFLICT / NOT_COMPARABLE  (rule 6: all three platforms must report the field)

How it is evaluated (RQ4) - no circularity
  * Two team members label the same 40 rows independently in their own copy of audit_template.xlsx,
    using only the written rules, BEFORE the tool is run (they never see the tool's output).
  * Agreement between the two auditors is reported (percentage agreement and Cohen's kappa).
  * A third member resolves the auditor-vs-auditor disagreements -> "Agreed" labels.
    The agreed labels are NEVER edited to match the tool.
  * The tool's flags are compared with the agreed labels for each discrepancy type
    (MISSING, OFFICIAL_MISMATCH, CROSS_PLATFORM_CONFLICT) and with a flag-everything baseline.

Usage
  python visitor_info_auditor_v3.py --data data.xlsx --out results                      # tool only
  python visitor_info_auditor_v3.py --data data.xlsx --audit-a A.xlsx --audit-b B.xlsx \
         --agreed agreed.xlsx --out results                                               # full evaluation
  python visitor_info_auditor_v3.py --data data.xlsx --make-template audit_template.xlsx  # blind template
"""
import argparse, csv, os, re, sys
from collections import Counter

PLATFORMS = ["google_maps", "visit_saudi", "tripadvisor"]
PNAME = {"google_maps": "Google Maps", "visit_saudi": "Visit Saudi", "tripadvisor": "Tripadvisor"}
PCOL = {"google_maps": 8, "visit_saudi": 9, "tripadvisor": 10}        
FIELDS = ["opening_hours", "ticket_prices", "location_entrance", "booking_channels"]
FNAME = {"opening_hours": "Opening hours", "ticket_prices": "Ticket prices",
         "location_entrance": "Location and entrance", "booking_channels": "Booking channels"}
FMAP = {"opening hours": "opening_hours", "ticket prices": "ticket_prices",
        "location and entrance details": "location_entrance", "booking channels": "booking_channels"}
MISS, M, X, NC = "MISSING", "MATCH", "MISMATCH", "NOT_COMPARABLE"
CON, CONF = "CONSISTENT", "CONFLICT"
PEND = "PENDING"
USD_TO_SAR = 3.75           

AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
def clean(s):
    s = str(s if s is not None else "").translate(AR_DIGITS)
    return re.sub(r"\s+", " ", s).strip()
def is_pending(v):  return clean(v).upper().startswith("PENDING")

THIRD_PARTY = r"third[- ]party|tours? (?:&|and) experiences|featured tours?|tour product|experiences only|tours only"
BOOKING_ELEMENT = r"ticket card|buy tickets?|book now|online booking|book(?:ing)? (?:link|button)|reserve|tickets can be purchased"
def is_missing(field, v, official=False):
    s = clean(v).lower()
    if s in ("", "-", "n/a", "na", "none", "not available", "غير متوفر", "لا يوجد") or s.startswith("not found"):
        return True
    if s.startswith("booking unavailable"):
        return True
    if official: return False
    if field == "ticket_prices" and re.search(THIRD_PARTY, s):
        return True                      
    if field == "booking_channels" and not re.search(BOOKING_ELEMENT, s):
        return True                      
    return False


DAYS = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"]
DAYRE = r"(sun|mon|tue|wed|thu|fri|sat)[a-z]*"
TOK = r"(?:midnight|noon|\d{1,2}(?:[:.]\d{2})?\s*(?:[ap]\.?m\.?)?)"
HOURS_NOT_COMPARABLE = r"prayer|dhuhr|asr|maghrib|no fixed|seasonal|ramadan|event only|subject to schedule"

def _minutes(tok, fallback_ampm=None):
    tok = tok.strip().lower()
    if tok == "midnight": return 0
    if tok == "noon": return 720
    m = re.fullmatch(r"(\d{1,2})(?:[:.](\d{2}))?\s*([ap])?\.?m?\.?", tok)
    h, mi, ap = int(m.group(1)), int(m.group(2) or 0), (m.group(3) or fallback_ampm)
    if ap == "p" and h < 12: h += 12
    if ap == "a" and h == 12: h = 0
    return (h % 24) * 60 + mi

def parse_ranges(text):
    out = []
    for m in re.finditer(rf"({TOK})\s*(?:-|–|—|to)\s*({TOK})", text):
        a, b = m.group(1), m.group(2)
        ap_b = re.search(r"([ap])\.?m", b.lower())
        ta = _minutes(a, ap_b.group(1) if ap_b and not re.search(r"[ap]\.?m", a.lower()) else None)
        tb = _minutes(b)
        out.append(f"{ta // 60:02d}:{ta % 60:02d}-{tb // 60:02d}:{tb % 60:02d}")
    return out

def expand_days(part):
    part = part.lower()
    if re.search(r"daily|every day|7 days", part): return list(DAYS)
    days = []
    for grp in re.split(r"\band\b|&|,|/", part):
        names = [m.group(1) for m in re.finditer(DAYRE, grp)]
        if len(names) == 2 and re.search(r"[-–—]|\bto\b", grp):
            i, j = DAYS.index(names[0]), DAYS.index(names[1])
            while True:
                days.append(DAYS[i])
                if i == j: break
                i = (i + 1) % 7
        else:
            days += names
    return days

DAYSPEC = rf"(?:daily|every day|{DAYRE}(?:\s*(?:-|–|—|to|&|and|,)\s*{DAYRE})*)"
def parse_hours(s):
    """-> {day: 'HH:MM-HH:MM|...' or 'closed'}; days not stated are absent.
    The text is cut at every day specification ('Sat-Wed', 'Thursday & Friday', 'Daily'); the times that
    follow a specification apply to all of its days."""
    s = clean(s).lower()
    s = re.sub(r"\(?last entry[^;)]*\)?|new visitors[^;]*", " ", s)
    s = re.sub(r"\([^)]*\)", " ", s)
    specs = list(re.finditer(rf"\b{DAYSPEC}\b", s))
    sched = {}
    for i, m in enumerate(specs):
        rest = s[m.end(): specs[i + 1].start() if i + 1 < len(specs) else len(s)]
        days = expand_days(m.group(0))
        rng = parse_ranges(rest)
        val = "|".join(sorted(rng)) if rng else ("closed" if "closed" in rest else None)
        if not days or val is None: continue
        for d in days: sched[d] = val
    if not sched:
        m = re.search(rf"from:?\s*({TOK})[;,\s]+to:?\s*({TOK})", s)
        if m:
            rng = parse_ranges(f"{m.group(1)} - {m.group(2)}")
            if rng: sched = {d: rng[0] for d in DAYS}
    return sched

def compare_hours(a, b):
    """MATCH if every day stated by both sources has the same times; NOT_COMPARABLE if no common day."""
    pa, pb = parse_hours(a), parse_hours(b)
    common = set(pa) & set(pb)
    if not common: return NC
    return M if all(pa[d] == pb[d] for d in common) else X

def parse_price(s):
    """-> ('free'|'exact'|'from', value in SAR) or (None, None)"""
    t = clean(s).lower().replace(",", "")
    m = re.search(r"adults?[^0-9]{0,20}(?:sar|sr)?\s*(\d+(?:\.\d+)?)", t)
    if m: return "exact", float(m.group(1))
    m = re.search(r"(?:from|starts? (?:at|from)|starting)\s*(?:at\s*)?(sar|sr|us\$|usd|\$)?\s*(\d+(?:\.\d+)?)", t)
    if m:
        v = float(m.group(2)); return "from", v * USD_TO_SAR if m.group(1) in ("us$", "usd", "$") else v
    m = re.search(r"(?:sar|sr)\s*(\d+(?:\.\d+)?)|(\d+(?:\.\d+)?)\s*(?:sar|sr|riyal)", t)
    if m: return "exact", float(m.group(1) or m.group(2))
    m = re.search(r"(?:us\$|usd|\$)\s*(\d+(?:\.\d+)?)", t)
    if m: return "exact", float(m.group(1)) * USD_TO_SAR
    if re.search(r"free admission|\bfree\b|don.t need tickets?|without a ticket|no ticket", t): return "free", 0.0
    return None, None

def compare_price(plat, off):
    kp, vp = parse_price(plat); ko, vo = parse_price(off)
    if kp is None or ko is None: return NC
    if "from" in (kp, ko) and kp != ko: return NC        
    return M if abs(vp - vo) < 0.01 else X

# ============================================================== Rule 4 - location and entrance (key tokens)
ALIASES = {  # spelling variants (Arabic / English) -> one token
    "makkah": ["makkah", "mecca", "مكة"], "jeddah": ["jeddah", "jiddah", "جدة"], "riyadh": ["riyadh", "الرياض"],
    "buraidah": ["buraidah", "buraydah", "buraida", "بريدة"], "qassim": ["qassim", "qassem", "القصيم"],
    "baha": ["baha", "bahah", "الباحة"], "alula": ["alula", "al ula", "العلا"], "dhahran": ["dhahran", "الظهران"],
    "diriyah": ["diriyah", "dir'iyah", "الدرعية"], "turaif": ["turaif", "traif", "الطريف"],
    "murabba": ["murabba", "مربع", "المربع"], "olaya": ["olaya", "olaia", "العليا"], "centre": ["centre", "center"],
    "corniche": ["corniche", "كورنيش"], "kurnaysh": ["kurnaysh", "kurnish"], "hegra": ["hegra", "al hijr", "الحجر"],
}
CITIES = {"makkah", "jeddah", "riyadh", "buraidah", "baha", "alula", "dhahran"}
REGIONS = {"qassim", "eastern", "province", "region"}
GENERIC = set("""saudi arabia kingdom of the and al el ad at in on to via near next off just located area access
city district governorate road rd street st avenue ave way bridge br highway hwy ring
map listing get directions link shown location listed approx m sa ksa north south east west
meeting point parking free visitor""".split()) - {"kingdom"}

def loc_tokens(s):
    t = clean(s).lower()
    for key, al in ALIASES.items():
        for a in sorted(al, key=len, reverse=True):
            t = re.sub(rf"(?<![a-z]){re.escape(a)}(?![a-z])", f" {key} ", t)
    t = re.sub(r"\bal[- ]", " ", t)
    t = re.sub(r"[^a-z؀-ۿ ]", " ", t)               # drop digits (postal codes, numbers) and punctuation
    toks = {w for w in t.split() if len(w) > 1 and w not in GENERIC}
    return toks

def split_loc(s):
    toks = loc_tokens(s)
    return toks & CITIES, toks - CITIES - REGIONS

def compare_location(plat, off):
    """MATCH: the platform names at least one of the official source's key tokens below city level
       (district, street, landmark or entrance); if the official source gives only a city, naming that city
       is enough. A platform that gives only a city/region when the official source is more specific = MISMATCH."""
    oc, osub = split_loc(off); pc, psub = split_loc(plat)
    if not oc and not osub: return NC
    if osub: return M if osub & psub else X
    return M if oc & pc else X

# ============================================================== Rule 5 - booking channels (category)
def book_cat(s):
    t = clean(s).lower()
    if re.search(r"no booking|not required|without (?:a )?booking", t): return "none"
    if re.search(r"on[- ]site only|at the venue only|reception", t) and not re.search(r"online|website|ticket card", t): return "onsite"
    if re.search(r"official|ithra\.com|experience ?alula|tickets\.|our website|online through|ticket card|book now", t): return "official"
    if re.search(THIRD_PARTY, t): return "third_party"
    return None                                               # cannot classify -> not comparable / manual review

def compare_booking(plat, off):
    cp, co = book_cat(plat), book_cat(off)
    if cp is None or co is None: return NC
    return M if cp == co else X

# ============================================================== official vs platform
def official_not_comparable(field, off):
    o = clean(off).lower()
    if field == "opening_hours" and re.search(HOURS_NOT_COMPARABLE, o): return True
    return False

def label_cell(field, plat, off):
    if is_missing(field, plat): return MISS
    if is_missing(field, off, official=True) or official_not_comparable(field, off): return NC
    return {"opening_hours": compare_hours, "ticket_prices": compare_price,
            "location_entrance": compare_location, "booking_channels": compare_booking}[field](plat, off)

# ============================================================== Rule 6 - cross-platform consistency
def label_cross(field, vals):
    if any(is_missing(field, v) for v in vals): return NC                    # all three must report the field
    if field == "opening_hours":
        sch = [parse_hours(v) for v in vals]
        common = set(sch[0]) & set(sch[1]) & set(sch[2])
        if not common: return NC
        return CON if all(sch[0][d] == sch[1][d] == sch[2][d] for d in common) else CONF
    if field == "ticket_prices":
        pr = [parse_price(v) for v in vals]
        if any(k is None for k, _ in pr) or len({k for k, _ in pr}) > 1: return NC
        return CON if max(v for _, v in pr) - min(v for _, v in pr) < 0.01 else CONF
    if field == "location_entrance":
        parts = [split_loc(v) for v in vals]
        subs = [s for _, s in parts]
        if all(not s for s in subs):
            cities = [c for c, _ in parts]
            if any(not c for c in cities): return NC
            return CON if cities[0] & cities[1] & cities[2] else CONF
        return CON if all(subs) and (subs[0] & subs[1] & subs[2]) else CONF
    cats = [book_cat(v) for v in vals]
    if None in cats: return NC
    return CON if len(set(cats)) == 1 else CONF

# ============================================================== I/O
def load_data(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True)["Data"]
    rows = []
    for r in range(2, ws.max_row + 1):
        if not ws.cell(r, 2).value: continue
        rows.append({"att": int(ws.cell(r, 1).value), "name": clean(ws.cell(r, 2).value),
                     "field": FMAP[clean(ws.cell(r, 5).value).lower()], "official": clean(ws.cell(r, 6).value),
                     "vals": {p: clean(ws.cell(r, c).value) for p, c in PCOL.items()},
                     "date": ws.cell(r, 11).value})
    return rows

LABELS_IN = {"missing": MISS, "match": M, "mismatch": X, "not comparable": NC, "not_comparable": NC,
             "consistent": CON, "same": CON, "conflict": CONF, "different": CONF, "inconsistent": CONF}
def load_audit(path):
    """audit workbook (sheet 'Audit'): Attr | Attraction | Category | ... | GM label | VS label | TA label | Cross label"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    if "Audit" not in wb.sheetnames:                      # labels written in the Data sheet (columns P-S)
        ws = wb["Data"]; out = {}
        for r in range(2, ws.max_row + 1):
            if not ws.cell(r, 2).value: continue
            a, f = int(ws.cell(r, 1).value), FMAP[clean(ws.cell(r, 5).value).lower()]
            for p, c in zip(PLATFORMS, (16, 17, 18)):
                out[(a, f, p)] = LABELS_IN.get(clean(ws.cell(r, c).value).lower())
            out[(a, f, "cross")] = LABELS_IN.get(clean(ws.cell(r, 19).value).lower())
        return out
    ws = wb["Audit"]
    hdr = [clean(c.value).lower() for c in ws[1]]
    col = lambda name: hdr.index(name) + 1
    out = {}
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value in (None, ""): continue
        a, f = int(ws.cell(r, 1).value), FMAP[clean(ws.cell(r, 3).value).lower()]
        for p in PLATFORMS:
            out[(a, f, p)] = LABELS_IN.get(clean(ws.cell(r, col(f"{PNAME[p].lower()} label")).value).lower())
        out[(a, f, "cross")] = LABELS_IN.get(clean(ws.cell(r, col("cross-platform label")).value).lower())
    return out

def make_template(rows, path):
    """Blind template for the auditors: values only, no labels from anyone, drop-down lists."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Audit"
    hdr = ["Attr", "Attraction", "Category", "Official value", "Google Maps value", "Visit Saudi value",
           "Tripadvisor value", "Google Maps label", "Visit Saudi label", "Tripadvisor label",
           "Cross-platform label", "Auditor comment"]
    ws.append(hdr)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for r in rows:
        ws.append([r["att"], r["name"], {"opening_hours": "Opening hours", "ticket_prices": "Ticket prices",
                   "location_entrance": "Location and entrance details", "booking_channels": "Booking channels"}[r["field"]],
                   r["official"], r["vals"]["google_maps"], r["vals"]["visit_saudi"], r["vals"]["tripadvisor"]])
    dv1 = DataValidation(type="list", formula1='"Missing,Match,Mismatch,Not comparable"', allow_blank=True)
    dv2 = DataValidation(type="list", formula1='"Consistent,Conflict,Not comparable"', allow_blank=True)
    ws.add_data_validation(dv1); ws.add_data_validation(dv2)
    dv1.add(f"H2:J{len(rows) + 1}"); dv2.add(f"K2:K{len(rows) + 1}")
    for i, w in enumerate([6, 24, 18, 40, 34, 34, 34, 14, 14, 14, 16, 30], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for c in row[7:11]:
            c.fill = PatternFill("solid", fgColor="FFF2CC")
    ws.freeze_panes = "H2"
    ins = wb.create_sheet("Rules")
    rules = [
        ("How to label", "Fill only the yellow columns. Use ONLY these written rules. Do not look at the other auditor's file or the tool's output."),
        ("Platform label", "Missing = the platform does not give the information. Match / Mismatch = platform vs official value under the category rule. Not comparable = the official value is missing or the category rule says the two cannot be compared."),
        ("All categories", "Empty, 'NOT FOUND' or 'not available' = Missing. If the official value is missing, a platform cell that has a value is Not comparable."),
        ("Opening hours", "Convert to 24-hour times per day. Match if the times are the same for every day that both sources state. Not comparable if the official hours are seasonal/event hours or defined relative to prayer times."),
        ("Ticket prices", "Use the standard adult admission price in SAR (free = 0; US$ x 3.75). Prices of third-party tours are not admission prices -> Missing. Match if equal. Not comparable if one source gives only a starting price ('from SAR …') or a different ticket type."),
        ("Location and entrance", "Key tokens = district, street, landmark or entrance names (ignore postal codes, word order, 'Saudi Arabia', 'Road'). Match if the platform names at least one key token of the official location below city level; if the official source gives only a city, naming that city is enough. A platform that gives only a city or region when the official source is more specific = Mismatch."),
        ("Booking channels", "Categories: official website/app, third-party platform, on-site only, no booking needed. A channel exists only if the platform shows a ticket or booking element (ticket card, Buy Tickets, Book now). A plain website link or generic 'Tours & experiences' = Missing. Match if both are in the same category."),
        ("Cross-platform label", "Consistent = all three platforms give the field and agree under the category rule. Conflict = all three give it but at least one disagrees. Not comparable = at least one platform does not give the field."),
    ]
    for r in rules: ins.append(list(r))
    ins.column_dimensions["A"].width = 22; ins.column_dimensions["B"].width = 120
    for row in ins.iter_rows():
        for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")
        row[0].font = Font(bold=True)
    wb.save(path)

# ============================================================== metrics
def pct(n, d): return None if not d else round(100 * n / d, 1)
def fmt(n, d): return f"{n}/{d} ({pct(n, d)}%)" if d else f"{n}/0 (n/a)"
def kappa(a, b):
    keys = [k for k in a if a[k] is not None and b.get(k) is not None]
    if not keys: return None, None, 0
    n = len(keys); po = sum(a[k] == b[k] for k in keys) / n
    ca, cb = Counter(a[k] for k in keys), Counter(b[k] for k in keys)
    pe = sum(ca[l] * cb[l] for l in set(ca) | set(cb)) / (n * n)
    return po, (None if pe == 1 else (po - pe) / (1 - pe)), n
def prf(tp, fp, fn):
    p = tp / (tp + fp) if tp + fp else None
    r = tp / (tp + fn) if tp + fn else None
    f = 2 * p * r / (p + r) if p and r else (0.0 if p is not None and r is not None else None)
    return p, r, f
def f3(x): return "n/a" if x is None else f"{x:.3f}"

def quality(labels, cells_keys, cross_keys):
    """RQ1-RQ3 from a label set {(att, field, platform|'cross'): label}"""
    res = {}
    for scope in ["ALL"] + PLATFORMS + FIELDS:
        ks = [k for k in cells_keys if scope == "ALL" or scope in (k[1], k[2])]
        ks = [k for k in ks if labels.get(k) not in (None, PEND)]
        req = len(ks); av = sum(labels[k] != MISS for k in ks)
        mt = sum(labels[k] == M for k in ks); mm = sum(labels[k] == X for k in ks)
        res[scope] = {"completeness": (av, req), "concordance": (mt, mt + mm)}
    for scope in ["ALL"] + FIELDS:
        ks = [k for k in cross_keys if scope == "ALL" or k[1] == scope]
        cn = sum(labels[k] == CON for k in ks); cf = sum(labels[k] == CONF for k in ks)
        res[scope]["consistency"] = (cn, cn + cf)
    return res

def evaluate(tool, ref, cell_keys, cross_keys):
    """Tool vs reference (agreed manual labels), per discrepancy type, with a flag-everything baseline."""
    out = []
    groups = {
        "MISSING": ([k for k in cell_keys if ref.get(k)], lambda l: l == MISS),
        "OFFICIAL_MISMATCH": ([k for k in cell_keys if ref.get(k) in (M, X)], lambda l: l == X),
        "CROSS_PLATFORM_CONFLICT": ([k for k in cross_keys if ref.get(k) in (CON, CONF)], lambda l: l == CONF),
    }
    for name, (elig, pos) in groups.items():
        for scope in ["ALL"] + FIELDS:
            el = [k for k in elig if scope == "ALL" or k[1] == scope]
            refpos = {k for k in el if pos(ref[k])}; flag = {k for k in el if pos(tool[k])}
            tp, fp, fn = len(refpos & flag), len(flag - refpos), len(refpos - flag)
            p, r, f = prf(tp, fp, fn)
            bp, br, bf = prf(len(refpos), len(el) - len(refpos), 0) if el else (None, None, None)
            out.append(dict(type=name, scope=scope, eligible=len(el), ref_pos=len(refpos), flagged=len(flag),
                            TP=tp, FP=fp, FN=fn, precision=p, recall=r, F1=f,
                            baseline_precision=bp, baseline_recall=br, baseline_F1=bf))
    return out

# ============================================================== main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True); ap.add_argument("--out", default="results_v3")
    ap.add_argument("--audit-a"); ap.add_argument("--audit-b"); ap.add_argument("--agreed")
    ap.add_argument("--make-template")
    a = ap.parse_args()
    rows = load_data(a.data)
    if a.make_template:
        make_template(rows, a.make_template); print("template ->", a.make_template); return
    pending = [(r["att"], r["field"], p) for r in rows for p in ["official"] + PLATFORMS
               if is_pending(r["official"] if p == "official" else r["vals"][p])]
    if pending:
        print(f"WARNING: {len(pending)} PENDING value(s) - results are provisional until they are re-captured:")
        for k in pending: print("   ", k)
    os.makedirs(a.out, exist_ok=True); P = lambda n: os.path.join(a.out, n)

    tool, cell_keys, cross_keys = {}, [], []
    for r in rows:
        f = r["field"]
        for p in PLATFORMS:
            k = (r["att"], f, p); cell_keys.append(k)
            if is_pending(r["vals"][p]): tool[k] = PEND                       # excluded until re-captured
            elif is_pending(r["official"]) and not is_missing(f, r["vals"][p]): tool[k] = NC
            else: tool[k] = label_cell(f, r["vals"][p], r["official"])
        k = (r["att"], f, "cross"); cross_keys.append(k)
        vals = [r["vals"][p] for p in PLATFORMS]
        tool[k] = NC if any(is_pending(v) for v in vals) else label_cross(f, vals)

    with open(P("tool_labels.csv"), "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh); w.writerow(["attraction", "name", "category", "official", "google_maps", "visit_saudi",
                                        "tripadvisor", "GM_label", "VS_label", "TA_label", "cross_label"])
        for r in rows:
            f = r["field"]
            w.writerow([r["att"], r["name"], FNAME[f], r["official"]] + [r["vals"][p] for p in PLATFORMS]
                       + [tool[(r["att"], f, p)] for p in PLATFORMS] + [tool[(r["att"], f, "cross")]])

    def report(title, labels):
        q = quality(labels, cell_keys, cross_keys)
        print(f"\n== {title}")
        print(f"  {'scope':22s}{'RQ1 completeness':>22s}{'RQ2 concordance':>22s}{'RQ3 consistency':>22s}")
        for s in ["ALL"] + PLATFORMS + FIELDS:
            c = q[s]
            print(f"  {s:22s}{fmt(*c['completeness']):>22s}{fmt(*c['concordance']):>22s}"
                  f"{(fmt(*c['consistency']) if 'consistency' in c else ''):>22s}")
        return q

    q_tool = report("Tool labels", tool)
    summary = [["source", "scope", "measure", "numerator", "denominator", "percent"]]
    for s, c in q_tool.items():
        for m_, (n, d) in c.items(): summary.append(["tool", s, m_, n, d, pct(n, d)])

    if a.agreed:
        ref = load_audit(a.agreed)
        q_ref = report("Agreed manual labels (reference) - use these for RQ1-RQ3 in the report", ref)
        for s, c in q_ref.items():
            for m_, (n, d) in c.items(): summary.append(["agreed", s, m_, n, d, pct(n, d)])
        if a.audit_a and a.audit_b:
            A, B = load_audit(a.audit_a), load_audit(a.audit_b)
            print("\n== Inter-auditor agreement (before adjudication)")
            agr = [["set", "n", "percent_agreement", "cohen_kappa"]]
            for name, keys in (("platform cells", cell_keys), ("cross-platform", cross_keys)):
                po, k, n = kappa({x: A.get(x) for x in keys}, {x: B.get(x) for x in keys})
                print(f"  {name:16s} n={n:3d}  agreement={f3(po)}  kappa={f3(k)}"); agr.append([name, n, po, k])
            with open(P("inter_auditor_agreement.csv"), "w", newline="", encoding="utf-8-sig") as fh:
                csv.writer(fh).writerows(agr)
            with open(P("auditor_disagreements.csv"), "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh); w.writerow(["attraction", "category", "platform", "auditor_A", "auditor_B"])
                for x in cell_keys + cross_keys:
                    if A.get(x) != B.get(x): w.writerow([x[0], FNAME[x[1]], x[2], A.get(x), B.get(x)])
        ev = evaluate(tool, ref, cell_keys, cross_keys)
        print("\n== RQ4 tool vs agreed labels, with flag-everything baseline")
        print(f"  {'type':25s}{'scope':20s}{'elig':>5}{'ref+':>5}{'flag':>5}{'TP':>4}{'FP':>4}{'FN':>4}"
              f"{'P':>7}{'R':>7}{'F1':>7}{'base P':>8}{'base F1':>8}")
        for e in ev:
            print(f"  {e['type']:25s}{e['scope']:20s}{e['eligible']:>5}{e['ref_pos']:>5}{e['flagged']:>5}{e['TP']:>4}"
                  f"{e['FP']:>4}{e['FN']:>4}{f3(e['precision']):>7}{f3(e['recall']):>7}{f3(e['F1']):>7}"
                  f"{f3(e['baseline_precision']):>8}{f3(e['baseline_F1']):>8}")
        with open(P("evaluation_v3.csv"), "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=list(ev[0])); w.writeheader(); w.writerows(ev)
        labs = [MISS, M, X, NC]
        conf = Counter((ref.get(k), tool[k]) for k in cell_keys)
        with open(P("confusion_platform_cells.csv"), "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh); w.writerow(["agreed \\ tool"] + labs)
            for l in labs: w.writerow([l] + [conf[(l, t)] for t in labs])
        with open(P("tool_vs_agreed_differences.csv"), "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh); w.writerow(["attraction", "category", "platform", "agreed", "tool"])
            for k in cell_keys + cross_keys:
                if ref.get(k) != tool[k]: w.writerow([k[0], FNAME[k[1]], k[2], ref.get(k), tool[k]])
        print("\n  Note: differences in tool_vs_agreed_differences.csv are tool errors to REPORT, not labels to change.")
    with open(P("quality_summary.csv"), "w", newline="", encoding="utf-8-sig") as fh:
        csv.writer(fh).writerows(summary)
    print(f"\noutputs -> {a.out}/")

if __name__ == "__main__":
    main()

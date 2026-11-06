"""Build the tracker's data files from the England and Wales welfare workbooks, one pair per year.

    pip install openpyxl
    python3 build_data.py

Reads  ../new_data/A Tale of Two Cities Welfare Analysis 2026 Provisional.xlsx          (2026: England, 543 seats)
       ../new_data/A Tale of Two Cities Welfare Analysis 2026 - Wales Provisional.xlsx  (2026: Wales, 32 seats)
       ../new_data/A Tale of Two Cities Welfare Analysis 2026 Revised.xlsx              (2025: England)
       ../new_data/A Tale of Two Cities Welfare Analysis 2026 - Wales Revised.xlsx      (2025: Wales)
           Each year matches pay, tax rules and welfare to one tax year: 2026 uses ONS ASHE April 2026
           (provisional), 2025/26 rules and Stat-Xplore Mar-26; 2025 uses ASHE April 2025 (revised),
           2024/25 rules and Stat-Xplore Feb-25 (the quarterly series have no Mar-25).
       ../Westminster_Parliamentary_Constituencies_July_2024_Boundaries_UK_BUC_*.geojson
       ../wales_pcon2024_simplified.geojson   Welsh boundaries, simplified from wales_pcon2024_bgc.geojson with
                                              npx mapshaper wales_pcon2024_bgc.geojson -simplify dp interval=300 keep-shapes
                                              -o wales_pcon2024_simplified.geojson precision=0.00001
       UK Parliament Members API (current MPs, fetched live)
Writes data.js   window.WT  = {latest: 2026, years: {2026: {england, wales, regions, parties, seats, dumbbell, method, ...},
                                                      2025: {...}}}
       geo.js    window.GEO = constituency boundaries, pre-projected to a flat plane
                             (x = lon * cos 52.7deg, y = lat) so the page can draw them with
                             d3.geoIdentity and no winding-order issues.

If a Wales workbook was saved without calculated values (openpyxl reads its formula cells as
empty until Excel has opened and saved it), its results are recomputed with model.py, which
follows the workbook's own formulas.

Re-run whenever a workbook changes. The page reads only data.js and geo.js.
Set WT_ROOT to read the input files from another folder.
"""
import glob, json, math, os, re, unicodedata, urllib.request
from datetime import date

import openpyxl

import model

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.environ.get("WT_ROOT", os.path.join(HERE, ".."))
DATA = os.path.join(ROOT, "new_data")
# tracker year -> England workbook, Wales workbook, and the pay, tax and welfare periods behind them (for credits and notes)
YEARS = {
    2026: {"xlsx": "A Tale of Two Cities Welfare Analysis 2026 Provisional.xlsx",
           "wales": "A Tale of Two Cities Welfare Analysis 2026 - Wales Provisional.xlsx",
           "ashe": "ASHE 2026 (provisional)", "payDate": "April 2026", "taxYear": "2025/26", "welDate": "March 2026"},
    2025: {"xlsx": "A Tale of Two Cities Welfare Analysis 2026 Revised.xlsx",
           "wales": "A Tale of Two Cities Welfare Analysis 2026 - Wales Revised.xlsx",
           "ashe": "ASHE 2025 (revised)", "payDate": "April 2025", "taxYear": "2024/25", "welDate": "February 2025"},
}
LATEST = max(YEARS)
# Stat-Xplore extract and MP party check behind both years' workbooks
WELFARE_PULLED = "6 November 2026"
PARTY_AS_OF = "November 2026"
GEOJSON = glob.glob(os.path.join(ROOT, "Westminster_Parliamentary_Constituencies_July_2024_*.geojson"))[0]
WALES_GEOJSON = os.path.join(ROOT, "wales_pcon2024_simplified.geojson")
MEMBERS_URL = "https://members-api.parliament.uk/api/Members/Search?House=1&IsCurrentMember=true&skip={}&take=20"

REGIONS = ["North East", "North West", "Yorkshire and The Humber", "East Midlands", "West Midlands",
           "East of England", "London", "South East", "South West"]
BENEFITS = ["Jobseeker's Allowance", "Universal Credit", "State Pension", "Pension Credit", "Housing Benefit",
            "Personal Independence Payment", "Disability Living Allowance", "Employment and Support Allowance",
            "Carer's Allowance", "Attendance Allowance"]
PCTS = [10, 20, 25, 30, 40, 60, 70, 75]


def key(name):
    s = unicodedata.normalize("NFKD", str(name)).encode("ascii", "ignore").decode().lower()
    s = s.replace("&", " and ")
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def num(v, nd=None):
    if v is None or isinstance(v, str):
        return None
    return round(float(v), nd) if nd is not None else float(v)


# ---------- MPs: UK Parliament Members API (fetched once, shared by both years) ----------
mps = {}
skip = 0
while True:
    with urllib.request.urlopen(MEMBERS_URL.format(skip), timeout=30) as resp:
        page = json.load(resp)
    for it in page["items"]:
        v = it["value"]
        m = v["latestHouseMembership"]
        # Drop courtesy prefixes (Mr, Mrs, Ms, Miss, Mx); keep titles such as Sir, Dame and Dr.
        name = re.sub(r"^(Mr|Mrs|Ms|Miss|Mx)\.?\s+", "", v["nameDisplayAs"])
        mps[key(m["membershipFrom"])] = {"name": name, "party": v["latestParty"]["name"],
                                         "since": m["membershipStartDate"][:10], "id": v["id"]}
    skip += 20
    if skip >= page["totalResults"]:
        break


def build(year, cfg):
    """One year's tracker data from its England and Wales workbooks."""
    xlsx, wales_xlsx = os.path.join(DATA, cfg["xlsx"]), os.path.join(DATA, cfg["wales"])
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)

    def rows(sheet, start=1):
        return [list(r) for r in wb[sheet].iter_rows(min_row=start, values_only=True)]

    # ---------- Constituency Summary: the fiscal core ----------
    seats = {}
    for r in rows("Constituency Summary", 3):
        if not r[0] or r[1] not in REGIONS:
            continue
        k = key(r[0])
        adults = int(r[12])
        tax, welfare = r[3] * 1e9, r[5] * 1e9
        seats[k] = {
            "name": r[0].strip(), "region": r[1], "party": r[2].strip(),
            "tax": round(tax), "tpw": round(r[4]), "wel": round(welfare), "wpa": round(r[6]),
            "net": round(tax - welfare), "npa": round((tax - welfare) / adults), "ratio": round(welfare / tax, 4),
            "jobs": int(r[8]), "adults": adults,
            "relR": num(r[9], 4), "relE": num(r[10], 4),
            "ben": [round(v) for v in r[13:23]],
            "meanEst": bool(r[23]), "imputed": int(r[24] or 0), "flag": bool(r[25]), "country": "England",
        }
    assert len(seats) == 543, len(seats)

    # ---------- Health & APS: ONS code + indicators for every seat ----------
    for r in rows("Health & APS Data", 5):
        if not r[0] or not r[1]:
            continue
        s = seats[key(r[0])]
        s["code"] = r[1]
        s["h"] = {"le": num(r[2], 2), "obesity": num(r[3], 2), "diabetes": num(r[4], 2), "hyper": num(r[5], 2),
                  "depress": num(r[6], 2), "asthma": num(r[7], 2), "emp": num(r[8], 4), "inact": num(r[9], 4)}

    # ---------- Other Data: ASHE mean and median pay ----------
    for r in rows("Other Data", 3):
        if r[9] and key(r[9]) in seats:
            s = seats[key(r[9])]
            s["meanPay"] = round(r[14]) if r[14] else None
            s["medPay"] = round(r[16]) if isinstance(r[16], (int, float)) else None
            s["pop1864"] = int(r[13]) if r[13] else None

    # ---------- Regional sheets: pay ladder, job split, benefit claims ----------
    for region in REGIONS:
        data = rows(region)
        for i, r in enumerate(data):
            k = key(r[0]) if isinstance(r[0], str) else None
            if k not in seats or not isinstance(r[1], (int, float)):
                continue
            s = seats[k]
            s["pay"] = [round(v) if isinstance(v, (int, float)) else None for v in r[1:9]]
            for j in range(i + 1, min(i + 24, len(data))):
                lab = data[j][0]
                if lab == "Number of Jobs":
                    s["jobsSplit"] = [int(data[j][2] or 0), int(data[j][3] or 0)]
                elif isinstance(lab, str) and lab.strip() in ("Welfare Type", "Welfare"):
                    s["claims"] = [int(data[j + 1 + b][1] or 0) for b in range(10)]
                    break
    missing = [s["name"] for s in seats.values() if "pay" not in s or "code" not in s or "claims" not in s]
    assert not missing, missing

    # ---------- Regional pay ladders (Annual Pay Across Percentiles) ----------
    reg_pay = {}
    for r in rows("Annual Pay Across Percentiles", 4)[:18]:
        if isinstance(r[0], str):
            name = r[0].replace("Yorkshire & The Humber", "Yorkshire and The Humber")
            reg_pay[name] = [round(v) for v in (r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9])]

    # ---------- Wales ----------
    wwb = openpyxl.load_workbook(wales_xlsx, read_only=True, data_only=True)
    wales_seats, wales_extra = model.compute_wales(openpyxl.load_workbook(wales_xlsx, data_only=False))
    # once the workbook has been opened and saved in Excel it carries Excel's own results: compare them
    saved = {str(r[0]).strip(): r for r in wwb["Constituency Summary"].iter_rows(min_row=3, values_only=True) if r[0] and isinstance(r[3], (int, float))}
    if saved:
        gap = max(max(abs(w["tax"] / 1e9 - saved[w["name"]][3]), abs(w["wel"] / 1e9 - saved[w["name"]][5])) for w in wales_seats)
        print(f"{year} Wales: recomputed figures differ from Excel's saved results by at most £{gap * 1e9:,.0f} per seat.")
        assert gap < 1e-6, gap  # £1,000
    wales_method = [list(r[:2]) for r in wwb["Methodology & Scope"].iter_rows(min_row=3, values_only=True)]
    wales_tax_note = wwb["Tax Assumptions"]["A23"].value
    for w in wales_seats:
        tax, welfare, adults = w["tax"], w["wel"], w["adults"]
        seats[key(w["name"])] = {
            "name": w["name"], "region": "Wales", "party": w["party"].strip(), "code": w["code"], "country": "Wales",
            "tax": round(tax), "tpw": round(w["tpw"]), "wel": round(welfare), "wpa": round(welfare / adults),
            "net": round(tax - welfare), "npa": round((tax - welfare) / adults), "ratio": round(welfare / tax, 4),
            "jobs": int(round(w["jobs"])), "adults": int(adults), "relR": None, "relE": None,
            "ben": [round(v) for v in w["ben"]], "claims": [int(c) for c in w["claims"]], "h": None,
            "pay": [round(v) for v in w["pay"]], "meanPay": round(w["meanPay"]), "medPay": round(w["medPay"]) if isinstance(w["medPay"], (int, float)) else None,
            "jobsSplit": [int(round(v)) for v in w["jobsSplit"]], "meanEst": w["meanEst"], "imputed": w["imputed"],
            "flag": w["meanEst"] or w["imputed"] >= 2, "mpSheet": w["mpSheet"],
        }
    reg_pay["Wales"] = [round(v) for v in wales_extra["regionPay"]]
    assert len(seats) == 575, len(seats)

    # ---------- MPs ----------
    party_mismatch = []
    for k, s in seats.items():
        mp = mps.get(k)
        if mp:
            s["mp"], s["mpSince"], s["mpId"] = mp["name"], mp["since"], mp["id"]
            if mp["party"] != s["party"]:
                party_mismatch.append((s["name"], s["party"], mp["party"]))
        elif s.get("mpSheet"):
            s["mp"] = re.sub(r"^(Mr|Mrs|Ms|Miss|Mx)\.?\s+", "", s["mpSheet"])
        else:
            s["mp"] = None
    no_mp = [s["name"] for s in seats.values() if not s["mp"]]

    # ---------- Aggregates ----------
    def agg(group):
        tax = sum(s["tax"] for s in group)
        wel = sum(s["wel"] for s in group)
        adults = sum(s["adults"] for s in group)
        jobs = sum(s["jobs"] for s in group)
        return {"seats": len(group), "tax": tax, "wel": wel, "net": tax - wel, "adults": adults, "jobs": jobs,
                "ratio": round(wel / tax, 4), "tpw": round(tax / jobs), "wpa": round(wel / adults),
                "npa": round((tax - wel) / adults), "surplus": sum(1 for s in group if s["net"] > 0),
                "ben": [sum(s["ben"][b] for s in group) for b in range(10)]}

    all_seats = list(seats.values())
    england_seats = [s for s in all_seats if s["country"] == "England"]
    england = agg(england_seats)
    england["h"] = {}
    wales = agg([s for s in all_seats if s["country"] == "Wales"])
    regions = []
    for name in REGIONS + ["Wales"]:
        g = [s for s in all_seats if s["region"] == name]
        regions.append({"name": name, **agg(g), "pay": reg_pay.get(name)})
    party_order = ["Labour", "Labour (Co-op)", "Conservative", "Liberal Democrat", "Reform UK", "Plaid Cymru", "Green Party",
                   "Independent", "Your Party", "Restore Britain", "Speaker"]
    assert set(s["party"] for s in all_seats) <= set(party_order), set(s["party"] for s in all_seats)
    parties = []
    for name in party_order:
        g = [s for s in all_seats if s["party"] == name]
        if g:
            best = max(g, key=lambda s: s["net"])
            worst = min(g, key=lambda s: s["net"])
            parties.append({"name": name, **agg(g), "best": best["name"], "worst": worst["name"]})

    # Cross-check against the workbook's own England totals.
    tot = {r[0]: r[1] for r in rows("Totals", 38)[:6] if r[0]}
    assert abs(england["tax"] / 1e9 - tot["Total Employee Tax Contribution (£ bn)"]) < 0.01
    assert abs(england["wel"] / 1e9 - tot["Total Welfare Spend (£ bn)"]) < 0.01

    # ---------- England averages for the indicator strips ----------
    demo = rows("Demographic Data")
    avg_rows = {r[0]: r for r in demo[86:93] if r[0]}
    nat = avg_rows["National Average (England)"]
    england["h"] = {"le": num(nat[6], 2), "obesity": num(nat[7], 2), "diabetes": num(nat[8], 2),
                    "hyper": num(nat[9], 2), "depress": num(nat[10], 2), "asthma": num(nat[11], 2),
                    "emp": num(nat[13], 4), "inact": num(nat[14], 4)}
    england["medPay"] = round(sum(s["medPay"] * s["jobs"] for s in england_seats if s.get("medPay")) /
                              sum(s["jobs"] for s in england_seats if s.get("medPay")))
    wales["medPay"] = round(sum(s["medPay"] * s["jobs"] for s in all_seats if s["country"] == "Wales" and s.get("medPay")) /
                            sum(s["jobs"] for s in all_seats if s["country"] == "Wales" and s.get("medPay")))

    # ---------- Dumbbell: 10 highest- vs 10 lowest-welfare seats ----------
    def block(title):
        i = next(n for n, r in enumerate(demo) if r[2] == title)
        names = []
        for r in demo[i + 3:i + 14]:
            if not r[0] or r[0] == "Average":
                break
            names.append(r[0])
        return names

    dumb_cols = [("Who lives there", [(2, "Average age", "age"), (3, "Born abroad", "pct"),
                                     (4, "Aged under 18", "pct"), (5, "Aged 65+", "pct"),
                                     (16, "Social housing", "pct")]),
                 ("Work and skills", [(13, "Employment rate", "pct"), (15, "Median weekly pay", "gbp"),
                                      (17, "No qualifications", "pct"), (19, "Degree-level qualifications", "pct")]),
                 ("Health", [(6, "Life expectancy", "yrs"), (7, "Obesity", "pct1"), (8, "Diabetes", "pct1"),
                             (10, "Depression", "pct1"), (12, "Disabled", "pct")])]
    hi = avg_rows["Worst per Capita (England)"]
    lo = avg_rows["Best per Capita (England)"]
    dumbbell = {
        "groups": [{"name": g, "rows": [{"label": lab, "fmt": fmt, "hi": num(hi[c], 4), "lo": num(lo[c], 4),
                                         "eng": num(nat[c], 4)} for c, lab, fmt in cols]} for g, cols in dumb_cols],
        "hiSeats": block("WORST PER CAPITA (ENGLAND)"),
        "loSeats": block("BEST PER CAPITA (ENGLAND)"),
    }

    # ---------- Methodology text ----------
    method = []
    for r in rows("Methodology & Scope", 3):
        if r[0] and r[1]:
            method.append([r[0], r[1]])
        elif r[0] and not r[1]:
            method.append([r[0], None])
        elif r[1]:
            method.append([None, r[1]])
    wm = {r[0]: r[1] for r in wales_method if r[0] and r[1]}
    wales_sources = next((r[1] for i, r in enumerate(wales_method) if not r[0] and r[1] and i and
                          str(wales_method[i - 1][0]).startswith("5.")), None)
    method += [["6. Wales", None],
               ["Tax rates", wales_tax_note],
               ["Welfare data", f"Stat-Xplore, downloaded {WELFARE_PULLED}, using the same databases, measures and period as the England extract."],
               ["Not included", "As for England, plus Welsh Government schemes such as the Discretionary Assistance Fund and the "
                                "Council Tax Reduction Scheme."],
               ["Periods", wm.get("Periods")],
               ["Sparse percentiles", wm.get("Sparse percentiles")],
               ["Validation", wm.get("Validation")],
               ["Demographics", "The demographic, health, and work comparisons in the tracker cover England only."],
               ["Sources", wales_sources]]
    assert all(v for _, v in method[-8:]), [k for k, v in method[-8:] if not v]
    pay = "; ".join(f"{y} uses ONS {c['ashe']}, for {c['payDate']}" for y, c in sorted(YEARS.items(), reverse=True))
    method += [["7. Years", None],
               ["Pay", f"The tracker shows {len(YEARS)} years of pay data: {pay}."],
               ["Welfare", "Each year's welfare data and tax rules match the tax year its pay covers: " +
                           "; ".join(f"{y} uses Stat-Xplore {c['welDate']} for all ten benefits and {c['taxYear']} tax rules"
                                     for y, c in sorted(YEARS.items(), reverse=True)) +
                           ". The change between years therefore reflects pay, benefit claims and awards together."]]

    # ---------- This year's data ----------
    out_seats = []
    for s in sorted(all_seats, key=lambda s: s["name"]):
        out_seats.append({k: s.get(k) for k in (
            "code", "name", "region", "party", "mp", "mpSince", "tax", "tpw", "wel", "wpa", "net", "npa", "ratio",
            "jobs", "adults", "relR", "relE", "ben", "claims", "h", "pay", "meanPay", "medPay", "jobsSplit",
            "meanEst", "imputed", "flag", "country")})

    print(f"{year} ({cfg['ashe']} pay):")
    for label, a in (("England", england), ("Wales", wales)):
        print(f"  {label}: tax £{a['tax'] / 1e9:.2f}bn, welfare £{a['wel'] / 1e9:.2f}bn, "
              f"net £{a['net'] / 1e9:.2f}bn, {a['surplus']} of {a['seats']} seats in surplus")
    if no_mp:
        print("  No current MP found for:", no_mp)
    if party_mismatch:
        print("  Party differs from Parliament API (workbook, API):", party_mismatch)

    return {"year": year, "ashe": cfg["ashe"], "payDate": cfg["payDate"], "welDate": cfg["welDate"],
            "partyAsOf": PARTY_AS_OF, "mpAsOf": date.today().isoformat(), "benefits": BENEFITS, "pcts": PCTS,
            "england": england, "wales": wales, "regions": regions, "parties": parties, "seats": out_seats,
            "dumbbell": dumbbell, "method": method}


# ---------- Write data.js ----------
years = {y: build(y, cfg) for y, cfg in sorted(YEARS.items(), reverse=True)}
codes = {s["code"] for s in years[LATEST]["seats"]}
assert all({s["code"] for s in y["seats"]} == codes for y in years.values())
WT = {"latest": LATEST, "years": years}
with open(os.path.join(HERE, "data.js"), "w", encoding="utf-8") as f:
    f.write("window.WT=")
    json.dump(WT, f, separators=(",", ":"), ensure_ascii=False)
    f.write(";\n")

# ---------- Write geo.js ----------
K = math.cos(math.radians(52.7))


def ring(coords):
    out, last = [], None
    for lon, lat in coords:
        p = [round(lon * K, 4), round(lat, 4)]
        if p != last:
            out.append(p)
            last = p
    return out


geo = json.load(open(GEOJSON, encoding="utf-8"))
geo["features"] += json.load(open(WALES_GEOJSON, encoding="utf-8"))["features"]
feats = []
for f in geo["features"]:
    code = f["properties"]["PCON24CD"]
    if code not in codes:
        continue
    g = f["geometry"]
    if g["type"] == "Polygon":
        coords = [ring(r) for r in g["coordinates"]]
    else:
        coords = [[ring(r) for r in poly] for poly in g["coordinates"]]
    feats.append({"type": "Feature", "id": code, "geometry": {"type": g["type"], "coordinates": coords}})
assert len(feats) == 575, len(feats)
with open(os.path.join(HERE, "geo.js"), "w", encoding="utf-8") as f:
    f.write("window.GEO=")
    json.dump({"type": "FeatureCollection", "features": feats}, f, separators=(",", ":"))
    f.write(";\n")

print(f"{len(codes)} seats x {len(years)} years -> data.js ({os.path.getsize(os.path.join(HERE, 'data.js')) / 1e3:.0f} KB); "
      f"geo.js ({os.path.getsize(os.path.join(HERE, 'geo.js')) / 1e3:.0f} KB)")

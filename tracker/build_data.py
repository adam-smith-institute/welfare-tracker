"""Build the tracker's data files from the welfare workbook.

    pip install openpyxl
    python3 build_data.py

Reads  ../A Tale of Two Cities (Welfare Analysis) 2026.xlsx
       ../Westminster_Parliamentary_Constituencies_July_2024_Boundaries_UK_BUC_*.geojson
       UK Parliament Members API (current MPs, fetched live)
Writes data.js   window.WT  = {england, regions, parties, seats, dumbbell, method, ...}
       geo.js    window.GEO = England constituency boundaries, pre-projected to a flat
                             plane (x = lon * cos 52.7deg, y = lat) so the page can draw
                             them with d3.geoIdentity and no winding-order issues.

Re-run whenever the workbook changes. The page reads only these two files.
"""
import glob, json, math, os, re, unicodedata, urllib.request
from datetime import date

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
XLSX = os.path.join(ROOT, "A Tale of Two Cities (Welfare Analysis) 2026.xlsx")
GEOJSON = glob.glob(os.path.join(ROOT, "Westminster_Parliamentary_Constituencies_July_2024_*.geojson"))[0]
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


wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)


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
        "meanEst": bool(r[23]), "imputed": int(r[24] or 0), "flag": bool(r[25]),
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

# ---------- MPs: UK Parliament Members API ----------
mps = {}
skip = 0
while True:
    with urllib.request.urlopen(MEMBERS_URL.format(skip), timeout=30) as resp:
        page = json.load(resp)
    for it in page["items"]:
        v = it["value"]
        m = v["latestHouseMembership"]
        mps[key(m["membershipFrom"])] = {"name": v["nameDisplayAs"], "party": v["latestParty"]["name"],
                                         "since": m["membershipStartDate"][:10], "id": v["id"]}
    skip += 20
    if skip >= page["totalResults"]:
        break
party_mismatch = []
for k, s in seats.items():
    mp = mps.get(k)
    if mp:
        s["mp"], s["mpSince"], s["mpId"] = mp["name"], mp["since"], mp["id"]
        if mp["party"] != s["party"]:
            party_mismatch.append((s["name"], s["party"], mp["party"]))
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
england = agg(all_seats)
england["h"] = {}
regions = []
for name in REGIONS:
    g = [s for s in all_seats if s["region"] == name]
    regions.append({"name": name, **agg(g), "pay": reg_pay.get(name)})
party_order = ["Labour", "Labour (Co-op)", "Conservative", "Liberal Democrat", "Reform UK", "Green Party",
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
england["medPay"] = round(sum(s["medPay"] * s["jobs"] for s in all_seats if s.get("medPay")) /
                          sum(s["jobs"] for s in all_seats if s.get("medPay")))

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

# ---------- Write data.js ----------
out_seats = []
for s in sorted(all_seats, key=lambda s: s["name"]):
    out_seats.append({k: s.get(k) for k in (
        "code", "name", "region", "party", "mp", "mpSince", "tax", "tpw", "wel", "wpa", "net", "npa", "ratio",
        "jobs", "adults", "relR", "relE", "ben", "claims", "h", "pay", "meanPay", "medPay", "jobsSplit",
        "meanEst", "imputed", "flag")})

WT = {"edition": 2026, "partyAsOf": "23 September 2026", "mpAsOf": date.today().isoformat(), "benefits": BENEFITS, "pcts": PCTS,
      "england": england, "regions": regions, "parties": parties, "seats": out_seats,
      "dumbbell": dumbbell, "method": method}
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
codes = {s["code"] for s in all_seats}
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
assert len(feats) == 543, len(feats)
with open(os.path.join(HERE, "geo.js"), "w", encoding="utf-8") as f:
    f.write("window.GEO=")
    json.dump({"type": "FeatureCollection", "features": feats}, f, separators=(",", ":"))
    f.write(";\n")

print(f"{len(out_seats)} seats -> data.js ({os.path.getsize(os.path.join(HERE, 'data.js')) / 1e3:.0f} KB); "
      f"geo.js ({os.path.getsize(os.path.join(HERE, 'geo.js')) / 1e3:.0f} KB)")
print(f"England: tax £{england['tax'] / 1e9:.1f}bn, welfare £{england['wel'] / 1e9:.1f}bn, "
      f"net £{england['net'] / 1e9:.1f}bn, {england['surplus']} seats in surplus")
if no_mp:
    print("No current MP found for:", no_mp)
if party_mismatch:
    print("Party differs from Parliament API (workbook, API):", party_mismatch)

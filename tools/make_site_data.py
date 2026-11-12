#!/usr/bin/env python3
"""data.json -> assets/data.js (sayt üçün yığcam data). İstifadə: python3 tools/make_site_data.py data.json assets/data.js"""
import json, sys, statistics
d = json.load(open(sys.argv[1]))
MONTHS = [5, 6, 7, 8, 9, 10]
inv = {}
for k, v in d["invoices"].items():
    inv[k] = {}
    for ab, i in v.items():
        ats = dict(i["ats"])
        gap = round(i["kwh"] - sum(ats.values()), 2)
        if abs(gap) >= 1:
            ats["digər"] = round(ats.get("digər", 0) + gap, 2)
        inv[k][ab] = dict(rem=i.get("removed_kwh", 0), period=i.get("period"), kwh=i["kwh"], net=i["net"], vat=i["vat"], total=i["total"], ats=ats, file=i["file"])

# Excel: xam sayğac fərqləri; səhv göstəricili aylar null olur
bad = set()
for f in d["flags"]:
    bad.add((f["serial"], f["month"]))
    if "üst-üstə düşmür" in f["why"]:
        y, m = map(int, f["month"].split("-"))
        bad.add((f["serial"], f"{y}-{m-1:02d}"))
meters = {}
for k in sorted(d["excel"]):
    for r in d["excel"][k]:
        m = meters.setdefault(r["serial"], dict(name=r["name"], d={}))
        m["d"][k] = None if (r["serial"], k) in bad else round(r["cur"] - r["prev"], 3)
        m["name"] = r["name"]
out = dict(
    inv=inv,
    meters=[dict(name=m["name"], serial=s,
                 y25=[m["d"].get(f"2025-{x:02d}") for x in MONTHS],
                 y26=[m["d"].get(f"2026-{x:02d}") for x in MONTHS]) for s, m in meters.items()],
    flags=d["flags"],
    removed=d.get("removed_names", []),
)
open(sys.argv[2], "w").write("window.PANEL_DATA=" + json.dumps(out, ensure_ascii=False, separators=(",", ":")) + ";\n")
print("ok", len(open(sys.argv[2]).read()) // 1024, "KB")

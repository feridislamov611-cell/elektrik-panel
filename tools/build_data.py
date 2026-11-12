#!/usr/bin/env python3
"""Excel (EATS göstəriciləri) və hesab-fakturalardan (PDF) data.json yaradır.
İstifadə: python3 tools/build_data.py <açılmış_qovluq> <rtf_pdf_qovluğu> <çıxış.json>
"""
import sys, re, glob, json, os, subprocess, datetime
import openpyxl, statistics
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from parse_rows import rows_of, GROUPS

SRC, PDF25, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
MONTHS = {"may":5,"iyun":6,"iyul":7,"avqust":8,"sentyabr":9,"oktyabr":10}
NUM = r"-?[\d.,]+"

def num(s, comma_dec=None):
    s = s.strip()
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    return float(s)

def pdftext(p):
    return subprocess.run(["pdftotext","-layout",p,"-"],capture_output=True,text=True).stdout

def folder_ym(d):
    y, m = os.path.basename(d).split()
    return int(y), MONTHS[m.lower()]

# ---------- Excel ----------
excel = {}
for d in sorted(glob.glob(os.path.join(SRC, "202* *"))):
    y, m = folder_ym(d)
    f = glob.glob(os.path.join(d, "*.xlsx"))[0]
    ws = openpyxl.load_workbook(f, data_only=True).worksheets[0]
    rows = []
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if r[0] is None or r[3] is None: continue
        name = str(r[1]).strip()
        rows.append(dict(sn=int(r[0]), name=name, serial=str(r[2]).strip(),
                         prev=float(r[3]), cur=float(r[4]),
                         rprev=(float(r[6]) if r[6] is not None else None),
                         rcur=(float(r[7]) if r[7] is not None else None)))
    excel[f"{y}-{m:02d}"] = rows

# ---------- Fakturalar ----------
def parse_invoice(text):
    """Yekun: kWh, ₼ (ƏDV-siz), ƏDV, ümumi ₼. Əmsallar: serial -> əmsal."""
    t = text
    res = {}
    # Ümumi ödəniləcək məbləğ ("Cəmi AZN ...") və ƏDV
    m = re.findall(r"C\S*mi\s+AZN\s+(" + NUM + ")", t)
    res["total"] = num(m[-1]) if m else None
    m = re.findall(r"C\S*mi\s+\S*DV\s+(" + NUM + ")", t)
    res["vat"] = num(m[-1]) if m else None
    m = re.findall(r"C\S*mi\s+M\S*bl\S*q\s+(" + NUM + ")", t)
    res["net"] = num(m[-1]) if m else None
    # Ümumi kWh: "Cəmi  <G.Sərf> <Əlavə> <TFərq> <Sərfiyyat> <Məbləğ> ..." sonuncu belə sətir
    kwh = None
    for line in t.splitlines():
        mm = re.match(r"\s*C\S*mi\s+(" + NUM + r")\s+(" + NUM + r")\s+(" + NUM + r")\s+(" + NUM + r")\s+(" + NUM + ")", line)
        if mm and "Aktiv" not in line:
            kwh = num(mm.group(4)); net_row = num(mm.group(5))
    res["kwh"] = kwh
    res["net_row"] = net_row if kwh is not None else None
    # Əmsallar
    coef = {}
    flat = re.sub(r"[ \t]+", " ", t)
    for mm in re.finditer(r"(\S+) 0,11 (\d+) ", flat):
        ser = re.split("M", mm.group(1))[0]
        coef[ser] = int(mm.group(2))
    for mm in re.finditer(r"(\S+) (\d+) \d\d\.\d\d\.\d{4} \d\d\.\d\d\.\d{4}", flat):
        ser = re.split("M", mm.group(1))[0]
        coef.setdefault(ser, int(mm.group(2)))
    res["coef"] = coef
    return res


# Excel: sayğac nömrəsi -> ATS qrupu
from parse_rows import ats_of
ser2ats = {}
for rows in excel.values():
    for r in rows:
        g = ats_of(r["name"].replace("-", " Ats-", 1)) if re.match(r"\d{3}", r["name"]) else None
        m = re.match(r"(\d{3}(?:/\d{3})?)", r["name"])
        if m: ser2ats[r["serial"]] = m.group(1)

RATE_NET = None
invoices = {}
srcs = []
for p in glob.glob(os.path.join(SRC, "2026 */*.pdf")): srcs.append((p, pdftext(p)))
for p in glob.glob(os.path.join(PDF25, "2025 */*.pdf")): srcs.append((p, pdftext(p)))
for p, text in sorted(srcs):
    y, m = folder_ym(os.path.dirname(p))
    b = os.path.basename(p)
    ab = "100-107" if "107" in b else "200-068"
    key = f"{y}-{m:02d}"
    inv = parse_invoice(text)
    rows = rows_of(text)
    ats = {}
    other = 0.0
    for r in rows:
        g = r["ats"] or ser2ats.get(r["serial"])
        if ab == "100-107": g = "107"
        if g is None: g = "digər"
        if r["kwh"] is None: continue
        ats[g] = ats.get(g, 0) + r["kwh"]
    inv["ats"] = {k: round(v, 2) for k, v in ats.items()}
    inv["rows_sum"] = round(sum(ats.values()), 2)
    inv["rows_ok"] = abs(inv["rows_sum"] - (inv["kwh"] or 0)) < 1.5
    inv["file"] = os.path.relpath(p, SRC) if p.startswith(SRC) else os.path.relpath(p, PDF25)
    inv.pop("coef"); inv.pop("net_row")
    inv["vat"] = round(inv["net"] * 0.18, 2)
    inv["total"] = round(inv["net"] + inv["vat"], 2)
    invoices.setdefault(key, {})
    invoices[key].setdefault(ab, inv)   # May 2025 200-068 iki nüsxədir

# ---------- Excel yoxlaması ----------
flags = []
by_serial = {}
for k in sorted(excel):
    for r in excel[k]:
        by_serial.setdefault(r["serial"], []).append((k, r))
for ser, lst in by_serial.items():
    diffs = [r["cur"] - r["prev"] for _, r in lst]
    med = statistics.median(diffs)
    for k, r in lst:
        d = r["cur"] - r["prev"]
        why = None
        if d < 0: why = "göstərici əvvəlkindən kiçikdir (mənfi fərq)"
        elif med > 0 and d > 6 * med and d > 5000: why = f"fərq median ({med:,.0f}) ilə müqayisədə {d/med:.0f} dəfə çoxdur — yazı səhvi ehtimalı"
        if why: flags.append(dict(month=k, name=r["name"], serial=ser, prev=r["prev"], cur=r["cur"], diff=round(d,2), why=why))
# davamlılıq: bu ayın "əvvəlki" göstəricisi keçən ayın "son" göstəricisi ilə uyğun olmalıdır
for k in sorted(excel):
    y, m = map(int, k.split("-"))
    pk = f"{y}-{m-1:02d}"
    if pk not in excel: continue
    last = {r["serial"]: r for r in excel[pk]}
    for r in excel[k]:
        p = last.get(r["serial"])
        if p and abs(p["cur"] - r["prev"]) > 1:
            flags.append(dict(month=k, name=r["name"], serial=r["serial"], prev=r["prev"], cur=p["cur"], diff=round(r["prev"]-p["cur"],2),
                              why=f"başlanğıc göstərici ({r['prev']:,.0f}) keçən ayın son göstəricisi ({p['cur']:,.0f}) ilə üst-üstə düşmür"))

json.dump(dict(excel=excel, invoices=invoices, flags=flags), open(OUT, "w"), ensure_ascii=False, indent=1)
print("ok", OUT, "| faktura:", sum(len(v) for v in invoices.values()), "| bayraq:", len(flags))

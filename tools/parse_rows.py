import re
N = r"-?(?:\d[\d.,]*|\.\d+)"
PAIR = {"430":"430/431","431":"430/431","432":"432/434","434":"432/434","538":"538/539","539":"538/539",
        "568":"568/569","569":"568/569","594":"594/595","595":"594/595"}
GROUPS = ["401","406","430/431","432/434","491","502","538/539","568/569","581","594/595","91"]

def val(s):
    s = s.strip()
    if "," in s: s = s.replace(".", "").replace(",", ".")
    if s.startswith("."): s = "0" + s
    if s.startswith("-."): s = "-0" + s[1:]
    return float(s)

def ats_of(name):
    """Uçot nöqtəsinin adından ATS qrupunu tap."""
    n = re.sub(r"^\d{10,15}\s*", "", name.strip())
    m = re.search(r"A432-434", n)
    if m: return "432/434"
    m = re.search(r"(?:ATS|Ats|ATC|АТС|ÀÒÑ)\s*-?\s*(\d{2,3})\b", n)
    if m:
        g = m.group(1)
        return PAIR.get(g, g)
    m = re.match(r"(\d{3})\s", n)           # "406 ÀÒÑ"
    if m: return PAIR.get(m.group(1), m.group(1))
    m = re.search(r"\b(\d{3})\b", n)
    if m and PAIR.get(m.group(1), m.group(1)) in GROUPS: return PAIR.get(m.group(1), m.group(1))
    return None

def rows_of(text):
    lines = [re.sub(r"[ \t]+", " ", l.strip()) for l in text.splitlines()]
    out = []
    for i, l in enumerate(lines):
        m = re.search(r"^(.*?)(\S*) 0,11 (\d+) (" + N + r") (" + N + r")(.*)$", l)
        if not m: continue
        head = m.group(1)
        ser = re.split("M", m.group(2))[0] if m.group(2) else ""
        rest = m.group(6)
        nxt = lines[i + 1] if i + 1 < len(lines) else ""
        kwh = None
        for cand in (rest, rest + " " + nxt, nxt):
            cand = re.sub(r"C\S*mi", " ", cand).strip()
            mm = re.match(r"(" + N + r") (" + N + r") (" + N + r") (" + N + r")\b", cand)
            if mm:
                kwh = val(mm.group(4)); break
        out.append(dict(serial=ser, coef=int(m.group(3)), kwh=kwh, name=head, ats=ats_of(head)))
    return out

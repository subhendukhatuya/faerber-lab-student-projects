"""Regenerate projects.json from Student_Projects.xlsx.
Usage: python build.py   (needs: pip install openpyxl)
Columns: Supervisor | Topic Title | Short Description | Expectations/ Deliverables |
         Prerequisites | Max. No of Students | Availability | Note
Text after a line 'References' in the description is shown as references.
"""
import json, re, openpyxl, pathlib

HERE = pathlib.Path(__file__).parent
# Content for rows whose sheet cells say "same content as the old portal" or are empty.
FALLBACK = json.loads((HERE / "fallback_content.json").read_text())

def bullets(text):
    if not text: return []
    t = text.strip()
    if t.startswith("(1)"):
        parts = re.split(r"\(\d\)\s*", t)
    else:
        parts = re.split(r"\n\s*[-•]\s*", "\n" + t)
        if len(parts) <= 2 and ";" in t: parts = t.split(";")
    return [p.strip(" \n-") for p in parts if p.strip(" \n-")]

def slug(s): return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")[:60]

wb = openpyxl.load_workbook(HERE / "Student_Projects.xlsx")
ws = wb.active
rows = list(ws.iter_rows(values_only=True))
out = []
for r in rows[1:]:
    if not r[0] or not r[1]: continue
    sup, title, desc, exp, pre, maxn, avail = [ (c.strip() if isinstance(c,str) else c) for c in r[:7] ]
    refs = []
    if desc:
        m = re.split(r"\n\s*References\s*\n", desc, maxsplit=1)
        desc = m[0].strip()
        if len(m) > 1: refs = bullets(m[1]) if "\n- " in "\n"+m[1] else [x.strip() for x in m[1].split("\n") if x.strip()]
    fb = FALLBACK.get(title, {})
    if not desc or desc.lower().startswith("same content"): desc = fb.get("description", "")
    p = {
        "id": slug(title), "title": title, "supervisor": sup,
        "description": [x.strip() for x in desc.split("\n") if x.strip()],
        "deliverables": bullets(exp) or fb.get("deliverables", []),
        "prerequisites": bullets(pre) or fb.get("prerequisites", []),
        "references": refs or fb.get("references", []),
        "places": int(maxn or 1),
        "status": (avail or "Open").strip().title(),
    }
    out.append(p)
(HERE / "projects.json").write_text(json.dumps(out, indent=2, ensure_ascii=False))
print(f"{len(out)} projects ->", [ (p['title'][:30], p['status']) for p in out])

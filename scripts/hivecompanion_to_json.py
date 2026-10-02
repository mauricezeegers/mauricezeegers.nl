#!/usr/bin/env python3
"""Turn a HiveCompanion export into the bees.log object of the site-data block.

Usage:
  python3 hivecompanion_to_json.py EXPORT [--asof YYYY-MM-DD] > log.json

EXPORT is either the .xlsx file itself or a Gmail message in RAW format saved as JSON
(the {"raw": "..."} result of get_message); the .xlsx attachment is then taken from it.
Python standard library only.

Cleaning rules (agreed with Maurice):
  * exact duplicate rows (same date, hive and amount/product) are double entries and count once;
  * all oxalic acid spellings (Ox vap, Ox var, Ox vapor, Oxalic acid) become "Oxalic acid";
  * the Apiaries sheet (GPS coordinates) and the Storage apiary are never exported;
  * behaviour score: 1 = calm, 5 = aggressive.
"""
import base64, email, io, json, re, sys, zipfile, datetime
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


def load_xlsx_bytes(path):
    if path.lower().endswith(".xlsx"):
        return open(path, "rb").read()
    raw = json.load(open(path))["raw"]
    raw += "=" * (-len(raw) % 4)
    msg = email.message_from_bytes(base64.urlsafe_b64decode(raw))
    for part in msg.walk():
        name = part.get_filename() or ""
        if name.lower().endswith(".xlsx"):
            return part.get_payload(decode=True)
    sys.exit("no .xlsx attachment found in " + path)


def col_index(ref):
    letters = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_sheets(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS):
            shared.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    sheets = {}
    for s in ET.fromstring(z.read("xl/workbook.xml")).find("m:sheets", NS):
        target = rels[s.get("{%s}id" % NS["r"])].lstrip("/")
        target = target if target.startswith("xl/") else "xl/" + target
        rows = []
        for row in ET.fromstring(z.read(target)).iter("{%s}row" % NS["m"]):
            vals = {}
            for c in row.findall("m:c", NS):
                t, v = c.get("t"), c.find("m:v", NS)
                if t == "inlineStr":
                    val = "".join(x.text or "" for x in c.iter("{%s}t" % NS["m"]))
                elif v is None:
                    val = ""
                elif t == "s":
                    val = shared[int(v.text)]
                else:
                    val = v.text or ""
                vals[col_index(c.get("r"))] = val
            rows.append([vals.get(i, "") for i in range(max(vals) + 1)] if vals else [])
        if rows:
            head = rows[0]
            sheets[s.get("name")] = [dict(zip(head, r + [""] * (len(head) - len(r)))) for r in rows[1:] if any(r)]
        else:
            sheets[s.get("name")] = []
    return sheets


def day(v):
    """YYYY-MM-DD from a text date or an Excel serial number."""
    v = (v or "").strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}", v):
        return v[:10]
    try:
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=float(v))).isoformat()
    except ValueError:
        return ""


def num(v):
    try:
        return round(float(v), 2)
    except (TypeError, ValueError):
        return 0


def dedupe(rows):
    seen, out = set(), []
    for r in rows:
        k = json.dumps(r)
        if k not in seen:
            seen.add(k)
            out.append(r)
    return out


def main():
    args = sys.argv[1:]
    if not args:
        sys.exit(__doc__)
    asof = datetime.date.today().isoformat()
    if "--asof" in args:
        asof = args[args.index("--asof") + 1]
    S = read_sheets(load_xlsx_bytes(args[0]))

    hives = []
    for h in S.get("Hives", []):
        if h.get("Apiary") in ("", "Storage"):
            continue
        hives.append({
            "name": h["Hive"], "apiary": h["Apiary"], "status": h.get("Status", ""),
            "queen": h.get("Has queen") == "Yes", "queenColor": (h.get("Queen color") or "").lower() or None,
            "bodies": int(num(h.get("Bodies"))), "supers": int(num(h.get("Supers"))),
        })
    names = {h["name"] for h in hives}

    def keep(rows):
        return [r for r in rows if r[1] in names and r[0]]

    insp = keep([[day(r["Date"]), r["Hive"], int(num(r.get("Population (1-5)"))), int(num(r.get("Food stores (1-5)"))),
                  int(num(r.get("Behavior (1-5)"))), (r.get("Notes") or "").strip()] for r in S.get("Inspections", [])])
    treat = []
    for r in S.get("Treatments", []):
        p = (r.get("Product used") or "").strip()
        if re.match(r"(?i)^ox", p):
            p = "Oxalic acid"
        treat.append([day(r["Date"]), r["Hive"], p])
    feed = keep([[day(r["Date"]), r["Hive"], num(r.get("Syrup")) + num(r.get("Honey")) + num(r.get("Fondant"))] for r in S.get("Feedings", [])])
    harv = keep([[day(r["Date"]), r["Hive"], num(r.get("Honey quantity"))] for r in S.get("Harvests", [])])
    splits = [[day(r["Date"]), r["From hive"], r["Destination hive"]] for r in S.get("Splits", []) if r.get("Date")]
    # open tasks, only those tied to a hive: by the Hive column or by a hive name in the title
    # ("59" also means Fifty nine). General tasks stay off the site.
    tasks = []
    for r in S.get("Tasks", []):
        title = (r.get("Title") or "").strip()
        if r.get("Completed") != "No" or not title:
            continue
        for n in sorted(names):
            short = re.sub(r"\s*\(\d+\)$", "", n).lower()
            alias = [short] + (["59"] if short == "fifty nine" else [])
            if r.get("Hive") == n or any(re.search(r"\b" + re.escape(a) + r"\b", title.lower()) for a in alias):
                tasks.append([title, n])

    log = {
        "asOf": asof, "source": "HiveCompanion",
        "hives": hives,
        "inspections": sorted(dedupe(insp)),
        "treatments": sorted(dedupe(keep(treat))),
        "feedings": sorted(dedupe(feed)),
        "harvests": sorted(dedupe(harv)),
        "splits": sorted(dedupe(splits)),
        "tasks": tasks,
    }
    json.dump(log, sys.stdout, ensure_ascii=False, indent=1)
    sys.stdout.write("\n")


if __name__ == "__main__":
    main()

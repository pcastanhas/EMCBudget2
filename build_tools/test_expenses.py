# EXP_ sheets: data tie-out, input flow to DETAIL / QTRLY, section order, duplicates across sections, refresh checks.
import sys, pandas as pd
from lo_driver import LO
from make_xlsm import embed_vba, cell
BLDG = sys.argv[2] if len(sys.argv) > 2 else "42"
from paths import MRI_DIR
SRC = MRI_DIR.rstrip("/") + "/"
MON = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]
v = pd.read_excel(SRC + "Q1-Q3_2026_EXPENSES_BY_VENDOR.xlsx", sheet_name="MRIQRY", dtype=str)
for c in ("entityid","vendid","acctnum"): v[c] = v[c].str.strip()
for m in MON: v[m] = pd.to_numeric(v[m])
v = v[v.entityid == BLDG]
lo = LO()
try:
    doc = lo.load(sys.argv[1]); embed_vba(lo, doc)
    run = lambda: (lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet"), doc.calculateAll())
    S = lambda: cell(doc, "DETAIL", "B3").String
    sh = lambda n: doc.Sheets.getByName(n)
    run(); print("0) built:", S(), "| QTRLY:", cell(doc, "QTRLY SUM", "B3").String)
    det = lambda: sh("DETAIL").getCellRangeByPosition(0, 0, 20, 120).getDataArray()
    print("   DETAIL section headings before inputs:", [r[0] for r in det() if isinstance(r[0], str) and r[0].isupper() and len(r[0]) > 6][:8])
    def block(sheet, acct):
        arr = sh(sheet).getCellRangeByPosition(0, 0, 20, 1500).getDataArray()
        h = [i for i, r in enumerate(arr) if r[0] == acct][0]
        rows = []
        for j in range(h + 1, len(arr)):
            if arr[j][0] != "" and isinstance(arr[j][0], float): break
            rows.append((j + 1, arr[j]))
            if "total" in str(arr[j][2]): break
        return h + 1, rows
    # 1) tie-out of vendor rows and account rows to the raw extract and GL data
    bad = 0; n_v = 0
    for _, _, name in [(4,"","EXP_UTILITIES"),(5,"","EXP_JANITORIAL"),(6,"","EXP_ELEVATOR"),(7,"","EXP_HVAC"),(8,"","EXP_LANDSCAPING"),(9,"","EXP_GENERAL_RM")]:
        arr = sh(name).getCellRangeByPosition(0, 0, 20, 1500).getDataArray()
        cur = None
        for r in arr[8:]:
            if isinstance(r[0], float) and r[0] > 1e6: cur = int(r[0])
            if r[4] == "VEND":
                x = v[(v.vendid == r[3]) & (v.acctnum == f"MR{cur}")]
                exp = [round(float(x[m].sum()), 2) for m in MON]
                got = [round(float(c), 2) for c in r[6:18]]
                n_v += 1; bad += exp != got
    print(f"1) vendor lines checked against the raw extract: {n_v} | mismatches: {bad}")
    # 2) enter budget amounts
    def first_input(sheet, acct):
        h, rows = block(sheet, acct)
        return [rn for rn, r in rows if r[4] == "INPUT"][0]
    r1 = first_input("EXP_UTILITIES", 5110000); r2 = first_input("EXP_ELEVATOR", 5265000); r3 = first_input("EXP_GENERAL_RM", 5410000); r4 = first_input("EXP_JANITORIAL", 5225000)
    for sheet, rr, val in (("EXP_UTILITIES", r1, 25000), ("EXP_ELEVATOR", r2, 700), ("EXP_GENERAL_RM", r3, 2500), ("EXP_JANITORIAL", r4, 9000)):
        for c in "GHIJKLMNOPQR": sh(sheet).getCellRangeByName(f"{c}{rr}").setValue(val)
    sh("EXP_GENERAL_RM").getCellRangeByName(f"C{r3}").setString("test")
    doc.calculateAll()
    print("2) amounts typed (Utilities 25,000/mo, Elevator 700/mo, Gen R&M 2,500/mo, Cleaning-Contract 9,000/mo) ->", S())
    run()
    d = det()
    heads = [(i + 1, r[0]) for i, r in enumerate(d) if isinstance(r[0], str) and r[0].isupper() and len(r[0]) > 6 and not r[0].startswith("TOTAL")]
    print("   after rebuild:", S(), "| QTRLY:", cell(doc, "QTRLY SUM", "B3").String)
    print("   section order:", [h for _, h in heads])
    tot = {r[1]: round(r[15]) for r in d if isinstance(r[1], str) and r[1].startswith("TOTAL")}
    print("   section totals:", tot, "(expect UTILITIES 300,000 | ELEVATOR 8,400 | GENERAL R&M 30,000 | JANITORIAL 108,000)")
    rows5225 = [(i + 1, r[15]) for i, r in enumerate(d) if r[0] == 5225000]
    print("   account 5225000 rows:", len(rows5225), "totals", [round(x) for _, x in rows5225], "(expect one row: 108,000 spend less 10,001 tenant-billed offset = 97,999)")
    q = sh("QTRLY SUM").getCellRangeByPosition(0, 0, 8, 120).getDataArray()
    print("   QTRLY 5110000 total:", [round(r[7]) for r in q if r[0] == 5110000], "| QTRLY status OK:", cell(doc, "QTRLY SUM", "B3").String)
    # 3) refresh checks: new vendor on an existing block, and a new qualifying account without a block
    dv = sh("DATA_VENDOR"); n = int(sum(1 for r in dv.getCellRangeByPosition(0, 1, 3, 400).getDataArray() if r[0] != "")) + 2
    for c, val in enumerate(["42", "NEWVND", "NEW VENDOR LLC", "MR5110000"]): dv.getCellByPosition(c, n - 1).setString(val)
    dv.getCellByPosition(4, n - 1).setValue(100); doc.calculateAll()
    print("3) new vendor on an existing account ->", S())
    ea = sh("EXPENSE_ACCOUNTS").getCellRangeByPosition(0, 1, 10, 1600).getDataArray()
    free = [int(r[0]) for r in ea if r[4] == 1.0 and r[8] == 0.0][0]       # in an expense section, no activity -> no block
    dv.getCellByPosition(3, n - 1).setString(f"MR{free}"); doc.calculateAll()
    ea = sh("EXPENSE_ACCOUNTS").getCellRangeByPosition(0, 1, 10, 1600).getDataArray()
    nob = [int(r[0]) for r in ea if r[10] == 1.0]
    print(f"   vendor activity on account {free} (in {[r[2] for r in ea if r[0]==free][0]}, no block) ->", S(), "| NO BLOCK accounts:", nob)
    doc.close(True)
finally:
    lo.close()

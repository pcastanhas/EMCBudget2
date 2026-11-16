# INSURANCE sheet: calculated row vs independent calc, increase %, Use calc switch, adjustments, flow to DETAIL/QTRLY, new-vendor flag.
import sys, pandas as pd
from lo_driver import LO
from make_xlsm import embed_vba, cell
BLDG = sys.argv[2] if len(sys.argv) > 2 else "42"
SRC = "/mnt/project/"
MON = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]
def load(f):
    d = pd.read_excel(SRC + f, sheet_name="MRIQRY", dtype=str); d["entityid"] = d.entityid.str.strip(); d["acctnum"] = d.acctnum.str.strip()
    for m in MON: d[m] = pd.to_numeric(d[m])
    return d[d.entityid == BLDG]
a26 = load("Q1-Q3_2026_ACTUALS.xlsx"); b26 = load("FY_2026_BUDGETS.xlsx")
def fcst(acct, thru=9):
    a = a26[a26.acctnum == f"MR{acct}"][MON].sum().tolist(); b = b26[b26.acctnum == f"MR{acct}"][MON].sum().tolist()
    return [a[i] if i < thru else b[i] for i in range(12)]
lo = LO()
try:
    doc = lo.load(sys.argv[1]); embed_vba(lo, doc)
    run = lambda: (lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet"), doc.calculateAll())
    S = lambda: cell(doc, "DETAIL", "B3").String
    sh = doc.Sheets.getByName("INSURANCE"); T = lambda ref: sh.getCellRangeByName(ref)
    run(); print("0) built:", S(), "| QTRLY:", cell(doc, "QTRLY SUM", "B3").String)
    arr = sh.getCellRangeByPosition(0, 0, 20, 60).getDataArray()
    blk = {}; cur = None
    for i, r in enumerate(arr):
        if isinstance(r[0], float) and r[0] > 1e6: cur = int(r[0]); blk[cur] = {"r0": i + 1}
        if cur and r[4] in ("CALC", "FINAL"): blk[cur][r[4]] = i + 1
        if cur and r[4] == "INPUT": blk[cur].setdefault("INPUT", []).append(i + 1)
    def mon(acct, tag): return [round(c, 2) for c in sh.getCellRangeByPosition(6, blk[acct][tag] - 1, 17, blk[acct][tag] - 1).getDataArray()[0]]
    for acct, pct in ((5750000, 0.0215), (5755000, 0.01)):
        exp = [round(x * (1 + pct), 2) for x in fcst(acct)]
        print(f"1) {acct} defaults (+{pct:.2%}): calc row match: {mon(acct,'CALC') == exp} | total {sum(mon(acct,'FINAL')):,.2f} (expect {sum(exp):,.2f}) | header % {T('H'+str(blk[acct]['r0'])).Value} use-calc {T('J'+str(blk[acct]['r0'])).String}")
    d = lambda: doc.Sheets.getByName("DETAIL").getCellRangeByPosition(0, 0, 20, 140).getDataArray()
    tot = [round(r[15], 2) for r in d() if r[1] == "TOTAL INSURANCE"]
    print("   DETAIL TOTAL INSURANCE:", tot, "| expect", round(sum(sum(fcst(5750000)) * 1.0215 for _ in [0]) + sum(fcst(5755000)) * 1.01, 2))
    # 2) increase % and Use calc
    r0 = blk[5750000]["r0"]
    T(f"H{r0}").setValue(0.05); doc.calculateAll(); print("2) 5750000 at 5%: total", round(sum(mon(5750000, 'FINAL')), 2), "(expect", round(sum(fcst(5750000)) * 1.05, 2), ")")
    T(f"J{r0}").setString("No"); doc.calculateAll(); print("   Use calc = No: calc row total", round(sum(mon(5750000, 'CALC')), 2), "| final", round(sum(mon(5750000, 'FINAL')), 2))
    # type a renewal premium on an open line (last INPUT row of the block): 20,000 in Mar and Sep
    last = blk[5750000]["INPUT"][-1]
    T(f"C{last}").setString("Quoted premium"); T(f"I{last}").setValue(20000); T(f"O{last}").setValue(20000)
    doc.calculateAll(); print("   typed 20,000 in Mar and Sep -> final total", round(sum(mon(5750000, 'FINAL')), 2), "(expect 40,000)", "| status:", S())
    run(); print("   rebuilt:", S())
    rows = {int(r[0]): round(r[15], 2) for r in d() if isinstance(r[0], float) and int(r[0]) in (5750000, 5755000)}
    q = doc.Sheets.getByName("QTRLY SUM").getCellRangeByPosition(0, 0, 8, 140).getDataArray()
    print("3) DETAIL rows:", rows, "| QTRLY 5750000:", [round(r[7], 2) for r in q if r[0] == 5750000], "| QTRLY status:", cell(doc, "QTRLY SUM", "B3").String)
    # 4) new vendor on an insurance account
    dv = doc.Sheets.getByName("DATA_VENDOR"); n = sum(1 for r in dv.getCellRangeByPosition(0, 1, 3, 400).getDataArray() if r[0] != "") + 2
    for c, v in enumerate(["42", "NEWINS", "NEW INSURER", "MR5750000"]): dv.getCellByPosition(c, n - 1).setString(v)
    dv.getCellByPosition(4, n - 1).setValue(100); doc.calculateAll(); print("4) new vendor on 5750000 ->", S())
    doc.close(True)
finally:
    lo.close()

# PROF_FEES: management fee = rate x rent base (flows from REVENUE / lease options), other accounts = forecast x (1 + %), DETAIL / QTRLY flow.
import sys, pandas as pd
from lo_driver import LO
from make_xlsm import embed_vba, cell
BLDG = sys.argv[2] if len(sys.argv) > 2 else "42"
from paths import MRI_DIR
SRC = MRI_DIR.rstrip("/") + "/"
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
    sh = doc.Sheets.getByName("PROF_FEES"); T = lambda ref: sh.getCellRangeByName(ref)
    run(); print("0) built:", S(), "| QTRLY:", cell(doc, "QTRLY SUM", "B3").String)
    det = lambda: doc.Sheets.getByName("DETAIL").getCellRangeByPosition(0, 0, 20, 140).getDataArray()
    def drow(acct, sec_after=None):
        for r in det():
            if r[0] == acct: return [c for c in r[3:15]]
    arr = sh.getCellRangeByPosition(0, 0, 20, 80).getDataArray()
    blk = {}; cur = None
    for i, r in enumerate(arr):
        if isinstance(r[0], float) and r[0] > 1e6: cur = int(r[0]); blk[cur] = {"r0": i + 1}
        if cur and r[4] in ("CALC", "FINAL"): blk[cur][r[4]] = i + 1
        if cur and r[4] == "INPUT": blk[cur].setdefault("INPUT", []).append(i + 1)
    def mon(acct, tag): return [round(c, 2) for c in sh.getCellRangeByPosition(6, blk[acct][tag] - 1, 17, blk[acct][tag] - 1).getDataArray()[0]]
    base = [x + y for x, y in zip(drow(4100000), drow(4150000))]
    exp_fee = [round(0.03 * b, 2) for b in base]
    print("1) management fee 3% x (base rent + concession):", "match" if mon(5800000, "CALC") == exp_fee else "MISMATCH", f"| total {sum(mon(5800000, 'FINAL')):,.2f}")
    for acct, pct in ((5805000, 0.05), (5855000, 0.0)):
        exp = [round(x * (1 + pct), 2) for x in fcst(acct)]
        print(f"   {acct} +{pct:.0%}: calc match {mon(acct, 'CALC') == exp} | total {sum(mon(acct, 'FINAL')):,.2f} (expect {sum(exp):,.2f})")
    tot = [round(r[15], 2) for r in det() if r[1] == "TOTAL PROFESSIONAL FEES/G&A"]
    print("   DETAIL TOTAL PROFESSIONAL FEES/G&A:", tot, "| expect", round(sum(exp_fee) + sum(fcst(5805000)) * 1.05 + sum(fcst(5855000)), 2))
    # 2) operating expense fee %
    T("G11").setValue(0.01); doc.calculateAll()
    opx = drow(4210000)
    exp2 = [round(0.03 * b + 0.01 * o, 2) for b, o in zip(base, opx)]
    print("2) +1% on operating expense reimbursement:", "match" if mon(5800000, "CALC") == exp2 else "MISMATCH", f"| fee total {sum(mon(5800000, 'FINAL')):,.2f}")
    T("G11").setValue(0.0)
    # 3) lease option flows into the fee
    cell(doc, "INPUT", "C31").setString("MTM"); doc.calculateAll()
    base_mtm = [x + y for x, y in zip(drow(4100000), drow(4150000))]
    print("3) lease option MTM -> rent base", round(sum(base)), "->", round(sum(base_mtm)), "| fee total", round(sum(mon(5800000, 'FINAL')), 2), "(expect", round(0.03 * sum(base_mtm), 2), ")")
    cell(doc, "INPUT", "C31").setString("Terminate"); doc.calculateAll()
    # 4) second base account blank, switch off, adjustments
    T("H12").setString(""); doc.calculateAll()
    print("4) concession account blank -> fee total", round(sum(mon(5800000, 'FINAL')), 2), "(expect", round(0.03 * sum(drow(4100000)), 2), ")")
    T("H12").setValue(4150000)
    r0 = blk[5805000]["r0"]; T(f"J{r0}").setString("No"); last = blk[5805000]["INPUT"][-1]
    T(f"C{last}").setString("Audit fee"); T(f"I{last}").setValue(23000); doc.calculateAll()
    print("   accounting Use calc=No + 23,000 typed in May ->", round(sum(mon(5805000, 'FINAL')), 2), "| status", S())
    run(); print("   rebuilt:", S())
    q = doc.Sheets.getByName("QTRLY SUM").getCellRangeByPosition(0, 0, 8, 140).getDataArray()
    print("5) QTRLY 5800000:", [round(r[7], 2) for r in q if r[0] == 5800000], "| DETAIL 5800000:", round(sum(drow(5800000)), 2), "| QTRLY status:", cell(doc, "QTRLY SUM", "B3").String)
    doc.close(True)
finally:
    lo.close()

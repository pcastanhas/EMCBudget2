# TAXES sheet: calculated payments / escrow / assessment vs an independent calculation, flow to DETAIL, option changes.
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
    tx = doc.Sheets.getByName("TAXES"); T = lambda ref: tx.getCellRangeByName(ref)
    run(); print("0) built:", S(), "| QTRLY:", cell(doc, "QTRLY SUM", "B3").String)
    arr = tx.getCellRangeByPosition(0, 0, 20, 80).getDataArray()
    rowof = {}
    cur = None
    for i, r in enumerate(arr):
        if isinstance(r[0], float) and r[0] > 1e6: cur = int(r[0])
        if cur and r[4] == "FINAL": rowof[cur] = i + 1
    def months(acct): return [round(c, 2) for c in tx.getCellRangeByPosition(6, rowof[acct] - 1, 17, rowof[acct] - 1).getDataArray()[0]]
    # expected (default inputs)
    py = sum(fcst(5710000)); by = py * 1.035; q12 = py / 4; q34 = (by - 2 * q12) / 2
    pay = {2: q12, 5: q12, 8: q34, 11: q34}
    exp_tax = [round(pay.get(m, 0), 2) for m in range(1, 13)]
    esc = max(a26[a26.acctnum == "MR1200000"][MON[:9]].sum().tolist())
    exp_esc = [round(esc - pay.get(m, 0), 2) for m in range(1, 13)]
    f45 = fcst(5745000); exp_bia = [round(x * 1.025, 2) for x in f45]
    print(f"1) defaults: prior-year tax {T('G9').Value:,.2f} (expect {py:,.2f}) | budget-year tax {T('G11').Value:,.2f} (expect {by:,.2f}) | escrow {T('G16').Value:,.2f} (expect {esc:,.2f})")
    print("   5710000 months match:", months(5710000) == exp_tax, "| 1200000 match:", months(1200000) == exp_esc, "| 5745000 match:", months(5745000) == exp_bia)
    d = lambda: doc.Sheets.getByName("DETAIL").getCellRangeByPosition(0, 0, 20, 120).getDataArray()
    heads = [r[0] for r in d() if r[0] in ("TAXES",)]; tot = [round(r[15], 2) for r in d() if r[1] == "TOTAL TAXES"]
    rows = {int(r[0]): round(r[15], 2) for r in d() if isinstance(r[0], float) and int(r[0]) in (1200000, 5710000, 5730000, 5735000, 5745000)}
    print("   DETAIL TAXES section present:", bool(heads), "| rows:", rows, "| TOTAL TAXES:", tot, "| expect", round(sum(exp_esc) + sum(exp_tax) + sum(exp_bia), 2))
    # 2) legacy tie: assessed value x rate, 3.5% (2026 budget as prepared in Allwood file)
    T("G7").setValue(19000000); T("G8").setValue(0.02531); T("G10").setValue(0.035); T("G16").setValue(38729); doc.calculateAll()
    print(f"2) legacy inputs (19,000,000 x 2.531%, +3.5%, escrow 38,729): prior-year tax {T('G9').Value:,.2f} | budget-year {T('G11').Value:,.2f} | payments {[round(T('H'+str(r)).Value,2) for r in (12,13,14,15)]}")
    print("   escrow row total:", round(sum(months(1200000)), 2), "(legacy -32,973) | 5710000 total", round(sum(months(5710000)), 2), "(legacy 497,721)")
    run(); print("   after rebuild:", S())
    # 3) typed lines + zero escrow + changed months
    cons = rowof[5735000] - 3
    T(f"C{cons}").setString("Tax consultant"); T(f"L{cons}").setValue(250)
    run(); print("   consulting typed, rebuilt ->", S())
    T("G16").setValue(0); T("G12").setValue(1); doc.calculateAll()
    print("3) escrow 0 -> escrow row total:", round(sum(months(1200000)), 2), "| first payment moved to Jan ->", months(5710000)[:3], "| consulting Jun 250 ->", months(5735000)[5], "| status:", S())
    T("G13").setValue(13); doc.calculateAll(); print("   payment month 13 ->", S())
    T("G13").setValue(5); T("G12").setValue(2); T("G16").setValue(38729); doc.calculateAll()
    print("   restored ->", S())
    rows = {int(r[0]): round(r[15], 2) for r in d() if isinstance(r[0], float) and int(r[0]) in (1200000, 5710000, 5730000, 5735000, 5745000)}
    q = doc.Sheets.getByName("QTRLY SUM").getCellRangeByPosition(0, 0, 8, 120).getDataArray()
    print("4) QTRLY 5710000:", [round(r[7], 2) for r in q if r[0] == 5710000], "| DETAIL 5710000:", rows.get(5710000), "| QTRLY status:", cell(doc, "QTRLY SUM", "B3").String)
    doc.close(True)
finally:
    lo.close()

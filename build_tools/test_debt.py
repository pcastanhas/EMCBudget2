# DEBT: primary-mortgage roll-forward vs an independent calculation from the raw MRI extracts; DETAIL / QTRLY flow; inputs (rate, payment, balance, day count, use calc, actuals-thru).
import sys, calendar, pandas as pd
from lo_driver import LO
from make_xlsm import embed_vba, cell
BLDG = sys.argv[2] if len(sys.argv) > 2 else "42"
SRC = "/mnt/project/"
MON = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]
def load(f, cols_num):
    d = pd.read_excel(SRC + f, sheet_name="MRIQRY", dtype=str); d["entityid"] = d.entityid.str.strip(); d["acctnum"] = d.acctnum.str.strip()
    for m in cols_num: d[m] = pd.to_numeric(d[m])
    return d[d.entityid == BLDG]
a26 = load("Q1-Q3_2026_ACTUALS.xlsx", MON)
bal = load("2026_STARTING_BALANCES.xlsx", ["starting_balance"])
PR, IN = "MR2600000", "MR5910000"
def acts(acct): return a26[a26.acctnum == acct][MON].sum().tolist()
def amort(rate, pay, thru=9, basis="30/360", override=None):
    """independent roll-forward: returns (PY interest[12], PY principal[12], PY ending, BY interest[12], BY principal[12])"""
    den = 365.0 if basis == "Actual/365" else 360.0
    days = lambda y, m: 30 if basis == "30/360" else calendar.monthrange(y, m)[1]
    ai, ap = acts(IN), acts(PR)
    o = -float(bal[bal.acctnum == PR].starting_balance.sum()); pi, pp = [], []
    for m in range(12):
        if m < thru: i_, p_ = ai[m], ap[m]
        else:
            i_ = round(o * rate * days(2026, m + 1) / den + 1e-9, 2); p_ = max(0.0, min(pay - i_, o))
        pi.append(i_); pp.append(p_); o -= p_
    end = o; o = override if override is not None else end; bi, bp = [], []
    for m in range(12):
        i_ = round(o * rate * days(2027, m + 1) / den + 1e-9, 2); p_ = max(0.0, min(pay - i_, o)); bi.append(i_); bp.append(p_); o -= p_
    return pi, pp, end, bi, bp
lo = LO()
try:
    doc = lo.load(sys.argv[1]); embed_vba(lo, doc)
    run = lambda: (lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet"), doc.calculateAll())
    S = lambda: cell(doc, "DETAIL", "B3").String
    sh = doc.Sheets.getByName("DEBT"); T = lambda ref: sh.getCellRangeByName(ref)
    rowv = lambda r: [round(x, 2) for x in sh.getCellRangeByPosition(6, r - 1, 17, r - 1).getDataArray()[0]]
    det = lambda: doc.Sheets.getByName("DETAIL").getCellRangeByPosition(0, 0, 20, 140).getDataArray()
    def drow(acct):
        for r in det():
            if r[0] == acct: return [round(c, 2) for c in r[3:15]]
    def dlab(lbl):
        for r in det():
            if r[1] == lbl or r[0] == lbl: return [round(c, 2) for c in r[3:15]]
    blk = {}
    arr = sh.getCellRangeByPosition(0, 0, 20, 90).getDataArray(); cur = None
    for i, r in enumerate(arr):
        if i >= 36 and isinstance(r[0], float) and r[0] > 1e6: cur = int(r[0]); blk[cur] = {"r0": i + 1, "INPUT": []}
        if cur and r[4] in ("CALC", "FINAL"): blk[cur][r[4]] = i + 1
        if cur and r[4] == "INPUT": blk[cur]["INPUT"].append(i + 1)
    ok_all = True
    def chk(label, got, exp, tol=0.011):
        global ok_all
        good = all(abs(a - b) <= tol for a, b in zip(got, exp)) and len(got) == len(exp); ok_all &= good
        print(f"   {label}: {'MATCH' if good else 'MISMATCH'} | got {sum(got):,.2f} expect {sum(exp):,.2f}")
    run(); print("0) built:", S(), "| QTRLY:", cell(doc, "QTRLY SUM", "B3").String, "| blocks:", sorted(blk))
    rate = T("G13").Value; pay = T("G18").Value
    print(f"   defaults: lender {T('G10').String!r} rate {rate:.5%} implied {T('G14').Value:.5%} payment {pay:,.2f} (latest actual {T('G17').Value:,.2f}) basis {T('G15').String}")
    pi, pp, end, bi, bp = amort(rate, pay)
    chk("1) 2026 interest (actual Jan-Sep, calc Oct-Dec)", rowv(26), pi); chk("   2026 principal", rowv(27), pp)
    print(f"   2026 ending balance {T('R28').Value:,.2f} vs independent {end:,.2f} | 2027 opening {T('G32').Value:,.2f} (must equal 2026 ending)")
    ok_all &= abs(T("R28").Value - end) < 0.011 and abs(T("G32").Value - T("R28").Value) < 0.005
    chk("   2027 interest", rowv(33), bi); chk("   2027 principal", rowv(34), bp)
    chk("   DEBT block 5910000 final = interest", rowv(blk[5910000]["FINAL"]), bi); chk("   DEBT block 2600000 final = principal", rowv(blk[2600000]["FINAL"]), bp)
    chk("   DETAIL 5910000", drow(5910000), bi); chk("   DETAIL 2600000", drow(2600000), bp)
    tds = dlab("TOTAL DEBT SERVICE"); chk("   DETAIL TOTAL DEBT SERVICE = payment x 12", tds, [pay] * 12)
    noi = dlab("OPERATING INCOME BEFORE DEBT SERVICE"); nbc = dlab("N.O.I. BEFORE CAPITAL")
    chk("   N.O.I. BEFORE CAPITAL = operating income before debt service - debt service (12 months)", nbc, [a - b for a, b in zip(noi, tds)])
    print("   2027 ending balance", f"{T('R35').Value:,.2f}", "| QTRLY 5910000:", [r[7] for r in cell(doc, "QTRLY SUM", "A1").Spreadsheet.getCellRangeByPosition(0, 0, 8, 120).getDataArray() if r[0] == 5910000], "| status", S())
    # 2) rate change
    T("G13").Value = 0.045; doc.calculateAll(); pi2, pp2, end2, bi2, bp2 = amort(0.045, pay)
    chk("2) rate 4.5%: 2027 interest", rowv(33), bi2); chk("   2027 principal", rowv(34), bp2); chk("   DETAIL 5910000", drow(5910000), bi2); chk("   2026 ending balance (Oct-Dec re-forecast)", [T("R28").Value], [end2])
    print("   status:", S()); T("G13").Value = rate
    # 3) payment override
    T("G16").Value = 65000; doc.calculateAll(); _, _, end3, bi3, bp3 = amort(rate, 65000)
    chk("3) payment 65,000: 2027 principal", rowv(34), bp3); chk("   DETAIL TOTAL DEBT SERVICE", dlab("TOTAL DEBT SERVICE"), [65000] * 12); print("   payment used", T("G18").Value)
    T("G16").String = ""; doc.calculateAll(); print("   cleared -> payment used", T("G18").Value)
    # 4) balance override
    T("G20").Value = 10000000; doc.calculateAll(); _, _, _, bi4, bp4 = amort(rate, pay, override=10000000.0)
    chk("4) opening balance 10,000,000: 2027 interest", rowv(33), bi4); print("   flag:", T("T35").String, "| opening", T("G32").Value); T("G20").String = ""
    # 5) day count
    T("G15").String = "Actual/360"; doc.calculateAll(); pi5, pp5, end5, bi5, bp5 = amort(rate, pay, basis="Actual/360")
    chk("5) Actual/360: 2027 interest", rowv(33), bi5); chk("   2026 ending", [T("R28").Value], [end5]); T("G15").String = "Actual/365"; doc.calculateAll()
    pi5, pp5, end5, bi5, bp5 = amort(rate, pay, basis="Actual/365"); chk("   Actual/365: 2027 interest", rowv(33), bi5); T("G15").String = "30/360"; doc.calculateAll()
    # 6) use calc = No, typed amounts, adjustment
    r0 = blk[5910000]["r0"]; sh.getCellRangeByName(f"J{r0}").String = "No"
    ins = blk[5910000]["INPUT"]; sh.getCellRangeByName(f"G{ins[0]}").Value = 400000; sh.getCellRangeByName(f"H{ins[1]}").Value = 1250
    doc.calculateAll(); print("6) 5910000 use calc=No, 400,000 Jan adjustment line, 1,250 Feb open line -> DEBT final", rowv(blk[5910000]["FINAL"])[:3], "| calc row", sum(rowv(blk[5910000]["CALC"])), "| status", S())
    run(); print("   rebuilt:", S(), "| DETAIL 5910000 Jan/Feb", drow(5910000)[:2], "| TOTAL DEBT SERVICE", round(sum(dlab("TOTAL DEBT SERVICE")), 2))
    sh.getCellRangeByName(f"J{r0}").String = "Yes"; sh.getCellRangeByName(f"G{ins[0]}").String = ""; sh.getCellRangeByName(f"H{ins[1]}").String = ""; run()
    # 7) actuals through June (INPUT C15)
    doc.Sheets.getByName("INPUT").getCellRangeByName("C15").Value = 6; run()
    pay7 = T("G18").Value; pi7, pp7, end7, bi7, bp7 = amort(rate, pay7, thru=6)
    chk("7) actuals through Jun: 2026 principal", rowv(27), pp7); chk("   2027 interest", rowv(33), bi7); print(f"   payment used {pay7:,.2f} (June actual) | implied rate {T('G14').Value:.5%} | status {S()}")
    doc.Sheets.getByName("INPUT").getCellRangeByName("C15").Value = 9; run()
    # 8) maturity flag
    T("G19").Value = 46752.0; doc.calculateAll(); print("8) maturity 12/2027 ->", repr(T("I19").String)); T("G19").String = ""
    print("ALL DEBT CHECKS MATCH:", ok_all)
    doc.close(True)
finally:
    lo.close()

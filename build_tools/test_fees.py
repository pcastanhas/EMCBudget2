# OTHER_PROF_FEES: forecast x (1 + increase %), monthly OVERRIDE, adjustment, Use calc, User blocks, no-block flag; DETAIL / QTRLY flow and the two N.O.I. lines vs an independent calc from the raw MRI extracts.
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
    sh = doc.Sheets.getByName("OTHER_PROF_FEES")
    det = lambda: doc.Sheets.getByName("DETAIL").getCellRangeByPosition(0, 0, 20, 160).getDataArray()
    def drow(acct, sec=None):
        for r in det():
            if r[0] == acct: return [round(c, 2) for c in r[3:15]]
    def dlab(lbl):
        for r in det():
            if r[1] == lbl or r[0] == lbl: return [round(c, 2) for c in r[3:15]]
    def scan():
        arr = sh.getCellRangeByPosition(0, 0, 20, 260).getDataArray(); blk = {}; users = []; cur = None
        for i, r in enumerate(arr):
            if i >= 8 and isinstance(r[20], str) and r[20].startswith("U") and r[20][1:].isdigit(): users.append(i + 1); cur = None
            if i >= 8 and isinstance(r[0], float) and r[0] > 1e6: cur = int(r[0]); blk[cur] = {"r0": i + 1, "INPUT": []}
            if cur and r[4] in ("CALC", "FINAL", "OVR", "USED", "FCST"): blk[cur][r[4]] = i + 1
            if cur and r[4] == "INPUT": blk[cur]["INPUT"].append(i + 1)
        return blk, users
    mon = lambda r: [round(c, 2) for c in sh.getCellRangeByPosition(6, r - 1, 17, r - 1).getDataArray()[0]]
    ok_all = True
    def chk(label, got, exp, tol=0.011):
        global ok_all
        good = len(got) == len(exp) and all(abs(a - b) <= tol for a, b in zip(got, exp)); ok_all &= good
        print(f"   {label}: {'MATCH' if good else 'MISMATCH'} | got {sum(got):,.2f} expect {sum(exp):,.2f}")
    run(); blk, users = scan()
    print("0) built:", S(), "| QTRLY:", cell(doc, "QTRLY SUM", "B3").String, "| blocks:", sorted(blk), "| user blocks at rows", users)
    ACC = sorted(blk)
    tot_exp = 0
    for a in ACC:
        e = [round(x * 1.01, 2) for x in fcst(a)]; tot_exp += sum(e)
        chk(f"1) {a} calc = forecast x 1.01", mon(blk[a]["CALC"]), e); chk(f"   {a} DETAIL", drow(a), e)
    print(f"   section total {sum(dlab('TOTAL OTHER PROFESSIONAL FEES')):,.2f} (expect {tot_exp:,.2f}); leasing account 5820400 has a block: {5820400 in blk}")
    noi = dlab("N.O.I. BEFORE OTHER PROFESSIONAL FEES & DEBT SERVICE"); fees = dlab("TOTAL OTHER PROFESSIONAL FEES"); oi = dlab("OPERATING INCOME BEFORE DEBT SERVICE")
    ds = dlab("TOTAL DEBT SERVICE"); nbc = dlab("N.O.I. BEFORE CAPITAL")
    chk("   OPERATING INCOME BEFORE DEBT SERVICE = N.O.I. - other professional fees", oi, [a - b for a, b in zip(noi, fees)])
    chk("   N.O.I. BEFORE CAPITAL = operating income - debt service", nbc, [a - b for a, b in zip(oi, ds)])
    want = ("N.O.I. BEFORE OTHER PROFESSIONAL FEES & DEBT SERVICE", "TOTAL OTHER PROFESSIONAL FEES", "OPERATING INCOME BEFORE DEBT SERVICE", "TOTAL DEBT SERVICE", "N.O.I. BEFORE CAPITAL")
    order = [r[1] if r[1] in want else r[0] for r in det() if r[1] in want or r[0] in want]
    ok_all &= order == list(want)
    print("   order:", order)
    # 2) increase %
    b = blk[5780100]; sh.getCellRangeByName(f"H{b['r0']}").Value = 0.05; doc.calculateAll()
    chk("2) BAIT increase 5%", mon(b["CALC"]), [round(x * 1.05, 2) for x in fcst(5780100)]); sh.getCellRangeByName(f"H{b['r0']}").Value = 0.01
    # 3) override by month (legacy BAIT installments), blank months use calculated
    inst = {0: 33500, 3: 39400, 5: 33500, 8: 33500}
    for m, v in inst.items(): sh.getCellByPosition(6 + m, b["OVR"] - 1).Value = v
    doc.calculateAll(); f1 = fcst(5780100); exp = [inst.get(m, round(f1[m] * 1.01, 2)) for m in range(12)]
    chk("3) BAIT override Jan/Apr/Jun/Sep (legacy installments)", mon(b["FINAL"]), exp); print("   total", f"{sum(mon(b['FINAL'])):,.2f}", "(legacy $139,900) | amount-used row matches final:", mon(b["USED"]) == mon(b["FINAL"]), "| status", S())
    # override 0 = budget zero
    c = blk[5815050]; sh.getCellByPosition(6, c["OVR"] - 1).Value = 0; doc.calculateAll(); fm = fcst(5815050)
    print(f"   typed 0 in Jan on 5815050: Jan final {mon(c['FINAL'])[0]} (calculated would be {round(fm[0]*1.01,2)}) | Feb unchanged {mon(c['FINAL'])[1]} (calc {round(fm[1]*1.01,2)})")
    ok_all &= abs(mon(c["FINAL"])[0]) < 0.005 and abs(mon(c["FINAL"])[1] - round(fm[1] * 1.01, 2)) < 0.011
    sh.getCellByPosition(6, c["OVR"] - 1).String = ""; doc.calculateAll()
    ok_all &= abs(mon(c["FINAL"])[0] - round(fm[0] * 1.01, 2)) < 0.011; print("   cleared -> Jan final back to calculated:", mon(c["FINAL"])[0])
    # 4) use calc = No
    sh.getCellRangeByName(f"J{b['r0']}").String = "No"; doc.calculateAll()
    exp4 = [inst.get(m, 0.0) for m in range(12)]; chk("4) Use calc = No with the four overrides (other months zero)", mon(b["FINAL"]), exp4)
    sh.getCellRangeByName(f"J{b['r0']}").String = "Yes"
    # 5) adjustment and open line added to amount used
    d = blk[5820000]; sh.getCellByPosition(6 + 2, d["INPUT"][0] - 1).Value = -100; sh.getCellByPosition(6 + 5, d["INPUT"][2] - 1).Value = 300; doc.calculateAll()
    e5 = [round(x * 1.01, 2) for x in fcst(5820000)]; e5[2] -= 100; e5[5] += 300; chk("5) 5820000 adjustment -100 (Mar) + open line 300 (Jun)", mon(d["FINAL"]), e5)
    run(); print("   rebuilt:", S(), "| DETAIL 5780100 total", round(sum(drow(5780100)), 2), "| DETAIL 5820000 total", round(sum(drow(5820000)), 2))
    q = [r for r in cell(doc, "QTRLY SUM", "A1").Spreadsheet.getCellRangeByPosition(0, 0, 8, 160).getDataArray() if r[0] in (5780100, 5820000)]
    print("   QTRLY totals:", [(int(r[0]), round(r[7], 2)) for r in q], "| QTRLY status:", cell(doc, "QTRLY SUM", "B3").String)
    ok_all &= abs(q[0][7] - sum(drow(5780100))) < 0.011 and abs(q[1][7] - sum(drow(5820000))) < 0.011
    # reset 3 and 5 for later steps
    for m in inst: sh.getCellByPosition(6 + m, b["OVR"] - 1).String = ""
    sh.getCellByPosition(8, d["INPUT"][0] - 1).String = ""; sh.getCellByPosition(11, d["INPUT"][2] - 1).String = ""; run()
    # 6) user blocks
    ex = doc.Sheets.getByName("EXPENSE_ACCOUNTS"); lst = [r[0] for r in ex.getCellRangeByPosition(24, 1, 24, 100).getDataArray() if r[0]]
    print(f"6) dropdown list: {len(lst)} entries | has 5820700: {any(x.startswith('5820700') for x in lst)} | has leasing 5820400: {any(x.startswith('5820400') for x in lst)}")
    ok_all &= any(x.startswith("5820700") for x in lst) and not any(x.startswith("5820400") for x in lst)
    u = users[0]; pick = next(x for x in lst if x.startswith("5820700")); sh.getCellRangeByName(f"C{u}").String = pick; doc.calculateAll()
    for rr in range(u, u + 12):
        if sh.getCellByPosition(4, rr - 1).String == "INPUT": ur = rr; break
    sh.getCellByPosition(6 + 3, ur - 1).Value = 12000; doc.calculateAll()
    print("   picked", repr(pick), "-> A:", sh.getCellByPosition(0, u - 1).String, "| total", round(sh.getCellByPosition(18, u + 10).Value, 2), "| status", S())
    run(); print("   rebuilt:", S(), "| DETAIL 5820700 Apr:", drow(5820700)[3] if drow(5820700) else None, "(expect 12000) | section total", round(sum(dlab("TOTAL OTHER PROFESSIONAL FEES")), 2))
    ok_all &= drow(5820700) is not None and abs(drow(5820700)[3] - 12000) < 0.011
    sh.getCellRangeByName(f"C{users[1]}").String = next(x for x in lst if x.startswith("5780100")); doc.calculateAll()
    print("   duplicate of automatic block 5780100 ->", repr(sh.getCellByPosition(19, users[1] - 1).String), "|", S())
    ok_all &= "DUPLICATE" in sh.getCellByPosition(19, users[1] - 1).String and S().startswith("CHECK DATA")
    sh.getCellRangeByName(f"C{users[1]}").String = "5820400  Legal-Leasing"; doc.calculateAll()
    print("   leasing account typed in ->", repr(sh.getCellByPosition(19, users[1] - 1).String), "|", S())
    run(); print("   after rebuild ->", S())
    ok_all &= "NOT AN OTHER" in sh.getCellByPosition(19, users[1] - 1).String and S().startswith("CHECK DATA")
    sh.getCellRangeByName(f"C{users[1]}").String = ""; sh.getCellRangeByName(f"C{u}").String = ""; sh.getCellByPosition(9, ur - 1).String = ""; run(); print("   cleared ->", S())
    # 7) activity on an OTHER PROFESSIONAL FEES account that has no block
    vs = doc.Sheets.getByName("DATA_VENDOR"); vs.unprotect("")
    r_ = 2
    while vs.getCellByPosition(0, r_ - 1).String != "": r_ += 1
    for c_, v_ in enumerate(["42", "TESTV", "TEST VENDOR", "MR5820100", 100]): 
        cc = vs.getCellByPosition(c_, r_ - 1)
        (cc.setValue(v_) if isinstance(v_, (int, float)) else cc.setString(v_))
    doc.calculateAll(); print("7) vendor activity on 5820100 (no block) ->", S())
    ok_all &= S().startswith("CHECK DATA")
    print("ALL FEES CHECKS MATCH:", ok_all)
    doc.close(True)
finally:
    lo.close()

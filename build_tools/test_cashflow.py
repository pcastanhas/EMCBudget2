# PAYROLL (INPUT -> payroll line of the account's EXP_ block -> DETAIL), PARTNER DISTRIBUTIONS and MAINTENANCE RESERVE (INPUT -> DETAIL cash lines).
import sys
from lo_driver import LO
from make_xlsm import embed_vba, cell
lo = LO()
try:
    doc = lo.load(sys.argv[1]); embed_vba(lo, doc)
    run = lambda: (lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet"), doc.calculateAll())
    S = lambda: cell(doc, "DETAIL", "B3").String
    sh = lambda n: doc.Sheets.getByName(n)
    inp = sh("INPUT")
    det = lambda: sh("DETAIL").getCellRangeByPosition(0, 0, 20, 220).getDataArray()
    def dlab(lbl):
        for r in det():
            if r[0] == lbl or r[1] == lbl: return [round(c, 2) for c in r[3:15]] + [round(r[15], 2)] + [round(c, 2) for c in r[17:21]]
    def dacct(a):
        for r in det():
            if r[0] == a: return [round(c, 2) for c in r[3:15]]
    ok_all = True
    def chk(label, got, exp, tol=0.011):
        global ok_all
        good = got is not None and len(got) == len(exp) and all(abs(a - b) <= tol for a, b in zip(got, exp)); ok_all &= good
        print(f"   {label}: {'MATCH' if good else 'MISMATCH'} | got {sum(got) if got else None} expect {sum(exp):,.2f}")
    def setc(ref, v):
        c = inp.getCellRangeByName(ref); (c.setString(v) if isinstance(v, str) else c.setValue(v))
    def clr(ref): inp.getCellRangeByName(ref).setString("")
    def cum(start, flows):
        out, b = [], start
        for f in flows: b += f; out.append(round(b, 2))
        return out
    def payroll(annual, weeks5, summer=0, s1=6, s2=8):
        return [annual / 52 * (5 if m in weeks5 else 4) + (summer if s1 <= m <= s2 else 0) for m in range(1, 13)]
    def cash_checks(tag, dist_by_month, res_by_month):
        nbc = dlab("N.O.I. BEFORE CAPITAL")[:12]; tot = dlab("TOTAL IMPRVMTS & COMMISSIONS"); tot = tot[:12] if tot else [0.0] * 12
        start, rs = inp.getCellRangeByName("C38").Value, inp.getCellRangeByName("C42").Value
        ncf = [a - b for a, b in zip(nbc, tot)]
        chk(f"{tag} distributions line", dlab("Less: partner distributions")[:12], dist_by_month)
        chk(f"{tag} reserve additions line", dlab("Less: additions to maintenance reserve")[:12], res_by_month)
        nca = [a - b - c for a, b, c in zip(ncf, dist_by_month, res_by_month)]
        chk(f"{tag} NET CASH FLOW AFTER DISTRIBUTIONS & RESERVES", dlab("NET CASH FLOW AFTER DISTRIBUTIONS & RESERVES")[:12], nca)
        mr = dlab("MAINTENANCE RESERVE BALANCE"); oc = dlab("PROJECTED OPERATING CASH BALANCE"); tc = dlab("PROJECTED CASH BALANCE")
        chk(f"{tag} MAINTENANCE RESERVE BALANCE = {rs:,.0f} + cumulative additions", mr[:12], cum(rs, res_by_month))
        chk(f"{tag} PROJECTED OPERATING CASH BALANCE = (start cash - reserve) + cumulative net after distributions & reserves", oc[:12], cum(start - rs, nca))
        chk(f"{tag} PROJECTED CASH BALANCE (total) = start cash + cumulative (net cash flow - distributions); reserve additions are transfers", tc[:12], cum(start, [a - b for a, b in zip(ncf, dist_by_month)]))
        qe = lambda row: [row[2], row[5], row[8], row[11]]
        good = all(abs(a - b) < 0.011 for row in (mr, oc, tc) for a, b in zip(row[13:17], qe(row))) and all(abs(row[12] - row[11]) < 0.011 for row in (mr, oc, tc)); globals()["ok_all"] &= good
        print(f"   {tag} total column = December and quarter columns = quarter-end for the three balance rows: {'MATCH' if good else 'MISMATCH'} | Dec total cash {tc[11]:,.2f}")
    run(); print("0) built:", S(), "| QTRLY:", cell(doc, "QTRLY SUM", "B3").String)
    print("   INPUT defaults: payroll annual", inp.getCellRangeByName("C56").Value, "| 5-week months", [inp.getCellRangeByName(x).Value for x in ("C57", "D57", "C58", "D58")], "| weeks", inp.getCellRangeByName("C61").Value,
          "| distributions", [inp.getCellByPosition(2, r).Value for r in range(49, 53)], "months", [inp.getCellByPosition(3, r).Value for r in range(49, 53)], "| reserve start", inp.getCellRangeByName("C42").Value)
    chk("1) payroll 5340000 default (65,000 / 52 x 4 or 5 weeks = legacy 5,000 / 6,250)", dacct(5340000), payroll(65000, {3, 6, 8, 11}))
    hv = sh("EXP_HVAC").getCellRangeByPosition(0, 0, 20, 400).getDataArray()
    prow = [i + 1 for i, r in enumerate(hv) if isinstance(r[2], str) and "payroll from INPUT" in r[2]]
    print("   EXP_HVAC payroll line rows:", prow, "| line total", [round(hv[i - 1][18], 2) for i in prow])
    ok_all &= len(prow) == 1 and abs(hv[prow[0] - 1][18] - 65000) < 0.011
    d0 = {0: 180000, 3: 180000, 6: 180000, 9: 180000}; dist0 = [float(d0.get(m, 0)) for m in range(12)]
    cash_checks("1)", dist0, [0.0] * 12)
    # 2) payroll changes
    setc("C56", 78000); setc("C57", 1); setc("D57", 4); setc("C58", 7); setc("D58", 10); setc("C59", 1000); doc.calculateAll()
    chk("2) payroll 78,000, 5-week months Jan/Apr/Jul/Oct, summer help 1,000 Jun-Aug", dacct(5340000), payroll(78000, {1, 4, 7, 10}, 1000)); print("   status:", S())
    setc("D58", ""); doc.calculateAll(); print("   only three 5-week months -> weeks budgeted", inp.getCellRangeByName("C61").Value, "| flag:", repr(inp.getCellRangeByName("E61").String[:40]))
    ok_all &= inp.getCellRangeByName("C61").Value == 51 and inp.getCellRangeByName("E61").String.startswith("NOT 52")
    setc("D58", 11); setc("C56", 65000); setc("C57", 3); setc("D57", 6); setc("C58", 8); setc("C59", 0); doc.calculateAll()
    chk("   restored", dacct(5340000), payroll(65000, {3, 6, 8, 11}))
    # 3) distributions
    setc("C51", 200000); setc("D51", 5); doc.calculateAll(); d1 = {0: 180000, 4: 200000, 6: 180000, 9: 180000}
    cash_checks("3)", [float(d1.get(m, 0)) for m in range(12)], [0.0] * 12)
    setc("C49", "No"); doc.calculateAll(); chk("   distributions anticipated = No", dlab("Less: partner distributions")[:12], [0.0] * 12)
    setc("C49", "Yes"); setc("C51", 180000); setc("D51", 4); doc.calculateAll()
    setc("D52", 13); doc.calculateAll(); print("   distribution month 13 ->", S()); ok_all &= S().startswith("CHECK DATA")
    setc("D52", 7); doc.calculateAll(); print("   restored ->", S())
    # 4) reserves
    setc("C43", 25000); setc("C45", 40000); doc.calculateAll(); r1 = {2: 25000, 8: 40000}
    cash_checks("4)", dist0, [float(r1.get(m, 0)) for m in range(12)])
    print("   a reserve addition moves cash between operating and reserve: total cash row unchanged by it (checked above)")
    setc("C42", 150000); doc.calculateAll(); cash_checks("   reserve start 150,000", dist0, [float(r1.get(m, 0)) for m in range(12)])
    setc("C42", 100000); setc("C43", 0); setc("C45", 0); setc("D46", 0); setc("C46", 5000); doc.calculateAll(); print("   reserve addition with month 0 ->", S()); ok_all &= S().startswith("CHECK DATA")
    setc("D46", 12); setc("C46", 0); doc.calculateAll(); print("   restored ->", S())
    # 5) rebuild and QTRLY unaffected
    run(); print("5) rebuilt:", S(), "| QTRLY:", cell(doc, "QTRLY SUM", "B3").String)
    print("ALL CASHFLOW CHECKS MATCH:", ok_all)
    doc.close(True)
finally:
    lo.close()

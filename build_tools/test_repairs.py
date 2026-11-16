# REPAIRS_CAPITAL: major repairs -> expense accounts on DETAIL (+ memo lines on EXP_ blocks), capital improvements -> one capital account in
# section 3 (IMPRVMTS & COMMISSIONS), NET CASH FLOW / PROJECTED CASH BALANCE, spreads, monthly all-or-nothing schedule, Include = No, flags, status, protection.
# Layout (build_v2.py): major repair lines rows 12-51, capital lines rows 59-83, capital account D56; columns B include, C description, F account text, I cost, J spent,
# K budget-year amount, L used, M start month, N months, P:AA monthly amounts (input), AB:AM budget by month, AN total, AO line check.
import sys
from lo_driver import LO
from make_xlsm import embed_vba, cell

REP0, CAP0, CAPACCT = 12, 59, 56
lo = LO()
try:
    doc = lo.load(sys.argv[1]); embed_vba(lo, doc)
    run = lambda: (lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet"), doc.calculateAll())
    S = lambda: cell(doc, "DETAIL", "B3").String
    Q = lambda: cell(doc, "QTRLY SUM", "B3").String
    sh = lambda n: doc.Sheets.getByName(n)
    rc, inp = sh("REPAIRS_CAPITAL"), sh("INPUT")
    det = lambda: sh("DETAIL").getCellRangeByPosition(0, 0, 20, 220).getDataArray()

    def dlab(lbl):
        for r in det():
            if r[0] == lbl or r[1] == lbl: return [round(c, 2) for c in r[3:15]] + [round(r[15], 2)]

    def dacct(a):
        for r in det():
            if r[0] == a: return [round(c, 2) for c in r[3:15]]
        return [0.0] * 12
    ok_all = True

    def chk(label, got, exp, tol=0.011):
        global ok_all
        good = got is not None and len(got) == len(exp) and all(abs(a - b) <= tol for a, b in zip(got, exp)); ok_all &= good
        print(f"   {label}: {'MATCH' if good else 'MISMATCH'} | got {round(sum(got), 2) if got else None} expect {sum(exp):,.2f}")

    def truth(label, cond):
        global ok_all
        ok_all &= bool(cond); print(f"   {label}: {'OK' if cond else 'FAIL'}")

    def setc(sheet, ref, v):
        c = sheet.getCellRangeByName(ref); (c.setString(v) if isinstance(v, str) else c.setValue(v))

    def clear(sheet, *refs):
        for ref in refs: sheet.getCellRangeByName(ref).setString("")

    def flag(r): return rc.getCellRangeByName(f"AO{r}").String

    def months(r): return [round(rc.getCellByPosition(26 + m, r - 1).Value, 2) for m in range(1, 13)]

    # account dropdown text from the EXPENSE_ACCOUNTS list
    ea = sh("EXPENSE_ACCOUNTS").getCellRangeByPosition(0, 0, 29, 2001).getDataArray()
    lst = {int(r[28]): r[29] for r in ea[1:] if r[28] not in ("", None)}          # AC -> AD
    txt = lambda a: lst[a]

    def cash_check(tag):
        nbc = dlab("N.O.I. BEFORE CAPITAL")[:12]; tot = dlab("TOTAL IMPRVMTS & COMMISSIONS"); tot = tot[:12] if tot else [0.0] * 12
        ncf = dlab("NET CASH FLOW")
        chk(f"{tag} NET CASH FLOW = N.O.I. before capital - total imprvmts & commissions", ncf[:12], [a - b for a, b in zip(nbc, tot)])

    # ---- 0) defaults
    run(); print("0) built:", S(), "| QTRLY:", Q())
    truth("status OK and QTRLY OK with both grids empty", S().startswith("OK") and Q().startswith("OK"))
    base_toe = dlab("TOTAL OPERATING EXPENSES")[:12]; base_ex = dlab("TOTAL EXPENSES")[:12]
    on_detail = {int(r[0]) for r in det() if isinstance(r[0], float) and r[0] >= 1000000}
    cands = [a for a in sorted(lst) if a in on_detail]
    new_acct = [a for a in sorted(lst) if a not in on_detail][0]
    print(f"repair dropdown accounts: {len(lst)} | accounts on DETAIL that are in the list: {len(cands)} | account with no DETAIL row yet: {new_acct}")
    # a) protection: unlocked cells = light-blue inputs only
    def locked(ref): return rc.getCellRangeByName(ref).CellProtection.IsLocked
    truth("repair line: B, C, F, G, H, I, J, K, M, N, P, AA, AP unlocked", all(not locked(f"{c}{REP0}") for c in ("B", "C", "F", "G", "H", "I", "J", "K", "M", "N", "P", "AA", "AP")))
    truth("repair line: D, E, L, O, AB, AM, AN, AO locked", all(locked(f"{c}{REP0}") for c in ("D", "E", "L", "O", "AB", "AM", "AN", "AO")))
    truth("capital line: F (account name) locked, D56 account unlocked", locked(f"F{CAP0}") and not locked(f"D{CAPACCT}"))
    truth("sheet protected (no inserting or deleting rows)", rc.isProtected())

    # ---- 1) major repairs on existing DETAIL accounts
    A1, A2 = 5225000, [a for a in cands if a != 5225000][0]               # 5225000 carries a tenant-billed offset
    b1, b2, bn = dacct(A1), dacct(A2), dacct(new_acct)
    r1, r2 = REP0, REP0 + 1
    setc(rc, f"F{r1}", txt(A1)); setc(rc, f"C{r1}", "Test repair 1"); setc(rc, f"I{r1}", 15000); setc(rc, f"M{r1}", 6)
    setc(rc, f"F{r2}", txt(A2)); setc(rc, f"C{r2}", "Test repair 2"); setc(rc, f"I{r2}", 65000); setc(rc, f"J{r2}", 20000); setc(rc, f"M{r2}", 5); setc(rc, f"N{r2}", 3)
    doc.calculateAll(); print(f"1) repair 1: {A1} 15,000 in Jun; repair 2: {A2} cost 65,000 less 20,000 spent over 3 months from May ->", S())
    truth("both lines counted (tag REPAIR) and no flags", rc.getCellRangeByName(f"E{r1}").String == "REPAIR" and flag(r1) == "" and flag(r2) == "")
    e1 = [0.0] * 12; e1[5] = 15000
    e2 = [0.0] * 12; e2[4] = e2[5] = e2[6] = 15000
    chk(f"DETAIL {A1} = before + 15,000 in June (tenant-billed offset kept)", dacct(A1), [a + b for a, b in zip(b1, e1)])
    chk(f"DETAIL {A2} = before + 45,000 spread May-Jul", dacct(A2), [a + b for a, b in zip(b2, e2)])
    truth("no rebuild needed when the account is already on DETAIL", S().startswith("OK"))
    chk("total on sheet (rows 7 / 52) = both repairs", [rc.getCellRangeByName("AN7").Value], [60000])
    toe = dlab("TOTAL OPERATING EXPENSES")[:12]
    chk("TOTAL OPERATING EXPENSES rises by the repairs (both accounts are operating expense sections)", [a - b for a, b in zip(toe, base_toe)], [x + y for x, y in zip(e1, e2)])
    # memo line on the EXP_ block
    nz = lambda row: [v if isinstance(v, float) else 0.0 for v in row]
    for shn in ("EXP_JANITORIAL", "EXP_HVAC", "EXP_LANDSCAPING", "EXP_GENERAL_RM", "EXP_ELEVATOR", "EXP_UTILITIES"):
        arr = sh(shn).getCellRangeByPosition(0, 0, 19, 700).getDataArray(); cur = None; memo = net = off = None; inp_sum = [0.0] * 12
        for r in arr:
            if isinstance(r[0], float) and r[0] >= 1000000: cur = int(r[0]); inp_sum = [0.0] * 12
            if cur == A1 and r[4] == "INPUT": inp_sum = [a + b for a, b in zip(inp_sum, nz(r[6:18]))]
            if cur == A1 and r[4] == "OFFSET": off = nz(r[6:18])
            if cur == A1 and r[4] == "MEMO": memo = nz(r[6:18])
            if cur == A1 and r[4] == "NET": net = nz(r[6:18])
        if memo is not None:
            chk(f"{shn} memo line for {A1}", memo, e1)
            chk(f"{shn} NET = inputs + tenant-billed offset + major repairs", net, [a + b + c for a, b, c in zip(inp_sum, off, memo)])
    # ---- 2) an account with no DETAIL row: needs a rebuild
    r3 = REP0 + 2
    setc(rc, f"F{r3}", txt(new_acct)); setc(rc, f"I{r3}", 8000); setc(rc, f"M{r3}", 2)
    doc.calculateAll(); print(f"2) repair on {new_acct} (no DETAIL row yet) ->", S())
    truth("status asks for a rebuild", S().startswith("REBUILD NEEDED"))
    run(); print("   rebuilt ->", S(), "| QTRLY:", Q())
    e3 = [0.0] * 12; e3[1] = 8000
    chk(f"DETAIL now has a row for {new_acct}", dacct(new_acct), [a + b for a, b in zip(bn, e3)]); truth("status OK after rebuild", S().startswith("OK") and Q().startswith("OK"))
    # ---- 3) monthly schedule: all or nothing
    setc(rc, f"I{r3}", 30000); setc(rc, f"N{r3}", 1)
    rc.getCellByPosition(15 + 4, r3 - 1).setValue(12000)                   # May (P=Jan -> col index 15; May = 19)
    rc.getCellByPosition(15 + 8, r3 - 1).setValue(18000)                   # Sep
    doc.calculateAll(); print("3) cost 30,000 with start month Feb, but May 12,000 + Sep 18,000 typed ->", S())
    ex = [0.0] * 12; ex[4] = 12000; ex[8] = 18000
    chk("typed months ARE the schedule (start month ignored, blank months = 0)", dacct(new_acct), [a + b for a, b in zip(bn, ex)])
    truth("no flag (typed months add to the budget-year amount)", flag(r3) == "" and S().startswith("OK"))
    rc.getCellByPosition(15 + 8, r3 - 1).setValue(10000)
    doc.calculateAll(); print("   Sep changed to 10,000 ->", flag(r3), "|", S())
    truth("flag when typed months do not add to the amount", flag(r3).startswith("CHECK - MONTHLY AMOUNTS") and S().startswith("CHECK DATA"))
    chk("typed schedule still used", dacct(new_acct), [a + b for a, b in zip(bn, [0, 0, 0, 0, 12000, 0, 0, 0, 10000, 0, 0, 0])])
    setc(rc, f"K{r3}", 22000); doc.calculateAll()
    truth("typing the budget-year amount 22,000 clears the flag", flag(r3) == "" and S().startswith("OK"))
    clear(rc, f"K{r3}")
    for m in range(12): rc.getCellByPosition(15 + m, r3 - 1).setString("")
    doc.calculateAll()
    chk("typed months cleared -> spread from the start month (30,000 in Feb)", dacct(new_acct), [a + b for a, b in zip(bn, [0, 30000] + [0] * 10)])
    # ---- 4) Include = No, months and spread flags
    setc(rc, f"B{r3}", "No"); doc.calculateAll()
    chk("Include = No removes the line from DETAIL", dacct(new_acct), bn); print("   flag:", flag(r3), "|", S())
    truth("excluded line is a NOTE, not a data error", flag(r3).startswith("NOTE - excluded") and S().startswith("OK"))
    setc(rc, f"B{r3}", "Yes"); clear(rc, f"M{r3}"); doc.calculateAll()
    print("4) cost but no start month ->", flag(r3), "|", S())
    truth("flag + CHECK DATA when there is no start month or monthly amounts", flag(r3).startswith("CHECK - ENTER A START MONTH") and S().startswith("CHECK DATA"))
    setc(rc, f"M{r3}", 11); setc(rc, f"N{r3}", 4); doc.calculateAll()
    print("   start Nov over 4 months ->", flag(r3), "| Nov, Dec:", rc.getCellByPosition(37, r3 - 1).Value, rc.getCellByPosition(38, r3 - 1).Value, "|", S())
    truth("spread past December is a NOTE only", flag(r3).startswith("NOTE - SPREAD RUNS") and S().startswith("OK"))
    # ---- 5) invalid account
    setc(rc, f"M{r3}", 2); setc(rc, f"N{r3}", 1)
    r4 = REP0 + 3
    setc(rc, f"F{r4}", "1345000  Building Improvements"); setc(rc, f"I{r4}", 5000); setc(rc, f"M{r4}", 3); doc.calculateAll()
    print("5) account that is not an expense account ->", rc.getCellRangeByName(f"E{r4}").String, "|", flag(r4), "|", S())
    truth("line ignored on DETAIL and flagged", rc.getCellRangeByName(f"E{r4}").String == "IGNORED" and flag(r4).startswith("CHECK - NOT A BUDGETABLE") and S().startswith("CHECK DATA"))
    clear(rc, f"F{r4}", f"I{r4}", f"M{r4}"); setc(rc, f"I{r4}", 5000); setc(rc, f"M{r4}", 3); doc.calculateAll()
    truth("no account -> CHECK - SELECT A GL ACCOUNT", flag(r4).startswith("CHECK - SELECT A GL"))
    clear(rc, f"I{r4}", f"M{r4}"); doc.calculateAll()
    # clear repair lines 2-3 for the capital tests; keep line 1
    for rr in (r2, r3): clear(rc, f"F{rr}", f"C{rr}", f"I{rr}", f"J{rr}", f"M{rr}"); setc(rc, f"N{rr}", 1)
    run(); print("   repairs 2-3 cleared and rebuilt ->", S())
    # ---- 6) capital improvements -> 1345000
    cap_base = dlab("TOTAL IMPRVMTS & COMMISSIONS"); cap_base = cap_base[:12] if cap_base else [0.0] * 12
    b45 = dacct(1345000)
    for k, (cost, mo, n) in enumerate([(25000, 4, 1), (72000, 8, 1), (10000, 10, 3)]):
        r = CAP0 + k; setc(rc, f"C{r}", f"Capital {k + 1}"); setc(rc, f"I{r}", cost); setc(rc, f"M{r}", mo); setc(rc, f"N{r}", n)
    doc.calculateAll(); print("6) capital: 25,000 Apr; 72,000 Aug; 10,000 over 3 months from Oct ->", S())
    run(); print("   rebuilt ->", S(), "| QTRLY:", Q())
    ec = [0.0] * 12; ec[3] = 25000; ec[7] = 72000; ec[9] = 3333.33; ec[10] = 3333.33; ec[11] = 3333.34
    chk("DETAIL 1345000 (one-month, 3-month spread, remainder in last month)", dacct(1345000), [a + b for a, b in zip(b45, ec)])
    tot = dlab("TOTAL IMPRVMTS & COMMISSIONS"); chk("TOTAL IMPRVMTS & COMMISSIONS", tot[:12], [a + b for a, b in zip(cap_base, ec)]); cash_check("6)")
    chk("capital total on the sheet (row 8)", [rc.getCellRangeByName("AN8").Value], [sum(ec)])
    # ---- 7) duplicate note across the two grids
    setc(rc, f"F{r2}", txt(A2)); setc(rc, f"I{r2}", 25000); setc(rc, f"M{r2}", 4); doc.calculateAll()
    print("7) repair 25,000 in April + capital line 25,000 in April ->", flag(r2), "|", flag(CAP0), "|", S())
    truth("both lines carry the duplicate NOTE and status is OK", flag(r2).startswith("NOTE - POSSIBLE DUPLICATE") and flag(CAP0).startswith("NOTE - POSSIBLE DUPLICATE") and S().startswith("OK"))
    clear(rc, f"F{r2}", f"I{r2}", f"M{r2}")
    # ---- 8) capital account changed
    setc(rc, f"D{CAPACCT}", 5860000); doc.calculateAll(); print("8) capital account set to 5860000 ->", S())
    run(); print("   rebuilt ->", S())
    truth("5860000 on DETAIL carries the capital amounts (>= 107,000)", sum(dacct(5860000)) >= 107000 - 0.01)
    chk("1345000 back to its starting amount", dacct(1345000), b45)
    setc(rc, f"D{CAPACCT}", 1234567); doc.calculateAll(); print("   capital account not in chart ->", flag(CAPACCT), "|", S())
    truth("flag + CHECK DATA for an unknown capital account", flag(CAPACCT).startswith("CHECK - CAPITAL GL") and S().startswith("CHECK DATA"))
    setc(rc, f"D{CAPACCT}", 1345000); run()
    # ---- 9) section 3 combination with a REVENUE cost line + starting cash input
    top = sh("REVENUE").getCellRangeByPosition(0, 0, 20, 60).getDataArray()
    cost = [i + 1 for i, r in enumerate(top) if r[4] == "COST" and r[3] == 5860000][0]
    before = dacct(5860000)
    for c in "GHI": sh("REVENUE").getCellRangeByName(f"{c}{cost}").setValue(4000)
    doc.calculateAll(); run(); chk("9) 5860000: capital lines now on 1345000; REVENUE cost line 4,000 x 3 months", dacct(5860000), [a + (4000 if i < 3 else 0) for i, a in enumerate(before)])
    for c in "GHI": sh("REVENUE").getCellRangeByName(f"{c}{cost}").setString("")
    cash_check("9)")
    run(); print("\nOVERALL:", "ALL PASS" if ok_all else "FAILURES ABOVE", "| final status:", S())
    doc.close(True)
finally:
    lo.close()

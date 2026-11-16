# DETAIL summary lines: TOTAL EXPENSES, TOTAL OPERATING EXPENSES, N.O.I. - position, formulas, flow when amounts change.
import sys
from lo_driver import LO
from make_xlsm import embed_vba, cell
lo = LO()
try:
    doc = lo.load(sys.argv[1]); embed_vba(lo, doc)
    run = lambda: (lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet"), doc.calculateAll())
    S = lambda: cell(doc, "DETAIL", "B3").String
    sh = lambda n: doc.Sheets.getByName(n)
    def lines():
        arr = sh("DETAIL").getCellRangeByPosition(0, 0, 20, 160).getDataArray()
        out = []
        for i, r in enumerate(arr, 1):
            lbl = r[1] if isinstance(r[1], str) and r[1] else (r[0] if isinstance(r[0], str) else "")
            if lbl: out.append((i, lbl, r[15] if isinstance(r[15], float) else None, [c for c in r[3:15]]))
        return out
    def show(tag):
        L = lines(); tot = {l: (i, v) for i, l, v, _ in L if l.startswith("TOTAL") or l.startswith("N.O.I.")}
        print(tag, S())
        for l, (i, v) in tot.items(): print(f"   row {i:3d}  {l[:56]:56s} {v:>14,.2f}")
        return L, tot
    run(); L, tot = show("0) default build ->")
    order = [l for _, l, v, _ in L if l.startswith("TOTAL") or l.startswith("N.O.I.")]
    print("   order:", order)
    def v(prefix): return [val for _, l, val, _ in L if l.startswith(prefix)][0]
    te = sum(val for _, l, val, _ in L if l.startswith("TOTAL") and l[6:] in ("UTILITIES", "JANITORIAL & CLEANING", "ELEVATOR", "HVAC & PAYROLL", "LANDSCAPING & GROUNDS", "GENERAL R&M"))
    print(f"   TOTAL EXPENSES {v('TOTAL EXPENSES'):,.2f} vs sum of the six sections {te:,.2f}")
    toe = te + v("TOTAL TAXES") + v("TOTAL INSURANCE") + v("TOTAL PROFESSIONAL")
    print(f"   TOTAL OPERATING EXPENSES {v('TOTAL OPERATING EXPENSES'):,.2f} vs expenses+taxes+insurance+prof fees {toe:,.2f}")
    print(f"   N.O.I. {v('N.O.I.'):,.2f} vs income - operating expenses {v('TOTAL INCOME') - toe:,.2f}")
    # monthly check of the NOI line
    rows = {l: m for _, l, _, m in L}
    noi_m = [m for _, l, _, m in L if l.startswith("N.O.I.")][0]; inc_m = rows["TOTAL INCOME"]; toe_m = rows["TOTAL OPERATING EXPENSES"]
    print("   monthly N.O.I. = income - operating expenses for all 12 months:", all(abs(noi_m[k] - (inc_m[k] - toe_m[k])) < 0.01 for k in range(12)))
    # 1) type expenses and rebuild
    def first_input(name, acct):
        arr = sh(name).getCellRangeByPosition(0, 0, 20, 1500).getDataArray()
        h = [i for i, r in enumerate(arr) if r[0] == acct][0]
        return [j + 1 for j in range(h + 1, len(arr)) if arr[j][4] == "INPUT"][0]
    for name, acct, val in (("EXP_UTILITIES", 5110000, 25000), ("EXP_ELEVATOR", 5265000, 700), ("EXP_GENERAL_RM", 5410000, 2500)):
        r = first_input(name, acct)
        for c in "GHIJKLMNOPQR": sh(name).getCellRangeByName(f"{c}{r}").setValue(val)
    run(); L, tot = show("1) typed Utilities 300,000 + Elevator 8,400 + General R&M 30,000 ->")
    v0 = lambda p: [val for _, l, val, _ in L if l.startswith(p)][0]
    print(f"   TOTAL EXPENSES {v0('TOTAL EXPENSES'):,.2f} (expect {300000 + 8400 + 30000 - 704.64 - 10000.80 :,.2f}: net of the tenant-billed offsets on 5435000 and 5225000)")
    print(f"   QTRLY status: {cell(doc, 'QTRLY SUM', 'B3').String}")
    # 2) a higher monthly escrow deposit raises TOTAL TAXES, operating expenses and lowers N.O.I. (a higher tax bill alone nets to zero inside escrow)
    t0 = (v0("TOTAL OPERATING EXPENSES"), v0("N.O.I."))
    sh("TAXES").getCellRangeByName("G16").setValue(45000); doc.calculateAll()
    L2 = lines(); a = [val for _, l, val, _ in L2 if l.startswith("TOTAL OPERATING EXPENSES")][0]; b = [val for _, l, val, _ in L2 if l.startswith("N.O.I.")][0]
    print(f"2) escrow deposit 38,729 -> 45,000 a month -> operating expenses {t0[0]:,.0f} -> {a:,.0f} (expect +{(45000 - 38729) * 12:,.0f}) | N.O.I. {t0[1]:,.0f} -> {b:,.0f} | sum unchanged {abs((a + b) - (t0[0] + t0[1])) < 0.01}")
    # 3) lease option MTM raises income and NOI
    cell(doc, "INPUT", "C31").setString("MTM"); doc.calculateAll()
    L3 = lines(); c_ = [val for _, l, val, _ in L3 if l.startswith("N.O.I.")][0]
    print(f"3) lease option MTM -> N.O.I. {b:,.0f} -> {c_:,.0f}")
    doc.close(True)
finally:
    lo.close()

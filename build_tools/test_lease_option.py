# Lease Options: Terminate / MTM / Renew vs an independent calculation from the raw MRI extract.
import sys, datetime as dt, pandas as pd
from lo_driver import LO
from make_xlsm import embed_vba, cell
BLDG = sys.argv[2] if len(sys.argv) > 2 else "42"
from paths import MRI_DIR
SRC = MRI_DIR.rstrip("/") + "/"
YEAR = 2027
te = pd.read_excel(SRC + "TENANT_LIST.xlsx", sheet_name="MRIQRY", dtype=str)
te["bldgid"] = te.bldgid.str.strip(); te = te[te.bldgid == BLDG].copy()
for c in ("leasid", "inccat", "suitid"): te[c] = te[c].str.strip()
for c in ("rentstrt", "expir", "vacate", "effdate", "enddate"): te[c] = pd.to_datetime(te[c], errors="coerce")
te["amount"] = pd.to_numeric(te.amount)
cc = pd.read_excel(SRC + "CHARGE_CODE_LIST.xlsx", sheet_name="MRIQRY", dtype=str)
gl = {r.Code.strip(): int(r.income_Account.strip()[2:]) for r in cc.itertuples()}
EXCL = {4150000}

def expected(option, pct):
    """{gl: total for the year} and {(lease, gl): monthly list}"""
    tot = {}
    for lease, g in te.groupby("leasid"):
        exp, vac = g.expir.iloc[0], g.vacate.iloc[0]
        for code, gc in g.groupby("inccat"):
            acct = gl[code]
            for m in range(1, 13):
                ms = pd.Timestamp(YEAR, m, 1)
                amt = 0.0
                in_term = (pd.isna(exp) or exp >= ms) and (pd.isna(vac) or vac >= ms)
                if in_term:
                    c = gc[gc.effdate <= ms].sort_values("effdate")
                    if not c.empty:
                        r = c.iloc[-1]
                        if pd.isna(r.enddate) or r.enddate >= ms: amt = r.amount
                elif option != "Terminate" and pd.notna(exp) and pd.isna(vac) and exp < ms and acct not in EXCL:
                    lm = pd.Timestamp(exp.year, exp.month, 1)
                    c = gc[gc.effdate <= lm].sort_values("effdate")
                    if not c.empty:
                        r = c.iloc[-1]
                        if pd.isna(r.enddate) or r.enddate >= lm:
                            amt = r.amount * ((1 + pct) if option == "Renew" else 1)
                tot[acct] = tot.get(acct, 0) + amt
    return tot

lo = LO()
try:
    doc = lo.load(sys.argv[1]); embed_vba(lo, doc)
    run = lambda: (lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet"), doc.calculateAll())
    run()
    inp = doc.Sheets.getByName("INPUT"); rev = doc.Sheets.getByName("LEASE_CALC"); det = doc.Sheets.getByName("DETAIL")
    print("status after build:", cell(doc, "DETAIL", "B3").String, "| default option:", inp.getCellRangeByName("C31").String, inp.getCellRangeByName("C32").Value)
    def rev_by_acct():
        arr = rev.getCellRangeByPosition(0, 7, 19, 1500).getDataArray(); out = {}
        for r in arr:
            if isinstance(r[3], float) and r[3] >= 1e6 and isinstance(r[18], float) and r[4] in ("MRI", "ADJ", "ADD", "CALC"): out[int(r[3])] = out.get(int(r[3]), 0) + r[18]
        return out
    def detail_total():
        return cell(doc, "DETAIL", "P12").Value
    ok = True
    for opt, pct in (("Terminate", .03), ("MTM", .03), ("Renew", .03), ("Renew", .05)):
        inp.getCellRangeByName("C31").setString(opt); inp.getCellRangeByName("C32").setValue(pct); doc.calculateAll()
        exp = expected(opt, pct); got = rev_by_acct()
        diffs = {a: (round(exp.get(a, 0), 2), round(got.get(a, 0), 2)) for a in set(exp) | set(got) if abs(exp.get(a, 0) - got.get(a, 0)) > 0.01 and a < 5000000}
        ok &= not diffs
        print(f"{opt:9s} {pct:.0%}: income-account totals vs independent calc -> {'MATCH' if not diffs else diffs} | base rent {got.get(4100000, 0):,.0f} | DETAIL total income {detail_total():,.0f} | status {cell(doc, 'DETAIL', 'B3').String}")
    # per-lease override: first lease expiring in the year -> Renew while global is Terminate
    inp.getCellRangeByName("C31").setString("Terminate"); doc.calculateAll()
    base = rev_by_acct()
    arr = rev.getCellRangeByPosition(0, 7, 20, 400).getDataArray()
    hdr = [i for i, r in enumerate(arr) if isinstance(r[20], float) and r[20] == 1.0][0] + 8   # block 1 header row
    exp_cell = rev.getCellRangeByName(f"F{hdr}").Value
    print("block-1 header row", hdr, "tenant", rev.getCellRangeByName(f"C{hdr}").String, "expiry serial", exp_cell)
    # find a lease that expires in the year
    for i, r in enumerate(arr):
        if isinstance(r[20], float) and r[20] >= 1 and isinstance(r[5], float) and r[5] > (dt.date(YEAR, 1, 1) - dt.date(1899, 12, 30)).days and r[5] < (dt.date(YEAR, 12, 31) - dt.date(1899, 12, 30)).days:
            h = i + 8; break
    print("override lease:", rev.getCellRangeByName(f"C{h}").String, "| expires", rev.getCellRangeByName(f"F{h}").String)
    ld = doc.Sheets.getByName("LEASE_DEALS"); lda = ld.getCellRangeByPosition(0, 0, 40, 300).getDataArray()
    hdr_ = list(lda[11]); c_lease, c_if, c_pct = hdr_.index("MRI lease renewed"), hdr_.index("If expired (in-place)"), hdr_.index("Renewal % (in-place)")
    ldr = [i for i, r in enumerate(lda) if r[0] == "MRI" and str(r[c_lease]).strip() == rev.getCellRangeByName(f"B{h}").String.strip()][0]
    ld.getCellByPosition(c_if, ldr).setString("Renew"); ld.getCellByPosition(c_pct, ldr).setValue(0.10); doc.calculateAll()   # override typed on LEASE_DEALS
    after = rev_by_acct()
    print("override Renew 10% on that lease only: base rent", f"{base.get(4100000, 0):,.0f}", "->", f"{after.get(4100000, 0):,.0f}", "| applied cells:", rev.getCellRangeByName(f"I{h}").String, rev.getCellRangeByName(f"J{h}").Value, "| status", cell(doc, "DETAIL", "B3").String)
    # flag text on the first MRI row of that block
    print("flag on first MRI row:", repr(rev.getCellRangeByName(f"T{h+1}").String))
    print("ALL OPTION TOTALS MATCH:", ok)
    doc.close(True)
finally:
    lo.close()

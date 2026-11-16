# LEASE_DEALS: one line per deal -> DEAL_SCHEDULE months vs an independent calculation, REVENUE "LEASE DEALS" section, DETAIL accounts,
# renewal stops the MRI lease's MTM/Renew forecast, line checks and the DETAIL status.   python3 test_lease_deals.py BASE.xlsx
import sys, datetime as dt
from lo_driver import LO
from make_xlsm import embed_vba, cell

LO_EPOCH = dt.date(1899, 12, 30)
serial = lambda d: (d - LO_EPOCH).days
fails = []
def chk(label, got, exp, tol=0.01):
    ok = all(abs(a - b) <= tol for a, b in zip(got, exp)) and len(got) == len(exp) if isinstance(exp, list) else abs(got - exp) <= tol
    print(("   ok   " if ok else "   FAIL ") + label + ("" if ok else f" | got {got} expected {exp}"))
    if not ok: fails.append(label)

lo = LO()
try:
    doc = lo.load(sys.argv[1]); embed_vba(lo, doc)
    run = lambda: (lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet"), doc.calculateAll())
    sh = doc.Sheets.getByName
    S = lambda: cell(doc, "DETAIL", "B3").String
    ld, ds, rv, inp = sh("LEASE_DEALS"), sh("DEAL_SCHEDULE"), sh("LEASE_CALC"), sh("INPUT")
    Y = int(inp.getCellRangeByName("C3").Value)
    # ---- layout read from the sheet (header row 12 labels, code row 13)
    hdr = ld.getCellRangeByPosition(0, 11, 40, 12).getDataArray()
    col_of = {}
    for j, t in enumerate(hdr[0]):
        if isinstance(t, str) and t and t not in col_of: col_of[t] = j
    codes = {str(hdr[1][j]): j for j in range(len(hdr[1])) if isinstance(hdr[1][j], str) and hdr[1][j] not in ("", "Charge code")}
    other = [j for j, t in enumerate(hdr[0]) if t == "(pick a code below)"]
    colA = [r[0] for r in ld.getCellRangeByPosition(0, 0, 0, 300).getDataArray()]
    FIRST = next(i for i, v in enumerate(colA, 1) if isinstance(v, float) and v == 1.0)      # deal line 1 (in-place rows above say "MRI")
    INP = [i for i, v in enumerate(colA, 1) if v == "MRI"]
    def put(line, label, v, j=None):
        c = ld.getCellByPosition(col_of[label] if j is None else j, FIRST + line - 2)
        if isinstance(v, dt.date): c.setValue(serial(v))
        elif isinstance(v, str): c.setString(v)
        else: c.setValue(v)
    run(); base = {}
    def detail():
        a = sh("DETAIL").getCellRangeByPosition(0, 0, 15, 200).getDataArray()
        return {int(r[0]): r[15] for r in a if isinstance(r[0], float) and r[0] >= 1e6}
    base = detail()
    print("0) empty deals:", S(), "| charge columns:", sorted(codes), "| pick-a-code columns:", len(other))
    # find the in-place lease in suite 1300 (expires during the budget year) for the renewal
    ts = sh("TENANT_SCHEDULE").getCellRangeByPosition(0, 0, 4, 400).getDataArray()
    kore = next((r[0], r[4]) for r in ts[1:] if str(r[1]).strip() == "1300")
    kore_id, kore_exp = str(kore[0]).strip(), LO_EPOCH + dt.timedelta(days=int(kore[1]))
    print("   renewal target: lease", kore_id, "expires", kore_exp)
    # ---- in-place section: one read-only row per REVENUE lease block, rates and 2027 revenue
    rva0 = rv.getCellRangeByPosition(0, 0, 20, 1500).getDataArray()
    blocks = [r for r in rva0 if isinstance(r[20], float) and str(r[1]).strip()]
    chk("in-place rows = REVENUE lease blocks", len(INP), len(blocks))
    lr = ld.getCellRangeByPosition(0, INP[0] - 1, 40, INP[-1] - 1).getDataArray()
    rc, rvc, stc, lc, sfc = col_of["Rent $/SF/yr"], col_of["2027 revenue ($)"], col_of["Line status"], col_of["MRI lease renewed"], col_of["Sq Ft"]
    tot_rows = [r[18] for r in rva0 if r[2] == "Total Rent for Lease"]
    chk("2027 revenue column = REVENUE lease totals", sum(r[rvc] for r in lr), sum(tot_rows[:len(INP)]))
    kr = next(r for r in lr if str(r[lc]).strip() == kore_id)
    tsd = sh("TENANT_SCHEDULE").getCellRangeByPosition(0, 0, 31, 400).getDataArray()
    feb = sum(r[19 + kore_exp.month - 1] for r in tsd[1:] if str(r[0]).strip() == kore_id and str(r[6]).strip() == "00")
    chk("expiring lease rent rate = expiry-month rent x 12 / SF", kr[rc], feb * 12 / kr[sfc])
    print("   suite 1300 status before renewal:", kr[stc])

    # ---- deals
    d1 = dict(suite="1700", tenant="NEW TENANT A", type="New lease", sf=3000, start=dt.date(Y, 3, 1), end=dt.date(Y + 3, 2, 28), ner=28, rent=30, incr=0.03,
              f1=(1, 2, None), f2=(13, 2, 0.5), oth={"50": 2.0, "10": 1.5}, legal=2, arch=1, permit=0.5, ti=25, comm=0.05, emc=0.015)
    d2 = dict(suite="1300", tenant="KORE RENEWAL", type="Renewal", lease=kore_id, sf=20013, start=dt.date(kore_exp.year, kore_exp.month, kore_exp.day) + dt.timedelta(days=1),
              end=dt.date(Y + 5, 2, 28), rent=25, incr=0, f1=None, f2=None, oth={}, legal=0, arch=0, permit=0, ti=0, comm=0, emc=0)
    d3 = dict(suite="2400", tenant="MID-MONTH START", type="New lease", sf=1000, start=dt.date(Y, 4, 15), end=dt.date(Y + 1, 12, 31), rent=24, incr=0.10,
              f1=None, f2=None, oth={"PICK:20": 1.0}, legal=0, arch=0, permit=0, ti=10, comm=0, emc=0)
    d4 = dict(suite="2800", tenant="STARTED LAST YEAR", type="Expansion", sf=2000, start=dt.date(Y - 1, 7, 1), end=dt.date(Y + 4, 6, 30), rent=24, incr=0.05,
              f1=None, f2=(13, 3, 0.5), oth={}, legal=0, arch=0, permit=0, ti=0, comm=0, emc=0)
    deals = [d1, d2, d3, d4]
    lab = {"suite": "Suite", "tenant": "Tenant / description", "type": "Type", "lease": "MRI lease renewed", "sf": "Sq Ft", "start": "Lease start",
           "end": "Lease end", "ner": "NER $/SF (opt.)", "rent": "Rent $/SF/yr", "incr": "Annual incr %", "legal": "Legal", "arch": "Architect",
           "permit": "Permits", "ti": "TI", "comm": "Broker %", "emc": "EMC %"}
    fr_cols = [j for j, t in enumerate(hdr[0]) if t in ("From lease month", "Months", "% of rent (blank = 100%)")]
    for n, d in enumerate(deals, 1):
        for k, l in lab.items():
            if d.get(k) not in (None, 0, ""): put(n, l, d[k])
        for p, per in enumerate((d["f1"], d["f2"])):
            if per:
                put(n, None, per[0], fr_cols[3 * p]); put(n, None, per[1], fr_cols[3 * p + 1])
                if per[2] is not None: put(n, None, per[2], fr_cols[3 * p + 2])
        for c, v in d["oth"].items():
            if c.startswith("PICK:"):
                ld.getCellByPosition(other[0], 12).setString(c[5:]); put(n, None, v, other[0])
            else:
                put(n, None, v, codes[c])
    doc.calculateAll()

    # ---- independent expectations
    def months(d):
        s = d["start"]; rs = dt.date(s.year + (s.month == 12 and s.day > 1), (s.month % 12) + 1 if s.day > 1 else s.month, 1)
        out = []
        for m in range(1, 13):
            md = dt.date(Y, m, 1); n = (md.year - rs.year) * 12 + md.month - rs.month + 1
            out.append(0 if (md > d["end"] or n < 1) else n)
        return out
    def exp(d):
        ix = months(d); sf = d["sf"]
        rent = [0 if i == 0 else sf * d["rent"] / 12 * (1 + d["incr"]) ** ((i - 1) // 12) for i in ix]
        def fr(i):
            t = 0
            for per in (d["f1"], d["f2"]):
                if per and per[0] <= i < per[0] + per[1]: t += 1 if per[2] is None else per[2]
            return t
        free = [-r * fr(i) for r, i in zip(rent, ix)]
        s = d["start"]
        mon = lambda off: dt.date(s.year + (s.month - 1 + off) // 12, (s.month - 1 + off) % 12 + 1, 1)
        lump = lambda off, amt: [amt if dt.date(Y, m, 1) == mon(off) else 0 for m in range(1, 13)]
        base_c = (d.get("ner") or d["rent"]) * sf * (d["end"] - d["start"]).days / 365
        comm = [a + b for a, b in zip(lump(-3, d["comm"] * base_c / 2), lump(3, d["comm"] * base_c / 2))]
        oth = {c: [0 if i == 0 else sf * v / 12 for i in ix] for c, v in d["oth"].items()}
        return dict(IDX=ix, RENT=rent, FREE=free, LEGAL=lump(-2, sf * d["legal"]), ARCH=lump(-2, sf * d["arch"]), PERMIT=lump(-2, sf * d["permit"]),
                    TI=lump(0, sf * d["ti"]), COMM=comm, EMC=lump(0, d["emc"] * base_c), OTH=oth)
    dsv = ds.getCellRangeByPosition(0, 0, 18, 700).getDataArray()
    rows = {}
    for r in dsv:
        if isinstance(r[0], float) and isinstance(r[1], str): rows[(int(r[0]), r[1])] = (r[2], list(r[6:18]))
    print("1) DEAL_SCHEDULE months vs independent calculation")
    code_role = {}
    for (n, role), (lbl, _) in rows.items():
        if n == 1 and role.startswith("C") and role[1:].isdigit(): code_role[str(lbl).strip().split(" ")[0]] = role
    for n, d in enumerate(deals, 1):
        e = exp(d)
        for role in ("IDX", "RENT", "FREE", "LEGAL", "ARCH", "PERMIT", "TI", "COMM", "EMC"):
            chk(f"deal {n} {role}", rows[(n, role)][1], e[role])
        for c, v in e["OTH"].items():
            chk(f"deal {n} charge {c.replace('PICK:', '')}", rows[(n, code_role[c.replace('PICK:', '')])][1], v)
    # ---- purple REVENUE tab: stale until Build Schedules, then by lease / deal
    print("   REVENUE status before rebuild:", cell(doc, "REVENUE", "B3").String)
    chk("REVENUE says REBUILD NEEDED after deals are entered", 0 if cell(doc, "REVENUE", "B3").String.startswith("REBUILD") else 1, 0)
    # ---- LEASE_CALC section, DETAIL
    run(); print("2) rebuilt ->", S(), "| REVENUE:", cell(doc, "REVENUE", "B3").String)
    chk("REVENUE status OK after rebuild", 0 if cell(doc, "REVENUE", "B3").String.startswith("OK") else 1, 0)
    rvp = sh("REVENUE").getCellRangeByPosition(0, 0, 16, 2000).getDataArray()
    tot_line = {str(x[0]): x[16] for x in rvp if isinstance(x[0], str) and x[0].startswith("TOTAL")}
    inc_exp = sum(sum(exp(d)[k]) for d in deals for k in ("RENT", "FREE")) + sum(sum(v) for d in deals for v in exp(d)["OTH"].values())
    cost_exp = sum(sum(exp(d)[k]) for d in deals for k in ("LEGAL", "ARCH", "PERMIT", "TI", "COMM", "EMC"))
    chk("REVENUE total new leases / renewals = deal income", tot_line["TOTAL NEW LEASES / RENEWALS / EXPANSIONS"], inc_exp)
    chk("REVENUE total leasing costs = deal costs", next(v for k, v in tot_line.items() if k.startswith("TOTAL LEASING COSTS")), cost_exp)
    chk("REVENUE total tenant revenue = LEASE_CALC lease totals", tot_line["TOTAL TENANT REVENUE"], cell(doc, "LEASE_CALC", "S6").Value)
    n_deal_hdr = sum(1 for x in rvp if str(x[1]).startswith("Deal "))
    chk("REVENUE lists one block per entered deal", n_deal_hdr, len(deals))
    chk("status OK after rebuild", 0 if S().startswith("OK") else 1, 0)
    det = detail()
    tot = lambda key: sum(sum(exp(d)[key]) for d in deals)
    tot_oth = lambda c: sum(sum(exp(d)["OTH"].get(c, [0])) for d in deals)
    # renewal: suite 1300 MRI lease after expiry is Terminate under every option, so base rent change = deals only (lease option default Terminate)
    chk("DETAIL 4100000 = before + deal rent", det[4100000] - base[4100000], tot("RENT"))
    chk("DETAIL 4150000 = before + free rent", det.get(4150000, 0) - base.get(4150000, 0), tot("FREE"))
    chk("DETAIL 4310000 = before + deal electric", det.get(4310000, 0) - base.get(4310000, 0), tot_oth("50"))
    chk("DETAIL 4210000 = before + deal OE", det.get(4210000, 0) - base.get(4210000, 0), tot_oth("10"))
    chk("DETAIL 4220000 (picked code 20) = deal CAM", det.get(4220000, 0) - base.get(4220000, 0), tot_oth("PICK:20"))
    chk("DETAIL 1365000 TI", det.get(1365000, 0) - base.get(1365000, 0), tot("TI"))
    chk("DETAIL 2105000 commissions (broker + EMC)", det.get(2105000, 0) - base.get(2105000, 0), tot("COMM") + tot("EMC"))
    chk("DETAIL 5820400 legal", det.get(5820400, 0) - base.get(5820400, 0), tot("LEGAL"))
    chk("DETAIL 5860000 architect", det.get(5860000, 0) - base.get(5860000, 0), tot("ARCH"))
    chk("DETAIL 5390000 permits", det.get(5390000, 0) - base.get(5390000, 0), tot("PERMIT"))
    # ---- renewal stops the MTM forecast of the renewed lease only
    inp.getCellRangeByName("C31").setString("MTM"); doc.calculateAll()
    rva = rv.getCellRangeByPosition(0, 0, 20, 1500).getDataArray()
    applied = {str(r[1]).strip(): r[8] for r in rva if str(r[1]).strip() and isinstance(r[8], str) and r[8] in ("Terminate", "MTM", "Renew")}
    chk("renewed lease applied option = Terminate under MTM", 0 if applied.get(kore_id) == "Terminate" else 1, 0)
    chk("other leases follow MTM", 0 if sum(v == "MTM" for k, v in applied.items() if k != kore_id) > 0 else 1, 0)
    kr = next(r for r in ld.getCellRangeByPosition(0, INP[0] - 1, 40, INP[-1] - 1).getDataArray() if str(r[lc]).strip() == kore_id)
    print("   suite 1300 status after renewal:", kr[stc]); chk("in-place row shows RENEWED", 0 if str(kr[stc]).startswith("RENEWED") else 1, 0)
    inp.getCellRangeByName("C31").setString("Terminate"); doc.calculateAll()
    # ---- line checks -> status
    print("3) line checks")
    put(5, "Tenant / description", "INCOMPLETE"); doc.calculateAll()
    print("   tenant only ->", S()); chk("missing SF/dates -> CHECK DATA", 0 if "LEASE_DEALS" in S() else 1, 0)
    put(5, "Sq Ft", 1000); put(5, "Lease start", dt.date(Y, 6, 1)); put(5, "Lease end", dt.date(Y, 5, 1)); doc.calculateAll()
    print("   end before start ->", S()); chk("end before start -> CHECK DATA", 0 if "LEASE_DEALS" in S() else 1, 0)
    put(5, "Lease end", dt.date(Y + 2, 5, 31)); put(5, "Type", "Renewal"); doc.calculateAll()
    print("   renewal without lease ->", S()); chk("renewal without MRI lease -> CHECK DATA", 0 if "LEASE_DEALS" in S() else 1, 0)
    put(5, "Type", "New lease"); doc.calculateAll()
    print("   fixed ->", S()); chk("fixed -> status back (REBUILD or OK)", 0 if not S().startswith("CHECK") else 1, 0)
    st = [r[col_of["Line status"]] for r in ld.getCellRangeByPosition(0, 14, 40, 18).getDataArray()]
    print("   line status column:", st)
    run(); print("   rebuilt ->", S())
    chk("final status OK", 0 if S().startswith("OK") else 1, 0)
    doc.close(True)
finally:
    lo.close()
print("OVERALL:", "ALL PASS" if not fails else f"{len(fails)} FAIL: {fails}")

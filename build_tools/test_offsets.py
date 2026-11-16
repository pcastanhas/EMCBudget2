# Tenant-billed offsets netted against expense / leasing accounts, offset+net rows on the section sheets, and the not-on-format flag.
import sys, pandas as pd
from lo_driver import LO
from make_xlsm import embed_vba, cell
lo = LO()
try:
    doc = lo.load(sys.argv[1]); embed_vba(lo, doc)
    run = lambda: (lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet"), doc.calculateAll())
    S = lambda: cell(doc, "DETAIL", "B3").String
    sh = lambda n: doc.Sheets.getByName(n)
    run(); print("0) built:", S())
    d = lambda: sh("DETAIL").getCellRangeByPosition(0, 0, 20, 160).getDataArray()
    def sections():
        out, cur = {}, None
        for r in d():
            if isinstance(r[0], str) and r[0].isupper() and len(r[0]) > 5 and r[0] not in ("ACCT",) and not isinstance(r[1], float): cur = r[0]; out[cur] = {}
            elif isinstance(r[0], float) and cur: out[cur][int(r[0])] = round(r[15], 2)
        return out
    sec = sections()
    print("   sections:", [k for k in sec if k not in ("ALLWOOD ASSOCIATES", "BUDGET FOR THE YEAR 2027")])
    print("   'TENANT CHARGES POSTED' section present:", any("TENANT CHARGES" in k for k in sec), "| 1527000 on DETAIL:", any(1527000 in v for v in sec.values()))
    j = [v for k, v in sec.items() if k.startswith("JANITORIAL")][0]; g = [v for k, v in sec.items() if k.startswith("GENERAL")][0]
    print("   5225000 in JANITORIAL (offset only):", j.get(5225000), "| 5435000 in GENERAL R&M:", g.get(5435000))
    rv = sh("LEASE_CALC").getCellRangeByPosition(0, 0, 22, 1000).getDataArray()
    fl = {}
    for r in rv:
        if isinstance(r[3], float) and r[21] != "": fl.setdefault(int(r[3]), set()).add(r[21])
    print("   LEASE_CALC flags:", fl)
    rev_by = {}
    for r in rv:
        if isinstance(r[3], float) and r[4] in ("MRI", "ADJ", "ADD", "CALC") and isinstance(r[18], float): rev_by[int(r[3])] = rev_by.get(int(r[3]), 0) + r[18]
    # 2) EXP input on 5225000
    arr = sh("EXP_JANITORIAL").getCellRangeByPosition(0, 0, 20, 400).getDataArray()
    h = [i for i, r in enumerate(arr) if r[0] == 5225000][0]
    inp = [k + 1 for k in range(h + 1, len(arr)) if arr[k][4] == "INPUT"][0]
    for c in "GHIJKLMNOPQR": sh("EXP_JANITORIAL").getCellRangeByName(f"{c}{inp}").setValue(9000)
    doc.calculateAll(); print("2) typed 9,000/mo on 5225000 ->", S())
    run(); sec = sections(); j = [v for k, v in sec.items() if k.startswith("JANITORIAL")][0]
    print("   after rebuild:", S(), "| DETAIL 5225000:", j.get(5225000), "(expect", round(108000 - rev_by[5225000], 2), ") | QTRLY status:", cell(doc, "QTRLY SUM", "B3").String)
    arr = sh("EXP_JANITORIAL").getCellRangeByPosition(0, 0, 20, 400).getDataArray()
    h = [i for i, r in enumerate(arr) if r[0] == 5225000][0]
    rows = {}
    for k in range(h + 1, len(arr)):
        if arr[k][4] in ("OFFSET", "NET", "FINAL") or arr[k][4] == "INPUT" and False: rows[arr[k][4]] = round(arr[k][18], 2)
        if arr[k][4] == "NET": break
    print("   block rows: total", rows.get("FINAL"), "| offset", rows.get("OFFSET"), "| net", rows.get("NET"))
    # 3) leasing cost: cost line on REVENUE + tenant-billed code to 1365000
    cc = sh("DATA_CHARGECODES").getCellRangeByPosition(0, 1, 3, 460).getDataArray()
    code = [r[0] for r in cc if r[3] == "MR1365000"][0]
    ld = sh("LEASE_DEALS"); lda = ld.getCellRangeByPosition(0, 0, 40, 300).getDataArray(); hdr_ = list(lda[11])
    col = lambda lbl: hdr_.index(lbl)
    pick = [j for j, v in enumerate(hdr_) if v == "(pick a code below)"][0]
    d1 = [i for i, r in enumerate(lda) if r[0] == 1.0][0]                      # deal line 1
    import datetime as _dt
    ser = lambda d: (d - _dt.date(1899, 12, 30)).days
    yr = int(cell(doc, "INPUT", "C3").Value)
    ld.getCellByPosition(pick, 12).setString(code)                              # pick-a-code column -> the code that posts to 1365000
    for lbl, v in (("Tenant / description", "OFFSET TEST"), ("Type", "New lease"), ("Sq Ft", 1000), ("Lease start", ser(_dt.date(yr - 1, 1, 1))), ("Lease end", ser(_dt.date(yr + 2, 12, 31)))):
        (ld.getCellByPosition(col(lbl), d1).setString(v) if isinstance(v, str) else ld.getCellByPosition(col(lbl), d1).setValue(v))
    ld.getCellByPosition(pick, d1).setValue(6.0)                                # 1,000 SF x 6.00 / 12 = 500 a month billed to the tenant
    ip = [i for i, r in enumerate(lda) if r[0] == "MRI" and isinstance(r[col("Sq Ft")], float) and r[col("Sq Ft")] > 0][0]
    sf_ = lda[ip][col("Sq Ft")]
    ld.getCellByPosition(col("TI"), ip).setValue(12000 / sf_); ld.getCellByPosition(col("Costs paid in month (in-place)"), ip).setValue(1)   # TI 12,000 in January
    doc.calculateAll(); run()
    sec = sections(); im = [v for k, v in sec.items() if k.startswith("IMPRVMTS")]
    print(f"3) tenant-billed 500/mo posted to 1365000 via code {code} on a LEASE_DEALS line + TI 12,000 on an in-place lease -> IMPRVMTS (expect 1365000 = 12,000 - 6,000 = 6,000):", im, "| status", S())
    # 4) account not on the Financial Format
    coa = sh("DATA_COA"); coa.unprotect(""); rowi = [i + 2 for i, r in enumerate(coa.getCellRangeByPosition(2, 1, 2, 1500).getDataArray()) if r[0] == 5225000][0]
    coa.getCellRangeByName(f"G{rowi}").setString(""); doc.calculateAll()
    rv = sh("LEASE_CALC").getCellRangeByPosition(0, 0, 22, 1000).getDataArray()
    print("4) 5225000 removed from the Financial Format -> LEASE_CALC flag:", {r[21] for r in rv if r[3] == 5225000 and r[21] != ""}, "| status:", S())
    run(); sec = sections(); j = [v for k, v in sec.items() if k.startswith("JANITORIAL")][0]
    print("   rebuilt: DETAIL 5225000 in JANITORIAL:", j.get(5225000), "(expect 108,000: offset ignored) | status:", S())
    doc.close(True)
finally:
    lo.close()

# Block capacity after an MRI data refresh: a lease gains a charge code / a new lease appears.
import sys
from lo_driver import LO
from make_xlsm import embed_vba, cell
lo = LO()
try:
    doc = lo.load(sys.argv[1] if len(sys.argv) > 1 else "base_v2.xlsx"); embed_vba(lo, doc)
    run = lambda: (lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet"), doc.calculateAll())
    S = lambda: cell(doc, "DETAIL", "B3").String
    run(); print("0) built                     ->", S())
    dt = doc.Sheets.getByName("DATA_TENANT_CHARGES"); cc = doc.Sheets.getByName("DATA_CHARGECODES")
    rows = dt.getCellRangeByPosition(0, 1, 11, 399).getDataArray()
    n = sum(1 for r in rows if r[1] != ""); nxt = n + 2          # first empty row (1-based)
    lease_codes = {}
    for r in rows:
        if r[1] != "": lease_codes.setdefault(r[1], set()).add(r[7])
    lease = rows[0][1]
    codes = [r[0] for r in cc.getCellRangeByPosition(0, 1, 3, 456).getDataArray() if r[0] and r[3] and r[0] not in lease_codes[lease]]
    new = codes[0]
    # copy charge row 2 of the first lease, give it an unused (valid) charge code
    for c in range(12):
        src = dt.getCellByPosition(c, 1); dst = dt.getCellByPosition(c, nxt - 1)
        if src.Type.value == "TEXT": dst.setString(src.String)
        elif src.Type.value == "VALUE": dst.setValue(src.Value)
        elif src.Type.value == "FORMULA": dst.setFormula(src.Formula)
    dt.getCellByPosition(7, nxt - 1).setString(new)
    doc.calculateAll()
    print("1) lease", lease, "gains code", new, "->", S())
    # new lease instead
    dt.getCellByPosition(7, nxt - 1).setString(rows[0][7]); dt.getCellByPosition(1, nxt - 1).setString("NEWLS9"); dt.getCellByPosition(2, nxt - 1).setString("9999")
    doc.calculateAll(); print("2) a new lease appears        ->", S())
    doc.close(True)
finally:
    lo.close()

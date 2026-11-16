import datetime as dt, pandas as pd
from lo_driver import LO
from make_xlsm import embed_vba, cell
def ser(d): return (d-dt.date(1899,12,30)).days
lo=LO()
try:
    import sys; doc=lo.load(sys.argv[1] if len(sys.argv)>1 else "base_v2.xlsx"); embed_vba(lo,doc)
    run=lambda: (lo.run_macro(doc,"VBAProject.modBuildAll.BuildAllSchedulesQuiet"), doc.calculateAll())
    S=lambda r: cell(doc,"DETAIL",r).String
    print("tab order:",[doc.Sheets.getByIndex(i).Name for i in range(doc.Sheets.Count)][:6])
    run(); print("1) first build           ->",S("B3"),"| INPUT:",cell(doc,"INPUT","C18").String)
    d=doc.Sheets.getByName("DETAIL"); col=d.getCellRangeByPosition(0,0,1,60).getDataArray()
    print("   'CHECKS' text anywhere on DETAIL/QTRLY:",any("CHECKS" in str(c) for row in col for c in row) or any("CHECKS" in str(c) for row in doc.Sheets.getByName("QTRLY SUM").getCellRangeByPosition(0,0,1,60).getDataArray() for c in row))
    rev=doc.Sheets.getByName("REVENUE"); R=lambda ref: rev.getCellRangeByName(ref)
    # locate block 1 rows from the source tags in column E (layout-independent: blocks are sized to each lease)
    E=[r[0] for r in rev.getCellRangeByPosition(4,0,4,80).getDataArray()]
    A1=E.index("ADD")+1; A2,A3=A1+1,A1+2; C1=E.index("COST")+1
    print("   block 1: ADD rows",A1,A2,A3,"| first COST row",C1)
    # (a) new accounts via additional line + TI -> REBUILD NEEDED, then OK after rebuild
    R(f"B{A1}").setString("07A")
    for c in "GHIJKLMNOPQR": R(f"{c}{A1}").setValue(500)
    R(f"G{C1}").setValue(10000)
    doc.calculateAll(); print("2) new accounts entered    ->",S("B3"))
    run(); print("3) after rebuild           ->",S("B3"))
    # (b) amount on an additional line that has no charge code (no GL account)
    print("   ADD3 code:",repr(R(f"B{A3}").String),"acct:",repr(R(f"D{A3}").String))
    R(f"G{A3}").setValue(777); doc.calculateAll(); print("4) amount on empty ADD line ->",S("B3"))
    run(); print("   rebuilt, still a data problem ->",S("B3"))
    from com.sun.star.sheet.CellFlags import VALUE
    rev.getCellRangeByName(f"G{A3}").clearContents(1)  # VALUE
    doc.calculateAll(); print("   cleared                  ->",S("B3"))
    # (c) unknown charge code on an additional line
    R(f"B{A2}").setString("ZZZ"); R(f"G{A2}").setValue(50); doc.calculateAll()
    print("5) unknown charge code     ->",S("B3"),"| acct:",R(f"D{A2}").String)
    R(f"B{A2}").setString(""); rev.getCellRangeByName(f"G{A2}").clearContents(1); doc.calculateAll()
    print("   fixed                    ->",S("B3"))
    # QTRLY status follows too
    doc.close(True)
finally:
    lo.close()

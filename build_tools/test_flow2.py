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
    ld=doc.Sheets.getByName("LEASE_DEALS"); lda=ld.getCellRangeByPosition(0,0,40,300).getDataArray(); hdr=list(lda[11])
    col=lambda lbl: hdr.index(lbl)
    picks=[j for j,v in enumerate(hdr) if v=="(pick a code below)"]
    d1=[i for i,r in enumerate(lda) if r[0]==1.0][0]                 # deal line 1 (0-based row index)
    yr=int(cell(doc,"INPUT","C3").Value)
    def L(lbl,v,row=d1,j=None):
        c=ld.getCellByPosition(col(lbl) if j is None else j,row); c.setString(v) if isinstance(v,str) else c.setValue(v)
    from com.sun.star.sheet.CellFlags import VALUE, STRING
    print("   deal line 1 = LEASE_DEALS row",d1+1,"| pick-a-code columns",picks)
    # (a) a deal with a charge code whose account is not on DETAIL yet (07A) + TI -> REBUILD NEEDED, then OK after rebuild
    for lbl,v in (("Tenant / description","FLOW TEST"),("Type","New lease"),("Sq Ft",1000),("Lease start",ser(dt.date(yr,1,1))),("Lease end",ser(dt.date(yr+2,12,31))),("TI",10)):
        L(lbl,v)
    ld.getCellByPosition(picks[0],12).setString("07A"); L(None,6.0,j=picks[0])
    doc.calculateAll(); print("2) new accounts entered    ->",S("B3"),"| REVENUE:",cell(doc,"REVENUE","B3").String)
    run(); print("3) after rebuild           ->",S("B3"),"| REVENUE:",cell(doc,"REVENUE","B3").String)
    # (b) amount in a pick-a-code column that has no charge code (no GL account)
    L(None,3.0,j=picks[1]); doc.calculateAll(); print("4) amount with no charge code ->",S("B3"))
    run(); print("   rebuilt, still a data problem ->",S("B3"))
    ld.getCellByPosition(picks[1],d1).clearContents(VALUE); doc.calculateAll(); print("   cleared                  ->",S("B3"))
    # (c) unknown charge code in a pick-a-code column
    ld.getCellByPosition(picks[1],12).setString("ZZZ"); L(None,3.0,j=picks[1]); doc.calculateAll()
    print("5) unknown charge code     ->",S("B3"),"| acct:",ld.getCellByPosition(picks[1],13).String)
    ld.getCellByPosition(picks[1],12).clearContents(STRING); ld.getCellByPosition(picks[1],d1).clearContents(VALUE); doc.calculateAll()
    print("   fixed                    ->",S("B3"))
    # QTRLY status follows too
    doc.close(True)
finally:
    lo.close()

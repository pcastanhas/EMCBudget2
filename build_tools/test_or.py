from lo_driver import LO
from make_xlsm import embed_vba, cell
lo=LO()
try:
    import sys; doc=lo.load(sys.argv[1] if len(sys.argv)>1 else "base_v2.xlsx"); embed_vba(lo,doc)
    run=lambda: (lo.run_macro(doc,"VBAProject.modBuildAll.BuildAllSchedulesQuiet"), doc.calculateAll())
    S=lambda: cell(doc,"DETAIL","B3").String
    orv=doc.Sheets.getByName("OTHER_REVENUE"); O=lambda ref: orv.getCellRangeByName(ref)
    det=doc.Sheets.getByName("DETAIL")
    def drow(acct):
        arr=det.getCellRangeByPosition(0,0,20,60).getDataArray()
        for i,r in enumerate(arr,1):
            if r[0]==acct: return i,r
        return None,None
    run(); base_total=cell(doc,"DETAIL","P12").Value; print("0) base build:",S(),"| TOTAL INCOME",round(base_total))
    cols="GHIJKLMNOPQR"
    # 1) enter interest + misc income
    O("C74").setString("Operating account interest"); O("C75").setString("Money market interest"); O("C84").setString("Vending commissions")
    for c in cols: O(f"{c}74").setValue(3000); O(f"{c}75").setValue(500); O(f"{c}84").setValue(200)
    doc.calculateAll()
    print("1) inputs entered       :",S(),"| block 4430000 2027 total (S77):",round(O("S77").Value),"(expect 42,000) | $/SF",round(O("T77").Value,2),"| misc total (S87):",round(O("S87").Value),"(expect 2,400)")
    print("   summary row 7 (all inputs):",round(O("S7").Value),"(expect 44,400)")
    run(); i,r=drow(4430000); j,m=drow(4440000)
    print("2) after rebuild        :",S(),"| DETAIL 4430000 total",round(r[15]),"| 4440000 total",round(m[15]),"| TOTAL INCOME",round(cell(doc,"DETAIL","P"+str(int(drow(None)[0] or 0) or 12)).Value) if False else "")
    # find TOTAL INCOME row
    arr=det.getCellRangeByPosition(0,0,20,60).getDataArray()
    ti=[(k+1,rw[15]) for k,rw in enumerate(arr) if rw[1]=="TOTAL INCOME"][0]
    print("   TOTAL INCOME now",round(ti[1]),"| increase vs base",round(ti[1]-base_total),"(expect 44,400)")
    print("   accounts in DETAIL income section:",[int(rw[0]) for rw in arr[6:20] if isinstance(rw[0],float)])
    # 3) additive on a tenant account (no silent replacement)
    O("G14").setValue(1000); doc.calculateAll()
    i,r=drow(4100000); print("3) +1,000 typed on base-rent block (tenant account): DETAIL Jan before rebuild",round(r[3]),"(live formula -> includes it)", "| status",S())
    O("G14").setValue(0)
    # 4) amount in an unused block
    O("G124").setValue(500); doc.calculateAll(); print("4) amount in an UNUSED block ->",S())
    O("G124").setValue(0); doc.calculateAll(); print("   cleared                    ->",S())
    # 5) actuals-through change reflows forecast without a rebuild
    ia=cell(doc,"INPUT","C15"); f_before=O("S73").Value
    ia.setValue(6); doc.calculateAll()
    print("5) actuals through Jun   : interest forecast",round(f_before),"->",round(O("S73").Value),"| Jul actual row",O("M70").Value,"| Jul budget row",round(O("M71").Value),"| label:",O("C71").String)
    ia.setValue(9); doc.calculateAll()
    # 6) QTRLY consistent with DETAIL
    q=doc.Sheets.getByName("QTRLY SUM").getCellRangeByPosition(0,0,8,60).getDataArray()
    qrow=[r for r in q if r[0]==4430000][0]; print("6) QTRLY 4430000 total",round(qrow[7]),"| status:",cell(doc,"QTRLY SUM","B3").String)
    doc.close(True)
finally: lo.close()

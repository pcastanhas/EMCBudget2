from lo_driver import LO
from make_xlsm import embed_vba, cell
lo=LO()
try:
    import sys; doc=lo.load(sys.argv[1] if len(sys.argv)>1 else "base_v2.xlsx"); embed_vba(lo,doc)
    run=lambda: (lo.run_macro(doc,"VBAProject.modBuildAll.BuildAllSchedulesQuiet"), doc.calculateAll())
    S=lambda: cell(doc,"DETAIL","B3").String
    fmt=lambda ref: doc.NumberFormats.getByKey(cell(doc,"DETAIL",ref).NumberFormat).FormatString
    w=lambda col: doc.Sheets.getByName("DETAIL").Columns.getByName(col).Width
    inp=cell(doc,"INPUT","C6"); print("INPUT C6 default:",repr(inp.String),"| B6:",cell(doc,"INPUT","B6").String)
    run(); print("1) build, decimals No   :",S(),"| D8 fmt:",fmt("D8"),"| $/SF fmt:",fmt("Q8"),"| width D",w("D"))
    inp.setString("Yes"); doc.calculateAll(); print("2) option set to Yes    :",S())
    run(); print("3) rebuilt, decimals Yes:",S(),"| D8 fmt:",fmt("D8"),"| P12 fmt:",fmt("P12"),"| width D",w("D"))
    inp.setString("No"); doc.calculateAll(); print("4) option back to No    :",S())
    run(); print("5) rebuilt              :",S(),"| D8 fmt:",fmt("D8"))
    doc.close(True)
finally: lo.close()

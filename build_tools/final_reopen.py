import sys
from lo_driver import LO
from make_xlsm import cell
lo=LO()
try:
    doc=lo.load(sys.argv[1])
    libs=doc.BasicLibraries; libs.loadLibrary("VBAProject")
    print("imported modules:",sorted(libs.getByName("VBAProject").getElementNames()))
    lo.run_macro(doc,"VBAProject.modBuildAll.BuildAllSchedulesQuiet"); doc.calculateAll()
    print("rebuild from imported VBA ->",cell(doc,"DETAIL","B3").String,"|",cell(doc,"QTRLY SUM","B3").String,"| BR",round(cell(doc,"DETAIL","P8").Value))
    doc.close(True)
finally: lo.close()

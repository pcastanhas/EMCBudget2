import sys, os, re, uno
from lo_driver import LO
MODS = ["modBuildCommon", "modFormat", "modBuildDetail", "modBuildQtrly", "modBuildAll"]

def module_source(name):
    txt = open(f"/home/claude/build/vba/{name}.bas").read()
    lines = [l for l in txt.splitlines() if not l.startswith("Attribute VB_Name")]
    return "Option VBASupport 1\n" + "\n".join(lines) + "\n"

def embed_vba(lo, doc):
    libs = doc.BasicLibraries
    libs.VBACompatibilityMode = True
    libs.ProjectName = "VBAProject"
    if not libs.hasByName("VBAProject"):
        libs.createLibrary("VBAProject")
    libs.loadLibrary("VBAProject")
    lib = libs.getByName("VBAProject")
    for m in MODS:
        mi = uno.createUnoStruct("com.sun.star.script.ModuleInfo")
        mi.ModuleType = 1
        lib.insertByName(m, module_source(m))
        lib.insertModuleInfo(m, mi)

def cell(doc, sheet, ref):
    return doc.Sheets.getByName(sheet).getCellRangeByName(ref)

if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    lo = LO()
    try:
        doc = lo.load(src)
        embed_vba(lo, doc)
        lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet")
        doc.calculateAll()
        print("DETAIL B3:", cell(doc, "DETAIL", "B3").String, "| B4:", cell(doc, "DETAIL", "B4").String)
        print("QTRLY B3:", cell(doc, "QTRLY SUM", "B3").String)
        lo.save(doc, dst, "Calc MS Excel 2007 VBA XML")
        doc.close(True)
    finally:
        lo.close()
    print("saved", dst)

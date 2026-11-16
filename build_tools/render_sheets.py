"""Render chosen sheet ranges to a PDF for a visual check (LibreOffice headless).
Usage: python3 render_sheets.py book.xlsm out.pdf "COVER:0,0,11,62" "DETAIL:0,0,20,70" [--set "INPUT!C31=Renew" ...] [--build]
  area = SHEET:firstCol,firstRow,lastCol,lastRow (zero-based). Sheets not listed print one empty cell (so pages = sheets).
  --set  sets a cell before rendering (numbers are parsed, everything else is text); --build runs BuildAllSchedulesQuiet first
  (needs a .xlsx base with the VBA embedded by make_xlsm.embed_vba, or an .xlsm that already carries the VBA project)."""
import os, sys, uno
from lo_driver import LO, pv
from com.sun.star.table import CellRangeAddress
args = sys.argv[1:]
book, out = args[0], args[1]
areas, sets, build = {}, [], False
i = 2
while i < len(args):
    a = args[i]
    if a == "--set": sets.append(args[i + 1]); i += 2; continue
    if a == "--build": build = True; i += 1; continue
    name, rng = a.rsplit(":", 1); areas[name] = tuple(int(x) for x in rng.split(",")); i += 1
lo = LO()
try:
    doc = lo.load(book)
    if build:
        from make_xlsm import embed_vba
        if not book.endswith(".xlsm"): embed_vba(lo, doc)
        lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet")
    for s in sets:
        ref, val = s.split("=", 1); sh, cell = ref.split("!")
        c = doc.Sheets.getByName(sh).getCellRangeByName(cell)
        try: c.setValue(float(val))
        except ValueError: c.setString(val)
    doc.calculateAll()
    ps = doc.StyleFamilies.getByName("PageStyles")
    for k in range(doc.Sheets.Count):
        sh = doc.Sheets.getByIndex(k); a = CellRangeAddress(); a.Sheet = k
        if sh.Name in areas: a.StartColumn, a.StartRow, a.EndColumn, a.EndRow = areas[sh.Name]
        else: a.StartColumn = a.StartRow = a.EndColumn = a.EndRow = 0
        sh.setPrintAreas((a,))
        st = ps.getByName(sh.PageStyle); st.IsLandscape = True
        w, h = st.Width, st.Height
        if w < h: st.Width, st.Height = h, w
        st.ScaleToPagesX = 1; st.ScaleToPagesY = 1
    doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(out)), (pv("FilterName", "calc_pdf_Export"),))
    doc.close(True)
finally:
    lo.close()
print("rendered", out)

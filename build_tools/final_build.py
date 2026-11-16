import sys, uno
from lo_driver import LO
from make_xlsm import embed_vba, cell
from com.sun.star.awt import Size, Point

def add_button(doc, sheet_name="INPUT"):
    sh = doc.Sheets.getByName(sheet_name)
    a = sh.getCellRangeByName("B20"); b = sh.getCellRangeByName("C21")
    x, y = a.Position.X, a.Position.Y
    w = b.Position.X + b.Size.Width - x
    h = b.Position.Y + b.Size.Height - y
    shape = doc.createInstance("com.sun.star.drawing.RectangleShape")
    shape.Size = Size(w, h); shape.Position = Point(x, y)
    sh.DrawPage.add(shape)
    shape.Name = "BuildSchedulesButton"
    shape.String = "Build Schedules"
    shape.FillColor = 0xFFC000; shape.LineColor = 0x000000
    shape.CharFontName = "Arial"; shape.CharHeight = 11; shape.CharWeight = 150; shape.CharColor = 0x000000

def col_letter(n):
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def export_font_colors(doc, sheet_name, rng, path):
    """LibreOffice 24.2 exports a stale theme colour instead of a macro-set font colour.
    Record every explicitly coloured cell (live document = what the VBA wrote) so that
    postprocess.py can write those colours into the file's styles."""
    import json
    sh = doc.Sheets.getByName(sheet_name)
    r = sh.getCellRangeByName(rng).RangeAddress
    out = {}
    for row in range(r.StartRow, r.EndRow + 1):
        for col in range(r.StartColumn, r.EndColumn + 1):
            cc = sh.getCellByPosition(col, row).CharColor
            if cc not in (-1, 0):
                out[f"{col_letter(col)}{row + 1}"] = "%06X" % cc
    json.dump(out, open(path, "w"))
    return len(out)

if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    lo = LO()
    try:
        doc = lo.load(src)
        embed_vba(lo, doc)
        lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet")
        doc.calculateAll()
        last = 1
        arr = doc.Sheets.getByName("DETAIL").getCellRangeByPosition(0, 0, 0, 200).getDataArray()
        for i, row in enumerate(arr, 1):
            if row[0] not in ("", None): last = i
        print("coloured DETAIL cells exported:", export_font_colors(doc, "DETAIL", f"A1:U{last + 2}", "/home/claude/build/colors_detail.json"))
        add_button(doc)
        print("status:", cell(doc, "DETAIL", "B3").String)
        lo.save(doc, dst, "Calc MS Excel 2007 VBA XML")
        doc.close(True)
    finally:
        lo.close()
    print("saved", dst)

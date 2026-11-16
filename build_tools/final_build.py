import sys, uno
from paths import COLORS_JSON
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

def font_colors(doc, sheet_name, rng):
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
    return out


def last_used_row(doc, sheet_name, rows=400):
    arr = doc.Sheets.getByName(sheet_name).getCellRangeByPosition(0, 0, 0, rows).getDataArray()
    return max([i for i, row in enumerate(arr, 1) if row[0] not in ("", None)] + [1])


def export_font_colors(doc, path):
    """{sheet: {cell: rgb}} for the macro-built schedules that use coloured text (postprocess.py writes them into the styles)."""
    import json
    out = {"REVENUE": font_colors(doc, "REVENUE", f"A1:Q{last_used_row(doc, 'REVENUE', 2000) + 2}"),
           "DETAIL": font_colors(doc, "DETAIL", f"A1:U{last_used_row(doc, 'DETAIL') + 2}"),
           "PROFORMA": font_colors(doc, "PROFORMA", f"A1:J{last_used_row(doc, 'PROFORMA') + 2}")}
    json.dump(out, open(path, "w"))
    return {k: len(v) for k, v in out.items()}

if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    colors_out = sys.argv[3] if len(sys.argv) > 3 else COLORS_JSON   # per-build file; default = build_tools/colors_detail.json
    lo = LO()
    try:
        doc = lo.load(src)
        embed_vba(lo, doc)
        lo.run_macro(doc, "VBAProject.modBuildAll.BuildAllSchedulesQuiet")
        doc.calculateAll()
        print("coloured cells exported:", export_font_colors(doc, colors_out))
        add_button(doc)
        print("status:", cell(doc, "DETAIL", "B3").String)
        lo.save(doc, dst, "Calc MS Excel 2007 VBA XML")
        doc.close(True)
    finally:
        lo.close()
    print("saved", dst)
